function replaceAllLinks() {
  // replaces all links in the document with perma.cc links

  api_key = promptForKey();

  var ui = DocumentApp.getUi();

  var doc = DocumentApp.getActiveDocument();

  // initialize array for all of the...
  var links = [];

  var selection = doc.getSelection();
  if (selection) {
    // if a selection is made
    ui.alert(
      "Detected a selection. Will convert only the links in the selection."
    );

    // ...links in the selection
    var selectionElements = selection.getRangeElements();
    for (var e = 0; e < selectionElements.length; e++) {
      links.push(...getAllLinks(selectionElements[e].getElement()));
    }
  } else {
    // if no selection is made
    ui.alert("Converting all links in the document.");

    // ...link elements in the body...
    links.push(...getAllLinks(doc.getBody()));

    // ... and in the header and footer...
    links.push(...getAllLinks(doc.getHeader()));
    links.push(...getAllLinks(doc.getFooter()));

    // ...and in the footnotes.
    var footnotes = doc.getFootnotes();
    for (var f = 0; f < footnotes.length; f++) {
      links.push(...getAllLinks(footnotes[f].getFootnoteContents()));
    }
  }

  // initialize object for all the links we're about to make
  var permalinks = {};

  // loop over links
  for (var l = 0; l < links.length; l++) {
    var link = links[l];

    // only replace the link if it's not already a permalink
    if (!link.url.includes("perma.cc")) {
      // if the link has already been permalinked on this run of the script (because it shows up twice), just use that
      if (link.url in permalinks) {
        var permalink = permalinks[link.url];
      }

      // otherwise make a permalink...
      else {
        var permalink = makePermalink(link.url, api_key);
        // ...and add it to the permalinks object in case it shows up again in the doc
        permalinks[link.url] = permalink;
      }

      // find the start and end location for the text we are about to edit
      var start = link.startOffset;
      var end = link.endOffsetInclusive; // TODO: this needs another +1 when the element ends with a linebreak. regardless of whether it's a URL or not. tricky!

      // get the text, as displayed
      var urlText = link.element.getText().slice(start, end + 1);

      if (isUrl(urlText)) {
        // if the displayed text appears to be a URL (ie it starts with 'http'), replace the displayed text and the link URL with the permalink
        link.element.deleteText(start, end);
        link.element.insertText(start, permalink);
        link.element.setLinkUrl(start, start + permalink.length - 1, permalink);
      } else {
        // if the displayed text is different than the URL, only replace the link URL with the permalink
        link.element.setLinkUrl(start, start + urlText.length - 1, permalink);
      }
    }
  }
}

function appendFootnoteLinks() {
  // append perma.cc links to paragraphs in footnotes that have only one link

  api_key = promptForKey();

  var doc = DocumentApp.getActiveDocument();

  // initialize arrays for all of the links in the footnotes, and all the permalinks we are about to make
  var links = [];
  var permalinks = {};

  var footnotes = doc.getFootnotes();
  for (var f = 0; f < footnotes.length; f++) {
    paragraphs = footnotes[f].getFootnoteContents().getParagraphs();

    // within each footnote, loop over the paragraphs
    for (var p = 0; p < paragraphs.length; p++) {
      paragraph = paragraphs[p];
      paragraphLinks = getAllLinks(paragraph);

      // only act on paragraphs that have a single link
      if (paragraphLinks.length == 1) {
        link = paragraphLinks[0];
        // only replace the link if it's not already a permalink
        if (!link.url.includes("perma.cc")) {
          // if the link has already been permalinked on this run of the script (because it shows up twice), just use that
          if (link.url in permalinks) {
            var permalink = permalinks[link.url];
          }

          // otherwise make a permalink...
          else {
            var permalink = makePermalink(link.url, api_key);
            // ...and add it to the permalinks object in case it shows up again in the doc
            permalinks[link.url] = permalink;
          }

          // ... and append the link to the footnote

          // eliminate trailing spaces
          while (paragraph.getText().endsWith(" ")) {
            paragraph
              .editAsText()
              .deleteText(
                paragraph.getText().length - 1,
                paragraph.getText().length - 1
              );
          }

          // add permalink in brackets
          var oldLength = paragraph.getText().length;
          paragraph.appendText(` [${permalink}]`);
          var newLength = paragraph.getText().length;

          // format properly
          paragraph.editAsText().setItalic(oldLength, newLength - 1, false);
          paragraph.editAsText().setLinkUrl(oldLength, newLength - 1, "");
          paragraph
            .editAsText()
            .setLinkUrl(oldLength + 2, newLength - 2, permalink);
        }
      }
    }
  }
}

function promptForKey() {
  // ask for the perma.cc API key
  var ui = DocumentApp.getUi();

  var result = ui.prompt(
    "Please enter your Perma.cc API key:",
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var api_key = result.getResponseText();
  if (
    button == ui.Button.CANCEL ||
    button == ui.Button.CLOSE ||
    api_key.length == 0
  ) {
    ui.alert("I can't run without an API key.");
    return;
  }
  return api_key;
}

function isUrl(string) {
  return string.startsWith("http");
}

function makePermalink(url, api_key) {
  // makes archive request to perma.cc and returns perma.cc URL

  var request_url = "https://api.perma.cc/v1/archives/?api_key=".concat(
    api_key
  );
  var data = {
    url: url,
    folder: 137682,
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(data),
  };

  var response = JSON.parse(
    UrlFetchApp.fetch(request_url, options).getContentText()
  );
  var permalink = "https://perma.cc/".concat(response["guid"]);

  return permalink;
}

function makeFakePermalink(url, api_key) {
  //  random response (for debugging, without making a real link)
  var permalink = "http://perma.cc/"
    .concat(url)
    .concat(String(Math.round(Math.random() * 1000)).padStart(3, "0"));

  return permalink;
}

/**
 * by @mogsdad: https://stackoverflow.com/questions/18727341/get-all-links-in-a-document
 * Get an array of all LinkUrls in the document. The function is
 * recursive, and if no element is provided, it will default to
 * the active document's Body element.
 *
 * @param {Element} element The document element to operate on.
 * .
 * @returns {Array}         Array of objects, vis
 *                              {element,
 *                               startOffset,
 *                               endOffsetInclusive,
 *                               url}
 */
function getAllLinks(element) {
  var links = [];

  if (
    element !== null &&
    element.hasOwnProperty("getType") &&
    element.getType() === DocumentApp.ElementType.TEXT
  ) {
    var textObj = element.editAsText();
    var text = element.getText();
    var inUrl = false;
    for (var ch = 0; ch < text.length; ch++) {
      var url = textObj.getLinkUrl(ch);
      if (url != null && ch != text.length - 1) {
        if (!inUrl) {
          // We are now!
          inUrl = true;
          var curUrl = {};
          curUrl.element = element;
          curUrl.url = String(url); // grab a copy
          curUrl.startOffset = ch;
        } else {
          curUrl.endOffsetInclusive = ch;
        }
      } else {
        if (inUrl) {
          // Not any more, we're not.
          inUrl = false;
          links.push(curUrl); // add to links
          curUrl = {};
        }
      }
    }
  } else {
    // Get number of child elements, for elements that can have child elements.
    try {
      var numChildren = element.getNumChildren();
    } catch (e) {
      numChildren = 0;
    }
    for (var i = 0; i < numChildren; i++) {
      links = links.concat(getAllLinks(element.getChild(i)));
    }
  }

  return links;
}

function onOpen() {
  // Add a menu including this link
  DocumentApp.getUi()
    .createMenu("Utils")
    .addItem("Replace all links with Perma.cc", "replaceAllLinks")
    .addItem("Append footnote links with Perma.cc", "appendFootnoteLinks")
    .addToUi();
}
