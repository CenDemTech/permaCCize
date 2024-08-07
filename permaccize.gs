const kErrorCode400 = "errorCode400"
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
    var permalinkError = false
    // only replace the link if it's not already a permalink
    if (!link.url.includes("perma.cc")) {
      // if the link has already been permalinked on this run of the script (because it shows up twice), just use that
      if (link.url in permalinks) {
        var permalink = permalinks[link.url];
      }

      // otherwise make a permalink...
      else {
        var permalink = makePermalink(link.url, api_key);
        if (permalink != kErrorCode400) {
          // ...and add it to the permalinks object in case it shows up again in the doc
          permalinks[link.url] = permalink;
        } else {
          permalinkError = true
        }
      }

      // find the start and end location for the text we are about to edit
      var start = link.startOffset;
      var end = link.endOffsetInclusive; // TODO: this needs another +1 when the element ends with a linebreak. regardless of whether it's a URL or not. tricky!

      if (permalinkError == true) {
        link.element.setForegroundColor(start, end+1, '#FF0000')
        continue
      }

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

function appendLinksToText(paragraphs, api_key, bluebook) {
  var links = [];
  var permalinks = {};

  for (var p = 0; p < paragraphs.length; p++) {
    paragraph = paragraphs[p];
    paragraphLinks = getAllLinks(paragraph);

    var permalinkError = false

    // only act on paragraphs that have a single link
    if (paragraphLinks.length == 1) {
      link = paragraphLinks[0];
      // only replace the link if it's not already a permalink
      if (!link.url.includes("perma.cc")) {
        // if the link has already been permalinked on this run of the script (because it shows up twice), just use that
        if (link.url in permalinks) {
          var permalink = permalinks[link.url];
        }

        // otherwise make a permalink
        else {
          // ...and add it to the permalinks object in case it shows up again in the doc
          var permalink = makePermalink(link.url, api_key);
          if (permalink != kErrorCode400) {
            permalinks[link.url] = permalink;
          } else {
            permalinkError = true
          }
        }

        // if there's an error, make the text red and keep going.
        if (permalinkError == true) {
          link.element.setForegroundColor(link.startOffset, link.endOffsetInclusive+1, '#FF0000')
          continue
        }

        // Otherwise, append the link to the footnote

        // eliminate trailing spaces
        while (paragraph.getText().endsWith(" ")) {
          paragraph
            .editAsText()
            .deleteText(
              paragraph.getText().length - 1,
              paragraph.getText().length - 1
            );
        }

        // eliminate trailing period(s) (Bluebook only)
        if (bluebook) {
          while (paragraph.getText().endsWith(".")) {
            paragraph
              .editAsText()
              .deleteText(
                paragraph.getText().length - 1,
                paragraph.getText().length - 1
              );
          }
        }

        // add permalink in brackets
        var oldLength = paragraph.getText().length;
        if (bluebook) {
          paragraph.appendText(` [${permalink}].`);
        } else {
          paragraph.appendText(` [${permalink.replace("https://", "")}]`);
        }
        var newLength = paragraph.getText().length;

        if (bluebook) {
          linkUrlOffset = 3;
        } else {
          linkUrlOffset = 2;
        }

        // format properly
        paragraph.editAsText().setItalic(oldLength, newLength - 1, false);
        paragraph.editAsText().setLinkUrl(oldLength, newLength - 1, "");
        paragraph
          .editAsText()
          .setLinkUrl(oldLength + 2, newLength - linkUrlOffset, permalink);
      }
    }
  }
}

function appendFootnoteLinks(bluebook = false) {
  // append perma.cc links to paragraphs in footnotes that have only one link
  api_key = promptForKey();

  var doc = DocumentApp.getActiveDocument();

  var footnotes = doc.getFootnotes();
  var paragraphs = []
  for (var f = 0; f < footnotes.length; f++) {
    paragraphs = paragraphs.concat(footnotes[f].getFootnoteContents().getParagraphs());
  }
  appendLinksToText(paragraphs, api_key, bluebook)

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
    throw "Script cannot be run without an API key.";
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
    muteHttpExceptions: true,
    payload: JSON.stringify(data),
  };

  var response = UrlFetchApp.fetch(request_url, options)

  var errorCode = response.getResponseCode()
  if(errorCode >= 400 && errorCode < 500) {
    return kErrorCode400
  }

  var responseJSON = JSON.parse(response.getContentText());
  Logger.log(responseJSON)
  var permalink = "https://perma.cc/".concat(responseJSON["guid"]);
  Logger.log(permalink)
  return permalink;
}

function makeFakePermalink(url, api_key) {
  //  random response (for debugging, without making a real link)
  var permalink = "https://perma.cc/"
    .concat(url)
    .concat(String(Math.round(Math.random() * 1000)).padStart(3, "0"));

  return permalink;
}

function appendFootnoteLinksBluebook() {
  appendFootnoteLinks((bluebook = true));
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

function appendBibliographyLinks() {
  var api_key = promptForKey();

  // Get the active document and its body
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var links = [];
  var permalinks = {};
  
  // Get all paragraphs in the document
  var paragraphs = body.getParagraphs();
  appendLinksToText(paragraphs, api_key, false)
}

function onOpen() {
  // Add a menu including this link
  DocumentApp.getUi()
    .createMenu("Perma.cc")
    .addItem("Replace all links with Perma.cc links", "replaceAllLinks")
    .addItem(
      "Append all links with bracketed Perma.cc links (bibliography)",
      "appendBibliographyLinks"
      )
    .addItem(
      "Append footnote links with bracketed Perma.cc links",
      "appendFootnoteLinks"
    )
    .addItem(
      "Append footnote links with bracketed Perma.cc links (Bluebook)",
      "appendFootnoteLinksBluebook"
    )
    .addToUi();
}