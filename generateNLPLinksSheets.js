/*------------------------------------------------
Google App Script to Parse Google Doc Content using Google Cloud Natural Language API
and store results in a new Google Sheet titled "NLP_Results_timestamp" in a tab titled "entitySentiment". 
Also Parses Google Doc Content for URLs and stores those URLs in a separate tab titled "Links" on the
same Google Sheet.
To run this script you will need a Google Cloud Natural Language API Key.
------------------------------------------------*/

// The function will take the document body and send it to the Google Cloud Natural Language API
// Documentation: https://cloud.google.com/natural-language/docs/reference/rest
function Main(line) {
    var doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const docText = body.getText();
    var apiKey = "XXXXXXXXXXXXXXXXXXXXX"; // Must provide your own API key

    var apiEndpoint = 'https://language.googleapis.com/v1/documents:analyzeEntitySentiment?key=' + apiKey;

    var nlData = {
        document: {
            language: 'en-us',
            type: 'PLAIN_TEXT',
            content: docText
        },
        encodingType: 'UTF8'
    };
    //  Package all of the options and the data together for the call
    var nlOptions = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(nlData)
    };
    //  And make the call
    var response = UrlFetchApp.fetch(apiEndpoint, nlOptions);
    var data = JSON.parse(response);

    // Send the results to the createSpreadsheet function
    createSpreadsheet(data)
    return;
};

// The function will create a spreadsheet and store the NLP results in one tab and all the document links in another tab.
// Documentation: https://cloud.google.com/blog/products/ai-machine-learning/analyzing-text-in-a-google-sheet-using-cloud-natural-language-api-and-apps-script
function createSpreadsheet(nlData) {
    // We want to create a unique name for the spreadsheet by using a timestamp.
    var filename = "NLP_Results_" + Math.floor(Date.now() / 1000);
    var ss = SpreadsheetApp.create(filename);

    // Checks to see if entitySentiment sheet is present; if not, creates new sheet and sets header row
    var entitySheet = ss.getSheetByName('entitySentiment');
    if (entitySheet == null) {
        ss.insertSheet('entitySentiment');
        var entitySheet = ss.getSheetByName('entitySentiment');
        var esHeaderRange = entitySheet.getRange(1, 1, 1, 6);
        var esHeader = [['Entity', 'Entity Type', 'Salience', 'Sentiment Score', 'Sentiment Magnitude', 'Number of mentions']];
        esHeaderRange.setValues(esHeader);
    };
    // Looping through the json results and writing rows to the spreadsheet
    var newValues = [];
    var entities = nlData.entities;
    entities.forEach(function (entity) {
        var row = [entity.name, entity.type, entity.salience, entity.sentiment.score,
        entity.sentiment.magnitude, entity.mentions.length,
        ];
        newValues.push(row);
    });

    if (newValues.length) {
        entitySheet.getRange(entitySheet.getLastRow() + 1, 1, newValues.length,
            newValues[0].length).setValues(newValues);
    }

    // Create a new tab called "Links"
    var linksSheet = ss.getSheetByName('Links');
    if (linksSheet == null) {
        ss.insertSheet('Links');
        var linksSheet = ss.getSheetByName('Links');
        var esHeaderRange = linksSheet.getRange(1, 1, 1, 1);
        var esHeader = [['URL']];
        esHeaderRange.setValues(esHeader);
    };

    var newValues1 = [];
    // Call the function getAlllinks which returns an array of parsed links from the document
    var links = getAllLinks();
    // Loop through the array and write to rows on the spreadsheet
    links.forEach(function (link) {
        var row = [link.url];
        newValues1.push(row);
    });
    if (newValues1.length) {
        linksSheet.getRange(linksSheet.getLastRow() + 1, 1, newValues1.length,
            newValues1[0].length).setValues(newValues1);
    }

}


/**
 * Pasted from: https://stackoverflow.com/questions/18727341/get-all-links-in-a-document
 * Returns a flat array of links which appear in the active document's body. 
 * Each link is represented by a simple Javascript object with the following 
 * keys:
 *   - "section": {ContainerElement} the document section in which the link is
 *     found. 
 *   - "isFirstPageSection": {Boolean} whether the given section is a first-page
 *     header/footer section.
 *   - "paragraph": {ContainerElement} contains a reference to the Paragraph 
 *     or ListItem element in which the link is found.
 *   - "text": the Text element in which the link is found.
 *   - "startOffset": {Number} the position (offset) in the link text begins.
 *   - "endOffsetInclusive": the position of the last character of the link
 *      text, or null if the link extends to the end of the text element.
 *   - "url": the URL of the link.
 *
 * @param {boolean} mergeAdjacent Whether consecutive links which carry 
 *     different attributes (for any reason) should be returned as a single 
 *     entry.
 * 
 * @returns {Array} the aforementioned flat array of links.
 */
function getAllLinks(mergeAdjacent) {
  var links = [];

  var doc = DocumentApp.getActiveDocument();


  iterateSections(doc, function(section, sectionIndex, isFirstPageSection) {
    if (!("getParagraphs" in section)) {
      // as we're using some undocumented API, adding this to avoid cryptic
      // messages upon possible API changes.
      throw new Error("An API change has caused this script to stop " + 
                      "working.\n" +
                      "Section #" + sectionIndex + " of type " + 
                      section.getType() + " has no .getParagraphs() method. " +
        "Stopping script.");
    }

    section.getParagraphs().forEach(function(par) { 
      // skip empty paragraphs
      if (par.getNumChildren() == 0) {
        return;
      }

      // go over all text elements in paragraph / list-item
      for (var el=par.getChild(0); el!=null; el=el.getNextSibling()) {
        if (el.getType() != DocumentApp.ElementType.TEXT) {
          continue;
        }

        // go over all styling segments in text element
        var attributeIndices = el.getTextAttributeIndices();
        var lastLink = null;
        attributeIndices.forEach(function(startOffset, i, attributeIndices) { 
          var url = el.getLinkUrl(startOffset);

          if (url != null) {
            // we hit a link
            var endOffsetInclusive = (i+1 < attributeIndices.length? 
                                      attributeIndices[i+1]-1 : null);

            // check if this and the last found link are continuous
            if (mergeAdjacent && lastLink != null && lastLink.url == url && 
                  lastLink.endOffsetInclusive == startOffset - 1) {
              // this and the previous style segment are continuous
              lastLink.endOffsetInclusive = endOffsetInclusive;
              return;
            }

            lastLink = {
              "section": section,
              "isFirstPageSection": isFirstPageSection,
              "paragraph": par,
              "textEl": el,
              "startOffset": startOffset,
              "endOffsetInclusive": endOffsetInclusive,
              "url": url
            };
console.log(lastLink.url)
            links.push(lastLink);
          }        
        });
      }
    });
  });


  return links;
}

/**
 * Calls the given function for each section of the document (body, header, 
 * etc.). Sections are children of the DocumentElement object.
 *
 * @param {Document} doc The Document object (such as the one obtained via
 *     a call to DocumentApp.getActiveDocument()) with the sections to iterate
 *     over.
 * @param {Function} func A callback function which will be called, for each
 *     section, with the following arguments (in order):
 *       - {ContainerElement} section - the section element
 *       - {Number} sectionIndex - the child index of the section, such that
 *         doc.getBody().getParent().getChild(sectionIndex) == section.
 *       - {Boolean} isFirstPageSection - whether the section is a first-page
 *         header/footer section.
 */
function iterateSections(doc, func) {
  // get the DocumentElement interface to iterate over all sections
  // this bit is undocumented API
  var docEl = doc.getBody().getParent();

  var regularHeaderSectionIndex = (doc.getHeader() == null? -1 : 
                                   docEl.getChildIndex(doc.getHeader()));
  var regularFooterSectionIndex = (doc.getFooter() == null? -1 : 
                                   docEl.getChildIndex(doc.getFooter()));

  for (var i=0; i<docEl.getNumChildren(); ++i) {
    var section = docEl.getChild(i);

    var sectionType = section.getType();
    var uniqueSectionName;
    var isFirstPageSection = (
      i != regularHeaderSectionIndex &&
      i != regularFooterSectionIndex && 
      (sectionType == DocumentApp.ElementType.HEADER_SECTION ||
       sectionType == DocumentApp.ElementType.FOOTER_SECTION));

    func(section, i, isFirstPageSection);
  }
}
