/*------------------------------------------------
Google App Script to Parse Google Doc Content using Google Cloud Natural Language API
and store results in a new Google Sheet titled "NLP_Results_timestamp" in a tab titled "entitySentiment". 
Also Parses Google Doc Content for URLs and stores those URLs in a separate tab titled "Links" on the
same Google Sheet.
To run this script you will need a Google Cloud Natural Language API Key.
------------------------------------------------*/

// The function will take the document body and send it to the Google Cloud Natural Language API
// Documentation: https://cloud.google.com/natural-language/docs/reference/rest
function retrieveEntitySentiment(line) {
    var doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const docText = body.getText();
    var apiKey = "XXXXXXXXXXXXXXXXXXX"; // Must provide your own API key

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
        var row = [link];
        newValues1.push(row);
    });
    if (newValues1.length) {
        linksSheet.getRange(linksSheet.getLastRow() + 1, 1, newValues1.length,
            newValues1[0].length).setValues(newValues1);
    }

}

// Returns a flat array of links which appear in the active document's body. 
// https://urlregex.com/
function getAllLinks() {
    var doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const docText = body.getText();
    var urlRegex = /((([A-Za-z]{3,9}:(?:\/\/)?)(?:[\-;:&=\+\$,\w]+@)?[A-Za-z0-9\.\-]+|(?:www\.|[\-;:&=\+\$,\w]+@)[A-Za-z0-9\.\-]+)((?:\/[\+~%\/\.\w\-_]*)?\??(?:[\-\+=&;%@\.\w_]*)#?(?:[\.\!\/\\\w]*))?)/g
    //console.log(docText.match(urlRegex))
    return docText.match(urlRegex);
}
