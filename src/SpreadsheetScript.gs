function doGet(e) {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = doc.getSheetByName("Sheet1");

  const action = e.parameter.action;
  if (action ==="getAllReviews") {
    const range = sheet.getRange(2,1,sheet.getLastRow()-1,7);
    const values = range.getDisplayValues();

    const result = values.map((h, index) => ({
      title: h[0],
      artist: h[1],
      author: h[2],
      score: h[3],
      url: h[4],
      image: h[5],
      quote: h[6]
      // text: RICHTEXT_TO_HTML(sheet.getRange(2+index,8))
    }));

    return ContentService.createTextOutput(JSON.stringify({data:result})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "getFeaturedReview") {
    const range = sheet.getRange("J2");
    const result = range.getDisplayValue();
    return ContentService.createTextOutput(JSON.stringify({data:result})).setMimeType(ContentService.MimeType.JSON);
  }
    
  else if (action === "getReview") {
    const index = parseInt(e.parameter.index);
    const range = sheet.getRange(index+2,1,index+2,9);
    const values = range.getDisplayValues();

    const result = values.map((h) => ({
      title: h[0],
      artist: h[1],
      author: h[2],
      score: h[3],
      url: h[4],
      image: h[5],
      quote: h[6],
      text: h[7],
      comments: h[8]
    }));

    return ContentService.createTextOutput(JSON.stringify({data:result})).setMimeType(ContentService.MimeType.JSON);
  }

  else if (action === "postComment") {
    try {
      // Extract parameters from the request
      const index = parseInt(e.parameter.index);
      const name = e.parameter.name;
      const text = e.parameter.text;
      const range = sheet.getRange(index + 2, 9);
      
      // Get the original value
      const original = range.getDisplayValue();

      const date = new Date().toISOString();
      
      // Construct the new comment
      let newComment = `{"commenter": "${name}", "text": "${text}", "date": "${date}"}`;
      if (original !== "") {
        newComment = '\n' + newComment;
      }

      // Set the new value in the range
      range.setValue(original + newComment);

      // Construct the result object
      const result = newComment;

      // Return the result as JSON
      return ContentService.createTextOutput(JSON.stringify({data: result}))
                          .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      // Handle any errors that occur
      const result = {
        status: "error",
        message: error.message
      };

      // Return the error message as JSON
      return ContentService.createTextOutput(JSON.stringify({data: result})).setMimeType(ContentService.MimeType.JSON);
    }
  }
}