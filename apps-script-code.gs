function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    const email = data.email;
    
    // Get all values in the email column (column C)
    const range = sheet.getRange('C:C');
    const values = range.getValues();
    
    // Flatten and check if email already exists
    const emailList = values.flat();
    if (emailList.includes(email)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'This email is already on the waitlist'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Email is unique, add to sheet
    sheet.appendRow([new Date(), data.name, email]);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Entry added to waitlist'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
