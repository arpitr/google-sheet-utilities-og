function sendConditionalEmails() {
  // Open the active spreadsheet and get the sheet you want to work with
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const notifySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notify-email");

  // Get the data range for all FDs
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // Get the data range for all Email Recipients
  const notifydataRange = notifySheet.getDataRange();
  const notifydata = notifydataRange.getValues();

  // Set other variables and consts
  const currentDate = new Date();
  let dayDifference = 0;
  let matureditemsList = [];
  let unmatureditemsList = [];
  let maturedSummary = '';
  let unmaturedSummary = '';
  let totalInvested = 0;
  let totalMaturity = 0;
  let totalInterest = 0;

  // Loop through each row, starting from the second row (assuming first row is headers)
  for (let i = 1; i < data.length; i++) {
    
    totalInvested += data[i][5];
    totalMaturity += data[i][6];
    totalInterest += data[i][11]; 
    const dateCell = data[i][8]; // Assume condition is in the 8th column (B)
    
    if (dateCell instanceof Date) { // Check if the cell contains a date
      const timeDifference = dateCell - currentDate; // Difference in milliseconds
      dayDifference = Math.floor(timeDifference / (1000 * 60 * 60 * 24)); // Convert to days

      // Log or write the difference in days back to the sheet
      Logger.log("Row " + i + ": " + dayDifference + " days");
      sheet.getRange(i + 1, 10).setValue(dayDifference); // Write to next column
    } else {
      Logger.log("Row " + i + ": No date found in this cell.");
    }
    // Check if the condition is met
    if (dayDifference <= 100) {
      // Build list of all matured items .
      matureditemsList.push(i); 
    } else {
      unmatureditemsList.push(i); 
      Logger.log("Policy ID " + data[i][4] + " of " + data[i][1] + " still has " + dayDifference + " days left for Maturity.");
    }
  }

  // Build mail content 
  for(let j=0; j < matureditemsList.length; j++) {
    let a = Number(matureditemsList[j]);
    maturedSummary += `\n\n ${j+1}. Policy ID ${data[a][4]} in the name of ${data[a][1]} with Bank ${data[a][3]} on Interest Rate ${data[a][10]} is about to mature in ${data[a][9]} Days `;
  } 

  for(let k=0; k < unmatureditemsList.length; k++) {
    let b = Number(unmatureditemsList[k]);
    unmaturedSummary += `\n\n ${k+1}. Policy ID ${data[b][4]} in the name of ${data[b][1]} with Bank ${data[b][3]} on Interest Rate ${data[b][10]} is about to mature in ${data[b][9]} Days `;
  } 
  Logger.log(maturedSummary);
  Logger.log(unmaturedSummary);
  // Send email 
  for (let i = 1; i < notifydata.length; i++) {
    const name = notifydata[i][0];
    const recipientEmail = notifydata[i][1];
    const subject = "FD Investment Maturity Summary";
    const body = `Hello ${name},\n\n` 
                  + `******************* Matured FD List ******************* ${maturedSummary} \n\n`
                  + `****************** Remaining FD  List ****************** ${unmaturedSummary}\n\n`
                  + `************************ Summary ************************ \n\n`
                  + `Total Investment = ${totalInvested}\n\n`
                  + `Total Maturity Amount = ${totalMaturity}\n\n`
                  + `Total Interest Received = ${totalInterest}\n\n`;
    // Send the email
    GmailApp.sendEmail(recipientEmail, subject, body);
  }
}
