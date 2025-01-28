function generateReport() {
  defineCell();
  // Get the active spreadsheet and the first sheet
  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheet = createNewSheet();
  buildHeader(sheet);

  //Year
  purchaseYear = getCellValue(this.purchaseYearCell[0],this.purchaseYearCell[1]);
  sellYear = getCellValue(this.sellYearCell[0],this.sellYearCell[1]);
  rentTillYear = sellYear-1;

  //Rental Income
  const cost= getCellValue(this.costRealEstate[0], this.costRealEstate[1]);
  const rentalYeild = getCellValue(this.rentalYeild[0], this.rentalYeild[1]);

  // Loan Calculations
  const loanAmount = getCellValue(this.loanAmount[0], this.loanAmount[1]);
  const loanInterestRate = getCellValue(this.loanInterestRate[0], this.loanInterestRate[1]);
  const emiAmount = getCellValue(this.emiAmount[0], this.emiAmount[1]);
  let principal = loanAmount;
  let rentalIncomeSD = 0;
  let additionTaxSaving=0;
  let yearlyEMI = 0;
  
  // Rent Tax Calculation
  const sdOnRent = getCellValue(this.sdOnRent[0], this.sdOnRent[1]);
  const taxOnRentPerc = getCellValue(this.taxOnRentPerc[0], this.taxOnRentPerc[1]);
  let netRentForTax = 0;
  let netTaxOnRent = 0;

  //Additional One time charges
  let additionalCost = 0;
  const registeryCostPerc = getCellValue(this.registeryCostPerc[0], this.registeryCostPerc[1]);
  const upfrontPayment = getCellValue(this.upfrontPayment[0], this.upfrontPayment[1]);
  const realEstateGrowthPerc = getCellValue(this.realEstateGrowthPerc[0], this.realEstateGrowthPerc[1]);
  const loanTenure = getCellValue(this.loanTenure[0], this.loanTenure[1]);
  
  // Net Flow
  let yearlynetFlow = 0;
  let monthlynetFlow = 0;
  let cycles=0;
  let fv=0;
  let xirrDates = [];
  let xirr = 0;
  let lastPrincipal = 0;
  let rowsAdded = [];
  let cashFlowsYearly = [];
  let netFlows = [];
  
  for(let i=purchaseYear; i <= rentTillYear; i++) {
   let xirrCashFlows =[];
   const date = "01/01/" + i;
   xirrDates.push(date);
   const yearlyLoanIntandPrin = calculateYearlyLoanIntandPrin(emiAmount, principal,loanInterestRate);
   if(principal) {
      yearlyEMI = emiAmount*12;
    } else {
      yearlyEMI = 0;
    } 
   lastPrincipal = principal;
   principal = Math.max(principal- yearlyLoanIntandPrin[0],0);
   if(i == purchaseYear) {
    rentalIncome = cost*rentalYeild;
    additionalCost = additionalCostCalculate(registeryCostPerc,cost,upfrontPayment);
   } else {
    rentalIncome = rentalIncome*1.1;
    additionalCost = 0;
    fv= calculateFV(cost, realEstateGrowthPerc,1, cycles,0)-lastPrincipal;
    xirrCashFlows = Array.from(cashFlowsYearly);
    xirrCashFlows.push(fv);
    xirr = calculateXirr(xirrCashFlows,xirrDates);
   }
    rentalIncomeSD = rentalIncome*(1-sdOnRent);
    netRentForTax = netRentForTaxation(rentalIncomeSD, yearlyLoanIntandPrin[1]);
    netTaxOnRent= taxOnRentPerc*netRentForTax;
    additionTaxSaving = additionTaxSaving24B(rentalIncomeSD, yearlyLoanIntandPrin[1]); 
    yearlynetFlow = (rentalIncome+additionTaxSaving)-(netTaxOnRent+yearlyEMI+additionalCost);
    monthlynetFlow = yearlynetFlow/12;
    cashFlowsYearly.push(yearlynetFlow);
    sheet.appendRow([date,rentalIncome, yearlyLoanIntandPrin[0],yearlyLoanIntandPrin[1],rentalIncomeSD,netRentForTax,netTaxOnRent,additionTaxSaving,yearlyEMI,principal,additionalCost,yearlynetFlow,monthlynetFlow, fv,xirr]);
    let lineItem = sheet.getLastRow();
    rowsAdded.push(lineItem);
    setColor(sheet, lineItem, color='#ffffff');
    cycles++;
  }
  const finalDate = "01/01/" + sellYear;
  fv = calculateFV(cost,realEstateGrowthPerc,1,  cycles ,0);
  if(loanTenure <= cycles) {
    additionalCost = 0;
  } else {
    additionalCost = principal;
    fv= fv-additionalCost;
  }
  sheet.appendRow([finalDate,0,0,0,0,0,0,0,0,"Future Value",additionalCost,fv,0]);
  let fvRow = sheet.getLastRow();
  setColor(sheet, fvRow, color='#90EE90');
  rowsAdded.push(fvRow);
  //sheet.appendRow([finalDate,0,0,0,0,0,0,0,0,"XIRR",0,0]);
  const lastRow = sheet.getLastRow();
  let firstRow=rowsAdded[0];
  let lastRow2=rowsAdded[rowsAdded.length-1];
  // Set a formula in the last row, e.g., in column D
  
  const formula = `=XIRR(L${firstRow}:L${lastRow2},A${firstRow}:A${lastRow2})`; // Example formula: Sum of columns A and B
  const lastRow1 = sheet.getLastRow();
  const range = sheet.getRange(lastRow1, 15);
  range.setNumberFormat("0.00%"); 
  range.setFormula(formula);
  createChart(sheet);
}

function getCellValue(row, column) {
  // Get the active spreadsheet and the active sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Investment Data";
  const sheet = spreadsheet.getSheetByName(sheetName);
  //const sheet = spreadSheet.getSheetByName(sheetName);
  const cellValue = sheet.getRange(row, column).getValue();
  // Return the value for further use
  return cellValue;
}

function calculateYearlyLoanIntandPrin(emi, principal, rateInterest) {
  let localPrincipal = principal;
  let monthlyPrincipal = 0;
  let yearlyPrincipal = 0;
  let yearlyInterest = 0;
  if(principal) {
    for(let i = 1; i<= 12 ; i++) {
      interest = (localPrincipal * rateInterest)/12;
      yearlyInterest += interest;
      monthlyPrincipal = emi - interest;
      localPrincipal = localPrincipal - monthlyPrincipal;
      yearlyPrincipal +=monthlyPrincipal; 
    }
  } else {
    yearlyPrincipal = 0;
    yearlyInterest = 0;
  }
  const yearlyLoanIntandPrin = [yearlyPrincipal, yearlyInterest];
  return yearlyLoanIntandPrin;
}

function netRentForTaxation(rentafterSD, loanInterest) {
  let netRentForTax= 0;
  netRentForTax = Math.max(rentafterSD-loanInterest,0);
  return netRentForTax;
}

function additionTaxSaving24B(rentafterSD, loanInterest) {
 let additionTaxSaving= 0;
  additionTaxSaving = ((Math.min(rentafterSD-loanInterest,0))*-1)*0.3;
  return additionTaxSaving;
}


function additionalCostCalculate(registeryCostPerc, cost, upfrontCost) {
  let additionalCost=0;
  additionalCost= (cost*registeryCostPerc)+upfrontCost;
  return additionalCost;
}

function calculateFV(presentValue, annualRate, periodsPerYear, years, payment) {
  var ratePerPeriod = annualRate / periodsPerYear;
  var totalPeriods = years * periodsPerYear;

  var fv = presentValue * Math.pow(1 + ratePerPeriod, totalPeriods) +
           payment * ((Math.pow(1 + ratePerPeriod, totalPeriods) - 1) / ratePerPeriod);

  return fv.toFixed(2); // Return the Future Value
}

function setColFormat(sheet) {
  let lastRow = sheet.getLastRow();
  const cols = ['A','B','C','D','E','F','G','H','I','K','L','M','N','O'];
  cols.forEach(function(col) {
    let columnRange = sheet.getRange(`${col}${lastRow}:${col}`);
    if(col == 'A') {
      columnRange.setNumberFormat("MM/dd/yyyy"); // Example date format
    } if (col == 'O') {
        columnRange.setNumberFormat("0.00%");
    } else {
      columnRange.setNumberFormat('₹#,##,##0.00');
    }
  });
}

function buildHeader(sheet) {
 // Build Header
  const startRow = 20;
  const headerRow = ["Date", "Rental Income", "Loan Principal Paid", "Loan Interest Paid", "Rent After Standard Deduction", "Net Rent For Tax","Tax On Rent","Additional Tax Saving","Loan EMI Yearly", "Loan Principal Remaining", "Additional Out", "Net Flow Yearly", "Net Flow Monthly", "Future Value", "XIRR"];
  const newRowRange = sheet.getRange(startRow, 1, 1, headerRow.length); // New row range
  newRowRange.setValues([headerRow]);
  setColFormat(sheet);
  const lastRow = sheet.getLastRow();
  setColor(sheet, lastRow, color='#808080');
}

function defineCell() {
  this.purchaseYearCell = [16,3];
  this.sellYearCell = [17,3];
  this.costRealEstate = [2,3];
  this.rentalYeild = [11,3];
  this.sdOnRent = [15,3];
  this.loanAmount = [4,3];
  this.loanInterestRate = [8,3];
  this.emiAmount = [10,3];
  this.taxOnRentPerc = [14,3];
  this.registeryCostPerc = [13,3];
  this.upfrontPayment = [3,3];
  this.realEstateGrowthPerc = [12,3];
  this.loanTenure = [5,3];
}

function setColor(sheet, rowNumberFrom, color) {
   const rowNumber = rowNumberFrom; // Change this to the row you want to color

  // Get the last column in the sheet to cover the entire row
  const lastColumn = sheet.getLastColumn();

  // Ensure the sheet has at least one column
  if (lastColumn === 0) {
    Logger.log("The sheet is empty or has no columns.");
    return;
  }

  // Get the range for the entire row
  const range = sheet.getRange(rowNumber, 1, 1, lastColumn);

  range.setHorizontalAlignment("center");
  // Set the background color for the row
  range.setBackground(color);
}

// Helper function to calculate XIRR
function calculateXirr(cashFlows, dates) {
  const MAX_ITER = 100; // Maximum number of iterations
  const PRECISION = 1e-6; // Desired precision

  let rate = 0.1; // Initial guess for the rate
  for (let i = 0; i < MAX_ITER; i++) {
    const f = cashFlows.reduce((acc, cf, idx) => acc + cf / Math.pow(1 + rate, daysBetween(dates[0], dates[idx]) / 365), 0);
    const fPrime = cashFlows.reduce((acc, cf, idx) => acc - (cf * daysBetween(dates[0], dates[idx]) / 365) / Math.pow(1 + rate, daysBetween(dates[0], dates[idx]) / 365 + 1), 0);
    
    const newRate = rate - f / fPrime;
    if (Math.abs(newRate - rate) < PRECISION) {
      return newRate;
    }
    rate = newRate;
  }

  throw new Error("XIRR calculation did not converge.");
}

// Helper function to calculate days between two dates
function daysBetween(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  return (d2 - d1) / (1000 * 60 * 60 * 24);
}

function createChart(sheet) {
  const lastRow = sheet.getLastRow();
  console.log(lastRow);
  // Define the range of data to use for the chart (example: data in columns A and B)
  //const range = sheet.getRange("A20:O20");
  const dataRange = sheet.getRange(22, 15, lastRow, 1); // From row 1, column 1, spanning 2 columns

  // Create the chart
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE) // Chart type (e.g., COLUMN, LINE, PIE, etc.)
    .addRange(dataRange) // Data range for the chart
    .setPosition(1, 1, 0, 0) // Position of the chart (row, column, offsetX, offsetY)
    .setOption('title', 'XIRR Tracker') // Chart title
    .setOption('legend', { position: 'bottom' }) // Chart legend position
    .setOption('hAxis', {
      title: 'Date',
      format: "MM/dd/yyyy",
      gridlines: { count: dataRange.getNumRows() },
      showTextEvery: 1,
      slantedText: true,
    })
    .setOption('pointSize', 5) // Highlight individual points
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { lineWidth: 2, pointSize: 5 }, // Configure series for visibility
    })
    .build();

  // Insert the chart into the sheet
  sheet.insertChart(chart);
}

function createNewSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  const sheetName = "Real Estate Return Tracker"; // Set the name of the new sheet
  
  // Check if a sheet with the same name already exists
  if (spreadsheet.getSheetByName(sheetName)) {
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetName));
  }
  
  // Create the new sheet
  const newSheet = spreadsheet.insertSheet(sheetName);

  if (newSheet) {
    //SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" has been created.`);
    return newSheet;
  } else {
    SpreadsheetApp.getUi().alert("Failed to create the new sheet.");
  }
}
