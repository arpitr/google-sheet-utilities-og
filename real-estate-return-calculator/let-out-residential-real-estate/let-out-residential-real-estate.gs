function generateReport() {
  defineCell();
  // Get the active spreadsheet and the first sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  clearData(sheet);


  // Build Header
  sheet.appendRow(["Date", "Rental Income", "Loan Principal Paid", "Loan Interest Paid", "Rent After Standard Deduction", "Net Rent For Tax","Tax On Rent","Additional Tax Saving","Loan EMI Yearly", "Loan Principal Remaining", "Additional Out", "Net Flow Yearly", "Net Flow Monthly"]);
  setColFormat(sheet);

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
  let rowsAdded = [];

  for(let i=purchaseYear; i <= rentTillYear; i++) {
   const date = "01-01-" + i;
   const yearlyLoanIntandPrin = calculateYearlyLoanIntandPrin(emiAmount, principal,loanInterestRate);
   if(principal) {
      yearlyEMI = emiAmount*12;
    } else {
      yearlyEMI = 0;
    }
   principal = Math.max(principal- yearlyLoanIntandPrin[0],0);
   if(i == purchaseYear) {
    rentalIncome = cost*rentalYeild;
    additionalCost = additionalCostCalculate(registeryCostPerc,cost,upfrontPayment);
   } else {
    rentalIncome = rentalIncome*1.1;
    additionalCost = 0;
   }
    rentalIncomeSD = rentalIncome*(1-sdOnRent);
    netRentForTax = netRentForTaxation(rentalIncomeSD, yearlyLoanIntandPrin[1]);
    netTaxOnRent= taxOnRentPerc*netRentForTax;
    additionTaxSaving = additionTaxSaving24B(rentalIncomeSD, yearlyLoanIntandPrin[1]);
    yearlynetFlow = (rentalIncome+additionTaxSaving)-(netTaxOnRent+yearlyEMI+additionalCost);
    monthlynetFlow = yearlynetFlow/12;
    sheet.appendRow([date,rentalIncome, yearlyLoanIntandPrin[0],yearlyLoanIntandPrin[1],rentalIncomeSD,netRentForTax,netTaxOnRent,additionTaxSaving,yearlyEMI,principal,additionalCost,yearlynetFlow,monthlynetFlow]);
    rowsAdded.push(sheet.getLastRow());
    cycles++;
  }
  const finalDate = "01-01-" + sellYear;
  fv = calculateFV(cost,realEstateGrowthPerc,1,  cycles ,0);
  if(loanTenure <= cycles) {
    additionalCost = 0;
    console.log(principal);
  } else {
    console.log(principal);
    additionalCost = principal;
    fv= fv-additionalCost;
  }
  sheet.appendRow([finalDate,0,0,0,0,0,0,0,0,"Future Value",additionalCost,fv,0]);
  rowsAdded.push(sheet.getLastRow());
  sheet.appendRow([finalDate,0,0,0,0,0,0,0,0,"XIRR",0,0]);
  const lastRow = sheet.getLastRow();
  let firstRow=rowsAdded[0];
  let lastRow2=rowsAdded[rowsAdded.length-1];
  // Set a formula in the last row, e.g., in column D

  const formula = `=XIRR(L${firstRow}:L${lastRow2},A${firstRow}:A${lastRow2})`; // Example formula: Sum of columns A and B
  const lastRow1 = sheet.getLastRow();
  const range = sheet.getRange(lastRow1, 12);
  range.setNumberFormat("0.00%");
  range.setFormula(formula);
}

function getCellValue(row, column) {
  // Get the active spreadsheet and the active sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  const cols = ['B','C','D','E','F','G','H','I','K','L'];
  cols.forEach(function(col) {
    let columnRange = sheet.getRange(`${col}${lastRow}:${col}`);
    columnRange.setNumberFormat('₹#,##,##0.00');
  });
}

function clearData(sheet) {
  const frozenRows = sheet.getFrozenRows(); // Get the number of frozen rows
  const lastRow = sheet.getLastRow();
  sheet.deleteRows(20, (lastRow-frozenRows)-1);
}

function defineCell() {
  this.purchaseYearCell = [18,3];
  this.sellYearCell = [19,3];
  this.costRealEstate = [2,3];
  this.rentalYeild = [12,3];
  this.sdOnRent = [17,3];
  this.loanAmount = [4,3];
  this.loanInterestRate = [8,3];
  this.emiAmount = [11,3];
  this.taxOnRentPerc = [16,3];
  this.registeryCostPerc = [15,3];
  this.upfrontPayment = [3,3];
  this.realEstateGrowthPerc = [14,3];
  this.loanTenure = [5,3];
}
