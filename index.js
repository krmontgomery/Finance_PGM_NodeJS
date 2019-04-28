// ==================================================================*
//          Program Description: Node program that creates           *
//          a spreadsheet that shows income & expenses in            *
//          both 'xlsx' & 'CSV' format. This program utilizes        *
//          exceljs npm package for most of it's functionality.      *
//                                                                   *   
//          Author: Kris Montgomery         Version: 01              *
//                                                                   *
//          Date_Created: 3/15/2019                                  *
//                                                                   *
// ==================================================================*
// Required packages as variables
const Excel = require('exceljs');
const workbook = new Excel.Workbook();
// ==================================================================*
// Workbook info configured
workbook.creator = 'Kris';
workbook.lastModifiedBy = 'Kris Montgomery';
workbook.created = new Date (2019, 3, 9);
workbook.modified = new Date();
workbook.lastPrinted = new Date();
// Set workbook dates to 1904 date system
workbook.properties.date1904 = true;
//Controls how many separate windows Excel will open when viewing
//Viewing the workbook
workbook.views = [
    {
      x: 0, y: 0, width: 10000, height: 20000,
      firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
  ]
// ==================================================================*
// Declaring a worksheet for workbook
var sheet = workbook.addWorksheet('Finance Spreadsheet');
var sheetTwo = workbook.addWorksheet('Misc Expenditures');
//Accessing worksheet
var worksheet = workbook.getWorksheet('Finance Spreadsheet');
var worksheetTwo = workbook.getWorksheet('Misc Expenditures');
// ==================================================================*
//               SpreadSheet Finances Worksheet                      *
// ==================================================================*
// Setting Columns
worksheet.columns = [
    {header:'Income', key: 'text', width:32},//Column One or A
    {header:'January', key: 'JAN', width: 15},
    {header:'February', key: 'FEB', width: 15},//Column Three or C
    {header:'March', key: 'MAR', width: 15},
    {header:'April', key: 'APR', width: 15},//Column Five or E
    {header:'May', key: 'MAY', width: 15},
    {header:'June', key: 'JUN', width: 15},//Column Seven or G
    {header:'July', key: 'JUL', width: 15},
    {header:'August', key: 'AUG', width: 15},//Column Nine or I
    {header:'September', key: 'SEP', width: 15},
    {header:'October', key: 'OCT', width: 15},//Column Eleven or K
    {header:'November', key: 'NOV', width: 15},
    {header:'December', key: 'DEC', width: 15},//Column Thirteen or M
    {header:'Year Total', key:'YearTotal', width: 15}
]
// ==================================================================*
//               SpreadSheet Main Layout                             *
// ==================================================================*
const prefix = '$';// variable made for property value
// Income portion of spreadsheet 
worksheet.addRow({text:"Income one", JAN:`${prefix}500`, FEB:`${prefix}500`, MAR:`${prefix}500`, APR: `${prefix}500`, MAY:`${prefix}500`});
worksheet.addRow({text:"Income two", JAN:`${prefix}2000`, FEB:`${prefix}2000`, MAR:`${prefix}2000`, APR:`${prefix}2000`,MAY:`${prefix}2000`});
worksheet.addRow({text:"Income three", JAN:`${prefix}1000`, FEB:`${prefix}1000`, MAR:`${prefix}1000`, APR:`${prefix}1000`, MAY:`${prefix}1000`});
worksheet.addRow({text:"Income four", JAN:`${prefix}10`, FEB:`${prefix}10`, MAR:`${prefix}10`, APR:`${prefix}10`, MAY:`${prefix}10`});
// Break*
const totalBreak = worksheet.getRow(6);// ** Break for total line **
totalBreak.addPageBreak();
// *
worksheet.addRow({text:'Income Total:'});// Income Total Line
worksheet.getCell('B7').value = {formula:'B2+B3+B4+B5', result:`${prefix}7`};
worksheet.getCell('C7').value = {formula:'C2+C3+C4+C5', result:`${prefix}7`};
worksheet.getCell('D7').value = {formula:'D2+D3+D4+D5', result:`${prefix}7`};
worksheet.getCell('E7').value = {formula:'E2+E3+E4+E5', result:`${prefix}7`};
worksheet.getCell('F7').value = {formula:'F2+F3+F4+F5', result:`${prefix}7`};
// Expenses Portion of Spreadsheet
// Break*
const incomeBreak = worksheet.getRow(8 && 9);// ** Break between income and expense **
incomeBreak.addPageBreak();
// *
worksheet.addRow({text:'Home Expenses'});// Expenses Header
worksheet.addRow({text:'Mortgage', JAN:`${prefix}500`, FEB:`${prefix}500`, MAR:`${prefix}500`, APR:`${prefix}500`, MAY:`${prefix}500`});
worksheet.addRow({text:'CarPymntOne', JAN:`${prefix}200`, FEB:`${prefix}200`, MAR:`${prefix}200`, APR:`${prefix}200`, MAY:`${prefix}200`});
worksheet.addRow({text:'CarPymntTwo', JAN:`${prefix}200`, FEB:`${prefix}200`, MAR:`${prefix}200`, APR:`${prefix}200`, MAY:`${prefix}200`});
worksheet.addRow({text:'CarInsuranceOne', JAN:`${prefix}75`, FEB:`${prefix}75`, MAR:`${prefix}75`, APR:`${prefix}75`, MAY:`${prefix}75`});
worksheet.addRow({text:'CarInsuranceTwo', JAN:`${prefix}100`, FEB:`${prefix}100`, MAR:`${prefix}100`, APR:`${prefix}100`, MAY:`${prefix}100`});
worksheet.addRow({text:'Gas(Travel)', JAN:`${prefix}100`, FEB:`${prefix}100`, MAR:`${prefix}100`, APR:`${prefix}100`, MAY:`${prefix}100`});
worksheet.addRow({text:'PhonePymtOne', JAN:`${prefix}50`, FEB:`${prefix}50`, MAR:`${prefix}50`, APR:`${prefix}50`, MAY:`${prefix}50`});
worksheet.addRow({text:'PhonePymtTwo', JAN:`${prefix}50`, FEB:`${prefix}50`, MAR:`${prefix}50`, APR:`${prefix}50`, MAY:`${prefix}50`});
worksheet.addRow({text:'Internet', JAN:`${prefix}80`, FEB:`${prefix}80`, MAR:`${prefix}80`, APR:`${prefix}80`, MAY:`${prefix}80`});
worksheet.addRow({text:'Electric', JAN:`${prefix}120`, FEB:`${prefix}120`, MAR:`${prefix}120`, APR:`${prefix}120`, MAY:`${prefix}120`});
worksheet.addRow({text:'Trash&Sewer', JAN:`${prefix}50`, FEB:`${prefix}50`, MAR:`${prefix}50`, APR:`${prefix}50`, MAY:`${prefix}50`});
worksheet.addRow({text:'Water', JAN:`${prefix}50`, FEB:`${prefix}50`, MAR:`${prefix}50`, APR:`${prefix}50`, MAY:`${prefix}50`});
worksheet.addRow({text:'Groceries', JAN:`${prefix}250`, FEB:`${prefix}250`, MAR:`${prefix}250`, APR:`${prefix}250`, MAY:`${prefix}250`});
worksheet.addRow({text:'Natural Gas', JAN:`${prefix}95`, FEB:`${prefix}95`, MAR:`${prefix}95`, APR:`${prefix}95`, MAY:`${prefix}95`});
worksheet.addRow({text:'CreditCardOne', JAN:`${prefix}50`, FEB:`${prefix}50`, MAR:`${prefix}50`, APR:`${prefix}50`, MAY:`${prefix}50`});
worksheet.addRow({text:'CreditCardTwo', JAN:`${prefix}100`, FEB:`${prefix}100`, MAR:`${prefix}100`, APR:`${prefix}100`, MAY:`${prefix}100`});
worksheet.addRow({text:'Savings', JAN:`${prefix}100`, FEB:`${prefix}100`, MAR:`${prefix}100`, APR:`${prefix}100`, MAY:`${prefix}100`});
worksheet.addRow({text:'Misc Expenditures', JAN:`${prefix}230`, FEB:`${prefix}230`, MAR:`${prefix}230`, APR:`${prefix}230`, MAY:`${prefix}230`});
// Break
const expTotBreak = worksheet.getRow(30);// ** Break for Expense total line **
expTotBreak.addPageBreak();
// *
worksheet.addRow({text:'Expense Totals:'});//Expense Total Line
worksheet.getCell('B31').value = {formula:'B11+B12+B13+B14+B15+B16+B17+B18+B19+B20+B21+B22+B23+B24+B25+B26+B27+B28+B29', result:`${prefix}7`};
worksheet.getCell('C31').value = {formula:'C11+C12+C13+C14+C15+C16+C17+C18+C19+C20+C21+C22+C23+C24+C25+C26+C27+C28+C29', result:`${prefix}7`};
worksheet.getCell('D31').value = {formula:'D11+D12+D13+D14+D15+D16+D17+D18+D19+D20+D21+D22+D23+D24+D25+D26+D27+D28+D29', result:`${prefix}7`};
worksheet.getCell('E31').value = {formula:'E11+E12+E13+E14+E15+E16+E17+E18+E19+E20+E21+E22+E23+E24+E25+E26+E27+E28+E29', result:`${prefix}7`};
worksheet.getCell('F31').value = {formula:'F11+F12+F13+F14+F15+F16+F17+F18+F19+F20+F21+F22+F23+F24+F25+F26+F27+F28+F29', result:`${prefix}7`};
// Break
const totalIncomeLeft = worksheet.getRow(32 && 33);
totalIncomeLeft.addPageBreak();
// *
worksheet.addRow({text:'Total Income Leftover'});//Income Leftover for the month
worksheet.getCell('B34').value = {formula:'B7-B31',result:`${prefix}7`};
worksheet.getCell('C34').value = {formula:'C7-C31',result:`${prefix}7`};
worksheet.getCell('D34').value = {formula:'D7-D31',result:`${prefix}7`};
worksheet.getCell('E34').value = {formula:'E7-E31',result:`${prefix}7`};
worksheet.getCell('F34').value = {formula:'F7-F31',result:`${prefix}7`};
// ==================================================================*
// ==================================== Accumulated Yearly Totals ===*
// Income totals for the year accumulated
worksheet.getCell('N2').value = {formula:'B2+C2+D2+E2+F2+G2+H2+I2+J2+K2+L2+M2', result:`${prefix}7`};
worksheet.getCell('N3').value = {formula:'B3+C3+D3+E3+F3+G3+H3+I3+J3+K3+L3+M3', result:`${prefix}7`};
worksheet.getCell('N4').value = {formula:'B4+C4+D4+E4+F4+G4+H4+I4+J4+K4+L4+M4', result:`${prefix}7`};
worksheet.getCell('N5').value = {formula:'B5+C5+D5+E5+F5+G5+H5+I5+J5+K5+L5+M5', result:`${prefix}7`};
// Expense total bill amount paid for the year
worksheet.getCell('N11').value = {formula:'B11+C11+D11+E11+F11+G11+H11+I11+J11+K11+L11+M11', result:`${prefix}7`};
worksheet.getCell('N12').value = {formula:'B12+C12+D12+E12+F12+G12+H12+I12+J12+K12+L12+M12', result:`${prefix}7`};
worksheet.getCell('N13').value = {formula:'B13+C13+D13+E13+F13+G13+H13+I13+J13+K13+L13+M13', result:`${prefix}7`};
worksheet.getCell('N14').value = {formula:'B14+C14+D14+E14+F14+G14+H14+I14+J14+K14+L14+M14', result:`${prefix}7`};
worksheet.getCell('N15').value = {formula:'B15+C15+D15+E15+F15+G15+H15+I15+J15+K15+L15+M15', result:`${prefix}7`};
worksheet.getCell('N16').value = {formula:'B16+C16+D16+E16+F16+G16+H16+I16+J16+K16+L16+M16', result:`${prefix}7`};
worksheet.getCell('N17').value = {formula:'B17+C17+D17+E17+F17+G17+H17+I17+J17+K17+L17+M17', result:`${prefix}7`};
worksheet.getCell('N18').value = {formula:'B18+C18+D18+E18+F18+G18+H18+I18+J18+K18+L18+M18', result:`${prefix}7`};
worksheet.getCell('N19').value = {formula:'B19+C19+D19+E19+F19+G19+H19+I19+J19+K19+L19+M19', result:`${prefix}7`};
worksheet.getCell('N20').value = {formula:'B20+C20+D20+E20+F20+G20+H20+I20+J20+K20+L20+M20', result:`${prefix}7`};
worksheet.getCell('N21').value = {formula:'B21+C21+D21+E21+F21+G21+H21+I21+J21+K21+L21+M21', result:`${prefix}7`};
worksheet.getCell('N22').value = {formula:'B22+C22+D22+E22+F22+G22+H22+I22+J22+K22+L22+M22', result:`${prefix}7`};
worksheet.getCell('N23').value = {formula:'B23+C23+D23+E23+F23+G23+H23+I23+J23+K23+L23+M23', result:`${prefix}7`};
worksheet.getCell('N24').value = {formula:'B24+C24+D24+E24+F24+G24+H24+I24+J24+K24+L24+M24', result:`${prefix}7`};
worksheet.getCell('N25').value = {formula:'B25+C25+D25+E25+F25+G25+H25+I25+J25+K25+L25+M25', result:`${prefix}7`};
worksheet.getCell('N26').value = {formula:'B26+C26+D26+E26+F26+G26+H26+I26+J26+K26+L26+M26', result:`${prefix}7`};
worksheet.getCell('N27').value = {formula:'B27+C27+D27+E27+F27+G27+H27+I27+J27+K27+L27+M27', result:`${prefix}7`};
worksheet.getCell('N28').value = {formula:'B28+C28+D28+E28+F28+G28+H28+I28+J28+K28+L28+M28', result:`${prefix}7`};
// Total lines accumulated
// Total Income Line Accumulated
worksheet.getCell('N7').value = {formula:'B7+C7+D7+E7+F7+G7+H7+I7+J7+K7+L7+M7', result:`${prefix}7`};
// Expenses Total line accumulated
worksheet.getCell('N31').value = {formula:'B31+C31+D31+E31+F31+G31+H31+I31+J31+K31+L31+M31', result:`${prefix}7`};
// Income leftover accumulated 
worksheet.getCell('N34').value = {formula:'B34+C34+D34+E34+F34+G34+H34+I34+J34+K34+L34+M34', result:`${prefix}7`};
worksheet.getCell('N10').value = 'Year Total';
// ==================================================================*
// ==================================== Styling for Headers =========*
// Coloring for the header and total lines
// Color/Styling for Headers
var firstRowHighlight = worksheet.getRow(1);
firstRowHighlight.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb: 'FF4682B4'}
}
firstRowHighlight.font = {
    color: {argb:'FFFFFFFF'}
}
var secondRowHighlight = worksheet.getRow(10);
secondRowHighlight.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb: 'FF4682B4'}
}
secondRowHighlight.font = {
    color: {argb:'FFFFFFFF'}
}
// Color/Styling for Income Total Lines
var totalIncomeLineOne = worksheet.getRow(7);
totalIncomeLineOne.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb: 'FFB0C4DE'}
}
var expenseTotalLine = worksheet.getRow(31);
expenseTotalLine.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb: 'FFB0C4DE'}
}
// Color/Styling for total money leftover line
var moneyLeftover = worksheet.getRow(34);
moneyLeftover.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor:{argb:'FF708090'}
}
moneyLeftover.font = {
    color: {argb:'FFFFFFFF'}
}
// ==================================================================*
//                              WRITING FILES                        *
// ==================================================================*
// Code that generates the files, both xlsx and CSV
workbook.xlsx.writeFile('Finance.xlsx')
  .then( (err) => {
      if(err) {
          console.error('xlsx file was *not* created',err);
      } else {
      console.log('xlsx file was created')
      }
  })
workbook.csv.writeFile('Finance.csv')
  .then( (err) => {
      if(err) {
          console.error('CSV file was *not* created',err);
      } else {
          console.log('CSV file was created')
      }
  })
// ==================================================================*