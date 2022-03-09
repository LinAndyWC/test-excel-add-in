/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { component } = require("vue/types/umd");

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable; 

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
     
  } 
});


async function createTable() {
  await Excel.run(async (context) => {

      // TODO1: Queue table creation logic here.
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add("A1:E1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      // TODO2: Queue commands to populate the table with data.

      // let rowData = [];
      // fetchComponentJSON().then((component)=>{

      //   component.forEach(element => {
      //     let temp_array = Object.keys(element)
      //     .map(function(key) {
      //         return element[key];
      //     });
      //     rowData.add(temp_array);          
      //   });
      // })        

      // expensesTable.rows.add(null,rowData);

              
      expensesTable.getHeaderRowRange().values =
      [["Demand", "SKU", "Component", "Qty","source"]];
  
      expensesTable.rows.add(null /*add at the end*/, [
          ['DN_4125908778_000010_2', 'IG1493002974', '6Y17B0709501', '35', 'FG01',],
          ['DN_4125200867_000010_1', 'IG1493002623', '6Y17B0709501', '41', 'FG01'],
          ['DN_4125876975_000010', 'IG1493002965', '6Y17B0709501', '35', 'FG01',],
          ['DN_4125881705_000010_6', 'IG1493002969', '6Y17B0709501', '35', 'FG01'],
          ['DN_4125881706_000010_4', 'IG1493002970', '6Y17B0709501', '35', 'Incoming'],
          ['DN_4125881706_000010_5', 'IG1493002970', '6Y17B0709501', '35', 'Incoming']
      ]);
 



      // TODO3: Queue commands to format the table.
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

// async function fetchComponentJSON() {
  
//   var part_list = []; part_list.add(document.getElementById('part').value);
//   var url = 'http://imx-fp-3/hp_fp_web/UM01_SKU/WebExportReport/GetRPTComponentVsDemand';
//   var data = { part_number_list : part_list};
    
//   const response = await fetch(url, {
//       method: 'POST', // or 'PUT'
//       body: JSON.stringify(data), // data can be `string` or {object}!
//       headers: new Headers({
//         'Content-Type': 'application/json'
//       })
//     }) ;
//   const component = await response.json();
//   return component;
// }
