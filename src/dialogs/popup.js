/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import dialogs from '../dialogs/dialogs.js';
import utils from '../common/utils.js';

Office.onReady(() => {
    console.log("Office.onReady from popup.js");

    // Notify that the dialog is ready
    // Office.context.ui.messageParent("dialogReady");

    // // Set up the event handler for receiving messages
    // Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, dialogs.receiveMessageFromParent, function(result) {
    //     if (result.status === Office.AsyncResultStatus.Succeeded) {
    //         console.log("Event handler was successfully added.");
    //         // Only now do we notify that the event handler is ready
    //         Office.context.ui.messageParent("eventHandlerReady");
    //     } else {
    //         console.error("Failed to set event handler: " + result.error.message);
    //     }
    // });
});
// https://stackoverflow.com/questions/58136833/how-to-show-range-selection-input-dialog-in-excel-using-officejs
// window.loadRangeAddressHandler = utils.loadRangeAddressHandler;
window.promptForAddressRange = utils.promptForAddressRange;

const queryString = window.location.search;
console.log('queryString=' + queryString);
const urlParams = new URLSearchParams(queryString);
console.log('contentFile = ' + urlParams.get('contentFile'));
let message = urlParams.get('message');
let contentFile = urlParams.get('contentFile');
if (message) {
    document.getElementById('message').innerHTML = message;
} else if (contentFile) {
    utils.loadHtmlPage(contentFile);
    
  }
    

// await utils.loadHtmlPage("UserForm3InputDataRng");
// let action = await utils.detectUnloadAction();
// if (action === 'submit') {
//   const dataRange = document.getElementsByName('data_range_id');
//   console.log("data range" + dataRange);
//   await selectRangeStart();
//   copyRangeToNewWorkbook();
// }

// Office.onReady(() => {
//     console.log("office.onready from popup.js");
//     Office.context.ui.messageParent("dialogReady");
//     Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveMessageFromParent);
// })

// function receiveMessageFromParent(arg) {
//     const message = arg.message;
//     console.log(message);
//     document.getElementById("messageText").innerText =  message;

//     // Process the message...

// };

// async function loadRangeAddressHandler(){
//     await Excel.run(async (context) => {
//       const worksheet = context.workbook.worksheets.getActiveWorksheet();     
//       worksheet.onSelectionChanged.add(rangeSelectionHandler);
//       await context.sync();
//     }); 
//   }
  
// async function rangeSelectionHandler(event){
// await Excel.run(async (context) => {

//     let range = context.workbook.getSelectedRange();
//     range.load("address");
//     await context.sync();
//     document.getElementById("range_address_id").value = range.address;
//     document.getElementById("submit_button_id").disabled = false;
    

//     console.log(`The address of the selected range is "${range.address}"`);

// });
// }

async function selectRangeStart() {
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let rangeAddressInput = document.getElementById("range_address_id");
    if (rangeAddressInput) {
    rangeAddressInput = rangeAddressInput.value;
    const initialCellAddress = rangeAddressInput.split(':')[0];
    const initialCell = sheet.getRange(initialCellAddress);
    initialCell.select();
    await context.sync();
    }
});
}
async function copyRangeToNewWorkbook() {
try {
    const rangeElement = document.getElementById("range_address_id");
    if (!rangeElement) {
    throw new Error("Element with id 'range_address_id' not found");
    }
    const dataRange = rangeElement.value;
    if (!dataRange) {
    throw new Error("No value found in 'range_address_id' element");
    }

    console.log("Selected range:", dataRange);

    let rangeValues, rangeFormat;

    await Excel.run(async (context) => {
    let sourceRange = context.workbook.worksheets.getActiveWorksheet().getRange(dataRange);
    sourceRange.load(["values", "format"]);
    await context.sync();

    rangeValues = sourceRange.values;
    rangeFormat = sourceRange.format;

    console.log("Source values:", rangeValues);
    });

    let newWorkbook = await Excel.createWorkbook();
    
    await Excel.run(newWorkbook, async (newContext) => {
    let newSheet = newContext.workbook.worksheets.getItem("Sheet1");
    let newRange = newSheet.getRange(dataRange);
    
    newRange.values = rangeValues;
    newRange.format.fill.color = rangeFormat.fill.color;
    newRange.format.font.color = rangeFormat.font.color;
    newRange.format.font.bold = rangeFormat.font.bold;

    await newContext.sync();
    
    console.log("Range pasted to new workbook");

    // Force Excel to recalculate the new workbook
    newContext.application.calculate(Excel.CalculationType.full);
    await newContext.sync();
    });

    console.log("Operation completed successfully");

} catch (error) {
    console.error("Error:", error);
}
}

// Create a map of button IDs to functions
const functionMap = {
    'SelectIntervalData': SelectIntervalData,
    'SelectBillingData': SelectBillingData,
    // Add all other button ID-function pairs here
  };
  
  export default functionMap;
  
  
  // Define your functions
  // function SelectIntervalData() {
  async function SelectBillingData() {
    console.log("SelectBillingData called");
    
    // dialogs.openDialog("SelectBillingData called !!!!!!!");
    dialogs.openDialog("UserForm3InputDataRng", true);
    
  
    
    return "SelectBillingData"; 
  }

  async function SelectIntervalData() {
    console.log("SelectIntervalData called");
  
    Office.addin.showAsTaskpane(); 
    state.set("strNrmlzBillingData", "No");
    // selectData();
  
    await utils.loadHtmlPage("UserForm3InputDataRng");
    let action = await utils.detectUnloadAction();
    if (action === 'submit') {
      const dataRange = document.getElementsByName('data_range_id');
      console.log("data range" + dataRange);
      await selectRangeStart();
      copyRangeToNewWorkbook();
      // Process the data range as needed
      // copy range
      // open new workbook
      // create sheet data
      // paste range in sheet data
      // create sheet dictionary
      // write headings Field_Name Units Description
      // write field names
      // set validation list to units column DateTime, Date, Time, kWh
      // close the first workbook
      // open the taskpane in the new workbook
    }
  
    return "SelectIntervalData"; 
  }
  