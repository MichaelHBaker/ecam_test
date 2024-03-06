/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

var iTimeCols;

export function setGlobal (var_name, value) {
  if (var_name in window) {
    eval(var_name + '=' + value);
    console.log("setGlobal iTimeCols " + iTimeCols);
  } else {
    throw `${var_name} has not been defined as a global variable`;
  }
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    // Assign event handlers and other initialization logic.
    document.getElementById("range_add_id").onclick = getAddress;

  }

  });

async function getAddress(event){
  // Additional Excel.run can be placed here if needed
  try {
    await Excel.run(async (context) => {
      // Asynchronous Excel operations here
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      

      worksheet.onSelectionChanged.add(changeHandler);
      await context.sync();
      console.log("Event handler added");
    }); 
  }
  catch(error){
    // need to figure out how to display dialogs on all catches
    console.error(error);
  }

}

async function changeHandler(event){
  await Excel.run(async (context) => {

    let range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("range_add_id").value = range.address;

    console.log(`The address of the selected range is "${range.address}"`);

  });
}