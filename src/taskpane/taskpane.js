/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */


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
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    await context.sync();
    
    document.getElementById("range_add_id").value = worksheet.name + "!" + event.address;

    console.log("event happended - address" + event.address);
    console.log("event happended - source" + event.source);
  });
}