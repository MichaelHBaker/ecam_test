/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */


var iTimeCols;
var strNrmlzBillingData;

// function setGlobal (var_name, value) {
//   if (var_name in window) {
//     eval(var_name + '=' + value);
//     console.log("setGlobal iTimeCols " + iTimeCols);
//   } else {
//     throw `${var_name} has not been defined as a global variable`;
//   }
// }

function setGlobal(var_name, value) {
  if (var_name in window) {
    window[var_name] = value;
    console.log(`setGlobal ${var_name} = ${window[var_name]}`);
  } else {
    throw new Error(`${var_name} has not been defined as a global variable`);
  }
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    // Assign event handlers and other initialization logic.
    document.getElementById("range_add_id").onclick = getAddress;
    document.getElementById("fetchBtn").onclick = fetchData;
    document.getElementById("writeBtn").onclick = writeData;
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

async function fetchData() {
  try {
    // Fetch Weather Data
    const weatherResponse = await fetch('/weatherdata');
    const jsonString = await weatherResponse.text();
    const weatherData = JSON.parse(jsonString); 

    // Extract max temperature
    const maxTempF = weatherData.forecast.forecastday[0].day.maxtemp_f;

    // Write to Excel
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("B1"); 
      range.values = maxTempF.toString(); 
      await context.sync(); 
    });

  } catch (error) {
    console.error("Error:", error); 
    // Handle the error appropriately for your UI (display an error message, etc.)
  }
}

async function writeData() {

  var maxTempF;

  // Read from Excel
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B1"); 
    range.load('values');
    await context.sync(); 
    maxTempF = range.values;
  });
  
  try { 

    const weatherResponse = await fetch(`/insertweatherdata?max_temp_f=${maxTempF}`);
    const jsonString = await weatherResponse.text();
    const weather_response = JSON.parse(jsonString); 
    console.log(weather_response);

    // Send Request for SQL Insertion
    // const sqlResult = await fetch('http://127.0.0.1:8000/insertweatherdata', {
    //     method: 'POST', 
    //     headers: { 'Content-Type': 'application/json'  }, 
    //     body: JSON.stringify({ temperature: maxTempF }) 
    // });

    if (!sqlResult.ok) {
        throw new Error('Error inserting into SQL');
    }

  } catch (error) {
    console.error("Error:", error); 
    // Handle the error appropriately for your UI (display an error message, etc.)
  }
}

// async function fetchData() {
//   try {
//     const response = await fetch('/weatherdata'); 
//     const jsonString = await response.text(); // Get raw JSON text
//     const weatherData = JSON.parse(jsonString); 

//     // Extract max temperature
//     const maxTempF = weatherData.forecast.forecastday[0].day.maxtemp_f;

//     // Using Office.js to write into Excel 
//     await Excel.run(async (context) => {
//       const sheet = context.workbook.worksheets.getActiveWorksheet();
//       const range = sheet.getRange("B1"); // Example: Place max temp in cell B1
//       range.values = maxTempF.toString();  

//       await context.sync(); 
//     });

//   } catch (error) {
//     console.error("Error fetching or processing data:", error); 
//     // Handle the error appropriately for your UI 
//   }
// }

// function SelectIntervalData() {
//   // ui load
//   // form button clicks
//   // other business code
//   return "SelectIntervalData";  

// }
