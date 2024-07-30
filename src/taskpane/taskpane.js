/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import state from './state.js';

window.stateSet = state.set;
window.stateGet = state.get;
window.getAddress = getAddress;
window.selectData = selectData;

  
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Host is Excel");
    await populateTestData();
  }
  console.log ("end of office onready in taskpane.js");
});

async function getAddress(){
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();     
    worksheet.onSelectionChanged.add(rangeSelectionHandler);
    await context.sync();
  }); 
}

async function rangeSelectionHandler(event){
  await Excel.run(async (context) => {

    let range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    document.getElementById("range_address_id").value = range.address;
    document.getElementById("submit_button_id").disabled = false;
    

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


    if (!sqlResult.ok) {
        throw new Error('Error inserting into SQL');
    }

  } catch (error) {
    console.error("Error:", error); 
    // Handle the error appropriately for your UI (display an error message, etc.)
  }
}

async function loadHtmlPage(pageName) {
  try {
    let response = await fetch(`/forms/${pageName}.html`);
    if (!response.ok) {
      throw new Error(`Failed to load the HTML page: ${response.statusText}`);
    }
    let htmlContent = await response.text();
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlContent;
    const scripts = tempDiv.querySelectorAll('script');
    const contentFrame = document.getElementById('content-frame');
    contentFrame.innerHTML = tempDiv.innerHTML; // Includes innerHTML without <script> tags
    for (const script of scripts){
      const scriptElement = document.createElement('script');
      scriptElement.type = 'text/javascript';
      if (script.type === 'module') {
        scriptElement.type = 'module';
      }
      scriptElement.textContent = script.textContent;
      document.body.appendChild(scriptElement); // Append to body and execute
    }
    console.log("Loaded HTML content successfully");
  } catch (error) {
    console.error('Error loading HTML content:', error);
  }
}


// Create a map of button IDs to functions
const functionMap = {
  'SelectIntervalData': SelectIntervalData,
  // Add all other button ID-function pairs here
};

export default functionMap;


// Define your functions
function SelectIntervalData() {
  console.log("SelectIntervalData called");
  Office.addin.showAsTaskpane(); 
  state.set("strNrmlzBillingData", "No");
  selectData();
  return "SelectIntervalData";  
}

async function detectTaskpaneUnloadAction() {
  return new Promise((resolve) => {
    const submitButton = document.getElementById("submit_button_id");
    const cancelButton = document.getElementById("cancel_button_id");
    const backButton = document.getElementById("back_button_id");
    async function handleClick(event) {
      if (event.target === submitButton) {
        resolve('submit');
      } else if (event.target === cancelButton) {
        resolve('cancel');
      } else if (event.target === backButton) {
        resolve('back');
      }
    }
    // Special event handler for click on taskpane close
    Office.addin.onVisibilityModeChanged(function(args) {
      if (args.visibilityMode == "Hidden") {
        resolve('close');
      }
    });
    submitButton.addEventListener("click", handleClick);
    cancelButton.addEventListener("click", handleClick);
  });
}

async function selectData(strAutomate = 'Manual') {
  let strNrmlzBillingData = state.get("strNrmlzBillingData");
  if (strAutomate != "Automate") {
    try {
      if (strNrmlzBillingData == "No") {
        await loadHtmlPage("UserForm4TimeStampCols");
        let action = await detectTaskpaneUnloadAction();
        if (action === 'submit') {
          await loadHtmlPage("UserForm3InputDataRng");
          action = await detectTaskpaneUnloadAction();
          if (action === 'submit') {
            const dataRange = document.getElementsByName('data_range_id');
            console.log("data range" + dataRange);
            // Process the data range as needed
          }
        }
      } else if (strNrmlzBillingData == "Yes") {
        console.log("Manual process with normalized billing data initiated");
        // Add specific logic for this case here, e.g., loading different forms
      }
    } catch (error) {
      console.error("Error in selectData:", error);
    } finally {
      await selectRangeStart();
      Office.addin.hide();
    }
  } else {
    if (strNrmlzBillingData == "No") {
      console.log("Automated process without normalized billing data initiated");
      // Add specific logic for this case here
    } else if (strNrmlzBillingData == "Yes") {
      console.log("Automated process with normalized billing data initiated");
      // Add specific logic for this case here
    }
  }
}

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


async function populateTestData() {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Helper function to generate random number between min and max
      const randomBetween = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;

      // Helper function to safely format date
      const formatDate = (date) => {
          if (isNaN(date.getTime())) return "Invalid Date";
          return date.toISOString().split('T')[0];
      };

      // Helper function to safely format time
      const formatTime = (date) => {
          if (isNaN(date.getTime())) return "Invalid Time";
          return date.toTimeString().split(' ')[0].substring(0, 5);
      };

      // Range 1: 15 minute interval data, 10 intervals, first column valid datetime, second column random kWh
      const startDate = new Date();
      const range1Data = [];
      for (let i = 0; i < 10; i++) {
          const dateTime = new Date(startDate.getTime() + i * 15 * 60000);
          range1Data.push([dateTime.toISOString().replace('T', ' ').replace('Z', ''), randomBetween(10, 100)]);

      }

      // Range 2: Similar to Range 1, but one value in column one is not a valid datetime
      const range2Data = range1Data.map((row, index) => index === 5 ? ["Invalid DateTime", randomBetween(10, 100)] : row);

      // Range 3: Similar to Range 1, but uses two columns for date and time, all valid
      const range3Data = range1Data.map(row => {
          const date = new Date(row[0]);
          return [formatDate(date), formatTime(date), row[1]];
      });

      // Range 4: Similar to Range 3, but one value in the date column is invalid
      const range4Data = range3Data.map((row, index) => index === 5 ? ["Invalid Date", row[1], row[2]] : row);

      // Fill the ranges with data
      const range1 = sheet.getRange("A1:B10");
      range1.values = range1Data;

      const range2 = sheet.getRange("A12:B21");
      range2.values = range2Data;

      const range3 = sheet.getRange("A23:C32");
      range3.values = range3Data;

      const range4 = sheet.getRange("A34:C43");
      range4.values = range4Data;

      await context.sync();
  });
}

async function validTimeSeriesRange(range) {
  try {
      await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const selectedRange = sheet.getRange(range);
          
          selectedRange.load(["values", "rowCount", "columnCount"]);
          await context.sync();

          const values = selectedRange.values;
          const firstRow = values[0];
          const firstColumn = values.map(row => row[0]);

          // Determine the number of time columns to check based on global state
          const iTimeCols = window.stateGet('iTimeCols'); // Assuming stateGet is a function that retrieves the global state

          // Check if the first row contains valid field names
          const validFieldName = name => /^[a-zA-Z_][a-zA-Z0-9_]*$/.test(name);
          const areFieldNamesValid = firstRow.every(fieldName => validFieldName(fieldName));

          // Check if the first column contains valid datetime values or date values
          const validDateTime = value => !isNaN(Date.parse(value));
          const validDate = value => !isNaN(Date.parse(value.split('T')[0]));
          const validTime = value => /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(value);

          let areTimeValuesValid;

          if (iTimeCols === 1) {
              // Only the first column should contain datetime values
              areTimeValuesValid = firstColumn.every(value => validDateTime(value));
          } else if (iTimeCols === 2) {
              // The first column should contain date values and the second column should contain time values
              const secondColumn = values.map(row => row[1]);
              areTimeValuesValid = firstColumn.every(value => validDate(value)) && secondColumn.every(value => validTime(value));
          } else {
              // Invalid iTimeCols value, return false
              areTimeValuesValid = false;
          }

          return areFieldNamesValid && areTimeValuesValid;
      });
  } catch (error) {
      console.error(error);
      return false;
  }
}
