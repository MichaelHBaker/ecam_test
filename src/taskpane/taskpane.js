/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import state from './state.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    console.log("Office.onReady in Taskpane run");

 }

  });

window.stateSet = state.set;
window.stateGet = state.get;
  
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
    // Fetch the HTML content
    let response = await fetch(`/forms/${pageName}.html`);
    if (!response.ok) {
      throw new Error(`Failed to load the HTML page: ${response.statusText}`);
    }

    let htmlContent = await response.text();
    // console.log(`Formed address of body page: ${htmlContent}`);
    
    // Create a temporary container to parse the fetched HTML
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlContent;

    // Extract script tags and their content
    const scripts = tempDiv.getElementsByTagName('script');
    const scriptContents = [];
    for (const script of scripts) {
      scriptContents.push(script.innerText);
      script.parentNode.removeChild(script); // Remove script tag from tempDiv
    }

    // Clear existing content and event listeners
    const contentFrame = document.getElementById('content-frame');
    const newContentFrame = contentFrame.cloneNode(false); // Clone without children to remove event listeners
    contentFrame.parentNode.replaceChild(newContentFrame, contentFrame);

    // Replace the body's content with the fetched HTML content (without script tags)
    newContentFrame.innerHTML = tempDiv.innerHTML;

    // Dynamically create and append script elements to the body
    for (const scriptContent of scriptContents) {
      const scriptElement = document.createElement('script');
      scriptElement.type = 'text/javascript';
      scriptElement.text = scriptContent;
      document.body.appendChild(scriptElement); // Append to body to execute the script
      // console.log('Executed script:', scriptContent);
    }

    console.log("Loaded HTML content successfully");
  } catch (error) {
    console.error('Error loading HTML content:', error);
  }
}

async function handleFormSubmission() {
  const submitButton = document.getElementById("submit_button_id");
  const cancelButton = document.getElementById("cancel_button_id");

  try {
    const clickedButton = await Promise.race([
      new Promise(resolve => submitButton.addEventListener("click", () => resolve(submitButton))),
      new Promise(resolve => cancelButton.addEventListener("click", () => resolve(cancelButton))),
    ]);

    // look for id attribute of clickebutton
    console.log("clicked button value:" + clickedButton + " type of:" + typeof clickedButton)
  } catch (error) {
    console.error("Error handling button clicks:", error);
    // Handle potential errors (e.g., button not found)
  }
}


// function waitForSubmit(buttonId) {
//   return new Promise((resolve) => {
//     const button = document.getElementById(buttonId);
//     button.addEventListener('click', () => {
//       resolve();
//     }, { once: true });
//   });
// }
// Create a map of button IDs to functions
const functionMap = {
  'SelectIntervalData': SelectIntervalData,
  // Add all other button ID-function pairs here
};

export default functionMap;


// Define your functions
function SelectIntervalData() {

  Office.addin.showAsTaskpane();

  console.log("SelectIntervalData called");
  
  state.set("strNrmlzBillingData", "No");
  SelectData();
  
  return "SelectIntervalData";  
  
}




async function SelectData(strAutomate='Manual') {
  let strNrmlzBillingData = state.get("strNrmlzBillingData");
  if (strAutomate != "Automate") {
    if (strNrmlzBillingData == "No") {
      await loadHtmlPage("UserForm4TimeStampCols");
      // await waitForSubmit('submit-button-id'); 
      await handleFormSubmission();
      // If iTimeCols = 5 Then GoTo FormTerminated
      if (state.get("iTimeCols") != 5) {
        await loadHtmlPage("UserForm3InputDataRng");

        // what do we do if cancel

        console.log("SelectData !!!");
      }
      else {
        Office.addin.hide();
      }
      }
    }
    return "";  
  }


