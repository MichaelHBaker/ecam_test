/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import state from './state.js';

window.stateSet = state.set;
window.stateGet = state.get;
window.getAddress = getAddress;
  
Office.onReady(info => { 
  if (info.host === Office.HostType.Excel) {
    console.log("Host is Excel");
  }
  console.log ("end of office onready in taskpane.js");
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
    document.getElementById("range_address_id").value = range.address;

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

// new version from chatgpt - simpler - also works
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

    // Clear and replace the content frame
    const contentFrame = document.getElementById('content-frame');
    contentFrame.innerHTML = tempDiv.innerHTML; // Includes innerHTML without <script> tags

    // Execute scripts
    scripts.forEach(script => {
      const scriptElement = document.createElement('script');
      scriptElement.type = 'text/javascript';
      // Check if it's a module and adjust accordingly
      if (script.type === 'module') {
        scriptElement.type = 'module';
      }
      scriptElement.textContent = script.textContent;
      document.body.appendChild(scriptElement); // Append to body to ensure global scope
    });

    console.log("Loaded HTML content successfully");
  } catch (error) {
    console.error('Error loading HTML content:', error);
  }
}

// Prior version -- this works
// async function loadHtmlPage(pageName) {
//   try {
//     // Fetch the HTML content
//     let response = await fetch(`/forms/${pageName}.html`);
//     if (!response.ok) {
//       throw new Error(`Failed to load the HTML page: ${response.statusText}`);
//     }

//     let htmlContent = await response.text();
//     // console.log(`Formed address of body page: ${htmlContent}`);
    
//     // Create a temporary container to parse the fetched HTML
//     const tempDiv = document.createElement('div');
//     tempDiv.innerHTML = htmlContent;

//     // Extract script tags and their content
//     const scripts = tempDiv.getElementsByTagName('script');
//     const scriptContents = [];
//     for (const script of scripts) {
//       scriptContents.push(script.innerText);
//       script.parentNode.removeChild(script); // Remove script tag from tempDiv
//     }

//     // Clear existing content and event listeners
//     const contentFrame = document.getElementById('content-frame');
//     const newContentFrame = contentFrame.cloneNode(false); // Clone without children to remove event listeners
//     contentFrame.parentNode.replaceChild(newContentFrame, contentFrame);

//     // Replace the body's content with the fetched HTML content (without script tags)
//     newContentFrame.innerHTML = tempDiv.innerHTML;

//     // Dynamically create and append script elements to the body
//     for (const scriptContent of scriptContents) {
//       const scriptElement = document.createElement('script');
//       scriptElement.type = 'text/javascript';
//       scriptElement.text = scriptContent;
//       document.body.appendChild(scriptElement); // Append to body to execute the script
//       // console.log('Executed script:', scriptContent);
//     }

//     console.log("Loaded HTML content successfully");
//   } catch (error) {
//     console.error('Error loading HTML content:', error);
//   }
// }

async function waitForButtonClick() {
  return new Promise((resolve) => {
    const submitButton = document.getElementById("submit_button_id");
    const cancelButton = document.getElementById("cancel_button_id");

    function handleClick(event) {
      cleanup();
      resolve(event.target === submitButton); 
    }

    function cleanup() {
      submitButton.removeEventListener("click", handleClick);
      cancelButton.removeEventListener("click", handleClick);
    }

    submitButton.addEventListener("click", handleClick);
    cancelButton.addEventListener("click", handleClick);
  });
}


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

async function SelectData(strAutomate = 'Manual') {
  let strNrmlzBillingData = state.get("strNrmlzBillingData");
  if (strAutomate != "Automate") {
    try {
      if (strNrmlzBillingData == "No") {
        await loadHtmlPage("UserForm4TimeStampCols");
        const firstSubmitButton = await waitForButtonClick();       
        if (firstSubmitButton) {
          await loadHtmlPage("UserForm3InputDataRng");
          const secondSubmitButton = await waitForButtonClick();
          if (secondSubmitButton) {
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
      console.error("Error in SelectData:", error);
    } finally {
      // Always hide the add-in after non-automated processes, regardless of outcome
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