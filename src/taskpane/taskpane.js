/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */


var iTimeCols;
var strNrmlzBillingData;
let dialog;
let message_from_parent;




Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    console.log("Office.onReady in Taskpane run");
    showTaskPane();

    // Assign event handlers and other initialization logic.
    document.getElementById("range_add_id").onclick = getAddress;
    document.getElementById("fetchBtn").onclick = fetchData;
    document.getElementById("writeBtn").onclick = writeData;
  }

  });

function setGlobal(var_name, value) {
  if (var_name in window) {
    window[var_name] = value;
    console.log(`setGlobal ${var_name} = ${window[var_name]}`);
  } else {
    throw new Error(`${var_name} has not been defined as a global variable`);
  }
}
  

async function showTaskPane() {
try {
    console.log("Line before Office.addin.showTaskPane()");
    await Office.addin.showAsTaskpane();
    console.log("Line after Office.addin.showTaskPane()");
} catch (error) {
    console.error("Error showing task pane: " + error);
    // Handle errors related to displaying the task pane here
}
}


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
    console.log(`Formed address of body page: ${htmlContent}`);
    
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
      console.log('Executed script:', scriptContent);
    }

    console.log("Loaded HTML content successfully");
  } catch (error) {
    console.error('Error loading HTML content:', error);
  }
}


function setMessage (message) {
    message_from_parent = message;
}

function openDialog() {
    const dialogUrl = 'https://localhost:3000/popup.html';
  
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog: " + asyncResult.error.message);
            return;
        }
  
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);
  
        
    });
  }
  
function processMessageFromDialog(arg) {
  if (arg.message === "dialogReady") {
    dialog.messageChild(message_from_parent);
  } else {
      console.log("arg message:" + arg.message);
  }
}

const button_to_form = {
  'SelectIntervalData': 'UserForm4TimeStampCols',
}


function SelectIntervalData() {
  
  setGlobal ("strNrmlzBillingData", "No");
  SelectData();
  
  return "SelectIntervalData";  
  
}
function SelectData() {

  // getglobal strmrlz
  // based on the value ex
  // ui.loadHtmlPage(name of the fragment);


  loadHtmlPage("UserForm4TimeStampCols");
  // loadHtmlPage("UserForm3InputDataRng");
  console.log("views.SelectData !!!");

  return "";  

}

async function OnAction_ECAM(event) {
  var function_name;

  console.log("Got to OnAction_ECAM");

  // Call function based on the button ID
  function_name = event.source['id'].replace(/^[a-z]+|\d+$/g, ''); //removes lower case prefix and numeric suffix
  
  //add process message from taskpane, add a listner to taskpane and then modify taskpane based on the button id
  //create this tutorial again https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator
  console.log(function_name);
  console.log(typeof window[function_name]);
  if (typeof window[function_name] === 'function') {
  let result = window[function_name]();
  setMessage("Button clicked for (" + result + ")");
  } else {
  setMessage("Button (" + function_name + ") not working yet!");
  }
  
  openDialog()


  event.completed();
}
  
Office.actions.associate("OnAction_ECAM", ribbon.OnAction_ECAM);

