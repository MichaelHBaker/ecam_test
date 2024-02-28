/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import * as ui from './ui';

export async function OnAction_ECAM(event) {
    var function_name;
  
    console.log("Got to OnAction_ECAM");
  
    // Call function based on the button ID
    function_name = event.source['id'].replace(/^[a-z]+|\d+$/g, ''); //removes lower case prefix and numeric suffix
  
    ui.setMessage("Button (" + function_name + ") not working yet!");
  
    //add process message from taskpane, add a listner to taskpane and then modify taskpane based on the button id
    //create this tutorial again https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator
    if (typeof window[function_name] === 'function') {
      ui.setMessage("Button clicked for (" + window[function_name]() + ")");
    } 
    
    ui.openDialog()
  
    try {
      // Code to show the task pane
      console.log("Line before showTaskPane()");
      await ui.showTaskPane();
      console.log("Line after showTaskPane");
  
      if (function_name === 'SelectIntervalData'){
        ui.loadHtmlPage('UserForm4TimeStampCols');
      }
  
      // Additional Excel.run can be placed here if needed
      // await Excel.run(async (context) => {
      //     // Asynchronous Excel operations here
      //     ...
      //     await context.sync();
      // });
    } catch (error) {
      // Handle any errors here
      console.error("Error: " + error);
    }
  
    event.completed();
  }
  