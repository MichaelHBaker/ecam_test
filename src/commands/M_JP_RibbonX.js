/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import * as ui from './ui';

const button_to_form = {
  'SelectIntervalData': 'UserForm4TimeStampCols',
}

export async function OnAction_ECAM(event) {
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
      ui.setMessage("Button clicked for (" + result + ")");
    } else {
      ui.setMessage("Button (" + function_name + ") not working yet!");
    }
    
    ui.openDialog()
  
    try {
      // Code to show the task pane
      console.log("Line before showTaskPane()");
      await ui.showTaskPane();
      console.log("Line after showTaskPane");
  
      ui.loadHtmlPage(button_to_form[function_name]);
  
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
  
