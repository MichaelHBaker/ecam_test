/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import * as ribbon from './M_JP_RibbonX.js';
//import named functions from specific legacy like folders that are one to one with button clicks


Office.onReady((info) => {
  //info can be used to customize UI
  console.log("Office.onready in command.js");
  console.log(info.host.toString());
  console.log(info.platform.toString());

});

// function SelectIntervalData() {
//   return "SelectIntervalData";  
// }




// Associate the function with Office actions
Office.actions.associate("OnAction_ECAM", ribbon.OnAction_ECAM);

let dialog;
let message_from_parent;

export function setMessage (message) {
    message_from_parent = message;
}

export function openDialog() {
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



console.log("end of commands.js");