/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

let dialog;
let message_from_parent;

import functionMap from '../taskpane/taskpane.js';

Office.onReady((info) => {
  //info can be used to customize UI
  console.log("Office.onready in command.js");
  console.log(info.host.toString());
  console.log(info.platform.toString());

});

async function OnAction_ECAM(event) {
  const buttonId = event.source.id;
  const functionName = buttonId.replace(/^[a-z]+|\d+$/g, ''); // removes lower case prefix and numeric suffix

  console.log("Got to OnAction_ECAM");
  console.log(functionName);


  const functionToCall = functionMap[functionName];
  if (typeof functionToCall === 'function') {
    let result = functionToCall();
    setMessage("Button clicked for (" + result + ")");
  } else {
    setMessage("Button (" + functionName + ") not working yet!");
  }

  openDialog();

  event.completed();
}

Office.actions.associate("OnAction_ECAM", OnAction_ECAM);

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



console.log("end of commands.js");