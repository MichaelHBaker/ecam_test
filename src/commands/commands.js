/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import functionMap from '../taskpane/taskpane.js';
import dialogs from '../dialogs/dialogs.js';

Office.onReady((info) => {
  //info can be used to customize UI
  console.log("Office.onready in command.js");
  console.log(info.host.toString());
  console.log(info.platform.toString());

});

function OnAction_ECAM(event) {
  const buttonId = event.source.id;
  const functionName = buttonId.replace(/^[a-z]+|\d+$/g, '');

  console.log("Got to OnAction_ECAM");
  console.log(functionName);

  const functionToCall = functionMap[functionName];

  let message;
  try {
    if (typeof functionToCall !== 'function') {
      message = `Button (${functionName}) not working yet!`;
    } else {
      functionToCall();
    }
  } catch (error) {
    console.error("Error in OnAction_ECAM:", error);
    message = `Error: ${error.message}`;
  }

  if (message) {
    dialogs.openDialog(message);
  }

  event.completed();
}

Office.actions.associate("OnAction_ECAM", OnAction_ECAM);

// function openDialog(message) {
//   const dialogUrl = 'https://localhost:3000/popup.html';

//   Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
//     if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//       console.error("Failed to open dialog: " + asyncResult.error.message);
//       return;
//     }

//     const dialog = asyncResult.value;
//     console.log("Dialog opened:", dialog);
//     dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => processMessageFromDialog(arg, dialog, message));
//   });
// }

// function processMessageFromDialog(arg, dialog, message) {
//   if (arg.message === "dialogReady") {
//     console.log("Dialog is ready");
//     // Wait for "eventHandlerReady"
//   } else if (arg.message === "eventHandlerReady") {
//     console.log("Event handler is ready, sending message to dialog");
//     dialog.messageChild(message);
//   } else {
//     console.log("Received message from dialog:", arg.message);
//     // Handle other messages if needed
//   }
// }



// const commands = {
//   openDialog,
//   processMessageFromDialog
// };

// export default commands;

console.log("end of commands.js");