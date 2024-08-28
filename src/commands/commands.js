/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

// // import functionMap from '../taskpane/taskpane.js';
// import dialogs from '../dialogs/dialogs.js';
// import functionMap from '../dialogs/popup.js';

// Office.onReady((info) => {
//   //info can be used to customize UI
//   console.log("Office.onready in command.js");
//   console.log(info.host.toString());
//   console.log(info.platform.toString());

// });

// function OnAction_ECAM(event) {
//   const buttonId = event.source.id;
//   const functionName = buttonId.replace(/^[a-z]+|\d+$/g, '');

//   console.log("Got to OnAction_ECAM");
//   console.log(functionName);

//   const functionToCall = functionMap[functionName];

//   let message;
//   try {
//     if (typeof functionToCall !== 'function') {
//       message = `Button (${functionName}) not working yet!`;
//     } else {
//       functionToCall();
//     }
//   } catch (error) {
//     console.error("Error in OnAction_ECAM:", error);
//     message = `Error: ${error.message}`;
//   }

//   if (message) {
//     dialogs.openDialog(message);
//   }

//   event.completed();
// }

// Office.actions.associate("OnAction_ECAM", OnAction_ECAM);

// // function openDialog(message) {
// //   const dialogUrl = 'https://localhost:3000/popup.html';

// //   Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
// //     if (asyncResult.status === Office.AsyncResultStatus.Failed) {
// //       console.error("Failed to open dialog: " + asyncResult.error.message);
// //       return;
// //     }

// //     const dialog = asyncResult.value;
// //     console.log("Dialog opened:", dialog);
// //     dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => processMessageFromDialog(arg, dialog, message));
// //   });
// // }

// // function processMessageFromDialog(arg, dialog, message) {
// //   if (arg.message === "dialogReady") {
// //     console.log("Dialog is ready");
// //     // Wait for "eventHandlerReady"
// //   } else if (arg.message === "eventHandlerReady") {
// //     console.log("Event handler is ready, sending message to dialog");
// //     dialog.messageChild(message);
// //   } else {
// //     console.log("Received message from dialog:", arg.message);
// //     // Handle other messages if needed
// //   }
// // }



// // const commands = {
// //   openDialog,
// //   processMessageFromDialog
// // };

// // export default commands;

// console.log("end of commands.js");

//Claude code
import dialogs from '../dialogs/dialogs.js';
import functionMap from '../dialogs/popup.js';

let isExcelApiReady = false;

Office.onReady((info) => {
  console.log("Office.onready in command.js");
  console.log(info.host.toString());
  console.log(info.platform.toString());

  if (info.host === Office.HostType.Excel) {
    // Load the Excel API
    Excel.run(async (context) => {
      // This ensures that the Excel API is fully loaded
      await context.sync();
      isExcelApiReady = true;
      console.log("Excel API is ready");
      setupSelectionChangeHandler();
    }).catch(function(error) {
      console.log("Error loading Excel API: " + error);
    });
  }
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

function setupSelectionChangeHandler() {
  if (!isExcelApiReady) {
    console.log("Excel API not ready yet. Delaying setup of selection change handler.");
    setTimeout(setupSelectionChangeHandler, 100); // Retry after 100ms
    return;
  }

  Excel.run(function(context) {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onSelectionChanged.add(handleSelectionChange);
    return context.sync();
  }).catch(function(error) {
    console.log("Error setting up selection change handler: " + error);
  });
}

function handleSelectionChange(event) {
  Excel.run(function(context) {
    let range = context.workbook.getSelectedRange();
    range.load("address");
    return context.sync().then(function() {
      dialogs.sendMessageToDialog(range.address);
    });
  }).catch(function(error) {
    console.log("Error handling selection change: " + error);
  });
}

console.log("end of commands.js");