/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

let dialog;
let message_from_parent;

Office.onReady((info) => {
  //info can be used to customize UI
  console.log(info.host.toString());
  console.log(info.platform.toString());
});

async function OnAction_ECAM(event) {
  var function_name;

  // Dynamically call the function based on the button's ID
  function_name = event.source['id'].replace(/^[a-z]+|\d+$/g, ''); //removes lower case prefix and numeric suffix

  if (typeof window[function_name] === 'function') {
      window[buttonId](event);
  } else {
      console.error('No function found for button ID:', buttonId);
      event.completed();
  }

  message_from_parent = "Button (" + function_name + ") not working yet!";
  
  openDialog();

  event.completed();
}

function IntervalData() {
  return "IntervalData";  
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


// Other functions and logic...


