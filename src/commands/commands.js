/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

let dialog;
let button_id;

Office.onReady((info) => {
  //info can be used to customize UI
  console.log(info.host.toString());
  console.log(info.platform.toString());
});

async function OnAction_ECAM(event) {

  // if (typeof me.onChange === "function") { 
  //   // safe to use the function
  // }

  button_id = event.source['id'];
  
  openDialog();

  event.completed();
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
    dialog.messageChild(button_id);
  } else {
      console.log("arg message:" + arg.message);
  }
}


// Other functions and logic...


