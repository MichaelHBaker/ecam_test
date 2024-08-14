/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

function openDialog(content, isHtmlFile = false) {
    let dialogUrl = 'https://localhost:3000/popup.html';
    if (isHtmlFile) {
        dialogUrl += `?contentFile=${content}`;
    } else {
        dialogUrl += `?message=${encodeURIComponent(content)}`;
    }

    Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20 }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog: " + asyncResult.error.message);
            return;
        }
        const dialog = asyncResult.value;
        console.log("Dialog opened:", dialog);
        // If you need to handle messages back from the dialog, set up here
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            processMessageFromDialog(arg, dialog, content);
        });
    });
}

function openDialog_old(message) {
    const dialogUrl = 'https://localhost:3000/popup.html';
  
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to open dialog: " + asyncResult.error.message);
        return;
      }
  
      const dialog = asyncResult.value;
      console.log("Dialog opened:", dialog);
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => processMessageFromDialog(arg, dialog, message));
    });
  }
  
  function processMessageFromDialog(arg, dialog, message) {
    if (arg.message === "dialogReady") {
      console.log("Dialog is ready");
      // Wait for "eventHandlerReady"
    } else if (arg.message === "eventHandlerReady") {
      console.log("Event handler is ready, sending message to dialog");
      dialog.messageChild(message);
    } else {
      console.log("Received message from dialog:", arg.message);
      // Handle other messages if needed
    }
  }
  
  function receiveMessageFromParent(arg) {
    const message = arg.message;
    console.log(message);
    document.getElementById("messageText").innerText =  message;
    console.log("message assigned to innerText:" + document.getElementById("messageText").innerText);

    // Process the message...

};

const dialogs = {
    openDialog,
    receiveMessageFromParent
  };
  
export default dialogs;
  