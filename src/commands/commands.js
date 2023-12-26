/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

// The command function.
async function OnAction_ECAM(event) {
  OpenDialog('ID of button clicked is:' + event.source['id']);
  event.completed();
}


function OpenDialog(message) {
  const dialogUrl = 'https://localhost:3000/popup.html';
  console.log(message);

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20 },
  (asyncResult) => {
      const dialog = asyncResult.value;

      // Send the element ID to the dialog
      const messageToSend = JSON.stringify({ message: message });
      console.log("Sending message to dialog:", messageToSend);
      dialog.messageChild(messageToSend);
    }
  );
  }

