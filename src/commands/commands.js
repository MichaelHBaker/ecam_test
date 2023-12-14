/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

// The command function.
async function OnAction_ECAM(event) {

  try {
    await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.values = event.source['id'];

          await context.sync();
          console.log(event.source['id']);
          console.log("hello world");

          openDialog(event.source['id']);
      });
  } catch (error) {
      // Note: In a production add-in, notify the user through your add-in's UI.
      console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

// You must register the function with the following line.
Office.actions.associate("OnAction_ECAM", OnAction_ECAM);

function openDialog(message) {
  // URL of your dialog HTML page
  // const dialogUrl = 'https://localhost:3000/messageDialog.html'; 
  const dialogUrl = 'https://localhost:3000/popup.html'; 
  Office.context.ui.displayDialogAsync(dialogUrl, { width: 20, height: 40 }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          var dialog = asyncResult.value;
          // dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);

          // You can also send an initial message to your dialog here, if needed
          // dialog.messageChild({ type: "initialMessage", value: "Hello Dialog!" });
          dialog.messageChild("Hello Dialog!");
          console.log("message sent");
      } else {
          console.error("Failed to open dialog: " + asyncResult.error.message);
      }
  });
}

function processMessageFromDialog(arg) {
  console.log("Message from dialog: " + arg.message); // Handle messages sent from the dialog
}
