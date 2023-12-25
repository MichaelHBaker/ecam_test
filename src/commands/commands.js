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
          OpenDialog("This is the message to display in the dialog box.");
      });
  } catch (error) {
      console.error(error);
  }

  event.completed();
}


function OpenDialog(message) {
  const dialogUrl = 'https://localhost:3000/popup.html'; 

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20 },
  (asyncResult) => {
      const dialog = asyncResult.value;
      });
  }

