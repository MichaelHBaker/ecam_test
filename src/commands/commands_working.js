/* eslint-disable prettier/prettier */
/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
async function toggleProtection(args) {
  try {
      await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        sheet.load('protection/protected');
        await context.sync();
        
        if (sheet.protection.protected) {
            sheet.protection.unprotect();
        } else {
            sheet.protection.protect();
        }
          await context.sync();
      });
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }

  args.completed();
}
Office.actions.associate("toggleProtection", toggleProtection);

// The command function.
async function highlightSelection(event) {
  console.log('log test');
  // Implement your custom code here. The following code is a simple Excel example.  
  try {
      await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = "yellow";
          await context.sync();
      });
  } catch (error) {
      // Note: In a production add-in, notify the user through your add-in's UI.
      console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

// You must register the function with the following line.
Office.actions.associate("highlightSelection", highlightSelection);

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

// Probably will need to async and have try/catch logic when further developed
function OnAction_ECAM(args) {
  // Your function logic here
  
  // Display a message after executing the function
  window.alert('Function executed successfully!');
  
  args.completed();
}

