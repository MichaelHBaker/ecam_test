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


// The command function.
async function highlightSelection(event) {
  // Implement your custom code here. The following code is a simple Excel example.  
  try {
      await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = "yellow";
          range.values = "ok"
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



