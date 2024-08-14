/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import dialogs from '../dialogs/dialogs.js';

Office.onReady(() => {
    console.log("Office.onReady from popup.js");

    // Notify that the dialog is ready
    Office.context.ui.messageParent("dialogReady");

    // Set up the event handler for receiving messages
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, dialogs.receiveMessageFromParent, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Event handler was successfully added.");
            // Only now do we notify that the event handler is ready
            Office.context.ui.messageParent("eventHandlerReady");
        } else {
            console.error("Failed to set event handler: " + result.error.message);
        }
    });
});



  
 

// Office.onReady(() => {
//     console.log("office.onready from popup.js");
//     Office.context.ui.messageParent("dialogReady");
//     Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveMessageFromParent);
// })

// function receiveMessageFromParent(arg) {
//     const message = arg.message;
//     console.log(message);
//     document.getElementById("messageText").innerText =  message;

//     // Process the message...

// };
