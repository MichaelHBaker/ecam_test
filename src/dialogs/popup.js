/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(() => {
   
    Office.context.ui.messageParent("dialogReady");
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveMessageFromParent);
})

function receiveMessageFromParent(arg) {
    const message = arg.message;
    console.log("Message received from parent:", message);
    document.getElementById("messageText").innerText = "Button clicked: " + message;
    // Process the message...

};
