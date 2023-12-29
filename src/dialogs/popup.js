/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(() => {
   
    Office.context.ui.messageParent("dialogReady");
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveMessageFromParent);
})

function receiveMessageFromParent(arg) {
    const message = arg.message;
    document.getElementById("messageText").innerText =  message;

    // Process the message...

};
