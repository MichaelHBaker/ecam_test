/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
Office.onReady(() => {
    // Office is ready
    setupInputBox();
});

function setupInputBox() {
    const inputBox = document.getElementById('columnInput');
    const instructions = document.getElementById('instructions');

    // Show instructions when the input box is clicked
    inputBox.addEventListener('focus', () => {
        instructions.style.display = 'block';
    });

    // Hide instructions and get column address when Enter is pressed
    inputBox.addEventListener('keydown', function(event) {
        if (event.key === 'Enter') {
            instructions.style.display = 'none';
            getCurrentColumnAddress();
        }
    });
}

function getCurrentColumnAddress() {
    Excel.run(async (context) => {
        // Get the currently selected range
        const range = context.workbook.getSelectedRange();

        // Load the address
        range.load('address');

        await context.sync();

        // Update the text box with the address
        document.getElementById('columnInput').value = range.address;
    }).catch(error => {
        console.error('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.error('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}
