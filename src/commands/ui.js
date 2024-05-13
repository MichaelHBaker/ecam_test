/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

let dialog;
let message_from_parent;

export function setMessage (message) {
    message_from_parent = message;
}

export function openDialog() {
    const dialogUrl = 'https://localhost:3000/popup.html';
  
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog: " + asyncResult.error.message);
            return;
        }
  
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);
  
        
    });
  }
  
function processMessageFromDialog(arg) {
  if (arg.message === "dialogReady") {
    dialog.messageChild(message_from_parent);
  } else {
      console.log("arg message:" + arg.message);
  }
}

export async function showTaskPane() {
  try {
      console.log("Line before Office.addin.showTaskPane()");
      await Office.addin.showAsTaskpane();
      console.log("Line after Office.addin.showTaskPane()");
  } catch (error) {
      console.error("Error showing task pane: " + error);
      // Handle errors related to displaying the task pane here
  }
}

// export async function loadHtmlPage(pageName) {
//     document.getElementById('content-frame').src = pageName + '.html';
//     console.log("inside loadhtml");
//     let htmlFile = await fetch("/forms/" + pageName + ".html");
//     // let htmlFile = await fetch(pageName + ".html");
//     let htmlSrc = await htmlFile.text();
//     console.log(`"formed address of body page" ${htmlSrc}`);
//     document.getElementById('content-frame').innerHTML = htmlSrc;
//   }

export async function loadHtmlPage(pageName) {
  try {
    // Fetch the HTML content
    let response = await fetch(`/forms/${pageName}.html`);
    if (!response.ok) {
      throw new Error(`Failed to load the HTML page: ${response.statusText}`);
    }

    let htmlContent = await response.text();
    console.log(`Formed address of body page: ${htmlContent}`);
    
    // Create a temporary container to parse the fetched HTML
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlContent;

    // Extract script tags and their content
    const scripts = tempDiv.getElementsByTagName('script');
    const scriptContents = [];
    for (const script of scripts) {
      scriptContents.push(script.innerText);
      script.parentNode.removeChild(script); // Remove script tag from tempDiv
    }

    // Clear existing content and event listeners
    const contentFrame = document.getElementById('content-frame');
    const newContentFrame = contentFrame.cloneNode(false); // Clone without children to remove event listeners
    contentFrame.parentNode.replaceChild(newContentFrame, contentFrame);

    // Replace the body's content with the fetched HTML content (without script tags)
    newContentFrame.innerHTML = tempDiv.innerHTML;

    // Dynamically create and append script elements to the body
    for (const scriptContent of scriptContents) {
      const scriptElement = document.createElement('script');
      scriptElement.type = 'text/javascript';
      scriptElement.text = scriptContent;
      document.body.appendChild(scriptElement); // Append to body to execute the script
      console.log('Executed script:', scriptContent);
    }

    console.log("Loaded HTML content successfully");
  } catch (error) {
    console.error('Error loading HTML content:', error);
  }
}

  
  