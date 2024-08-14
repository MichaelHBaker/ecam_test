/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

async function loadHtmlPage(pageName) {
    try {
      let response = await fetch(`/forms/${pageName}.html`);
      if (!response.ok) {
        throw new Error(`Failed to load the HTML page: ${response.statusText}`);
      }
      let htmlContent = await response.text();
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = htmlContent;
      const scripts = tempDiv.querySelectorAll('script');
      const contentFrame = document.getElementById('content-frame');
      contentFrame.innerHTML = tempDiv.innerHTML; // Includes innerHTML without <script> tags
      for (const script of scripts){
        const scriptElement = document.createElement('script');
        scriptElement.type = 'text/javascript';
        if (script.type === 'module') {
          scriptElement.type = 'module';
        }
        scriptElement.textContent = script.textContent;
        document.body.appendChild(scriptElement); // Append to body and execute
      }
      console.log("Loaded HTML content successfully");
    } catch (error) {
      console.error('Error loading HTML content:', error);
    }
  }

async function detectUnloadAction() {
  return new Promise((resolve) => {
    const submitButton = document.getElementById("submit_button_id");
    const cancelButton = document.getElementById("cancel_button_id");
    const backButton = document.getElementById("back_button_id");
    async function handleClick(event) {
      if (event.target === submitButton) {
        resolve('submit');
      } else if (event.target === cancelButton) {
        resolve('cancel');
      } else if (event.target === backButton) {
        resolve('back');
      }
    }
    // Special event handler for click on taskpane close
    Office.addin.onVisibilityModeChanged(function(args) {
      if (args.visibilityMode == "Hidden") {
        resolve('close');
      }
    });
    submitButton.addEventListener("click", handleClick);
    cancelButton.addEventListener("click", handleClick);
  });
}
const utils = {
  loadHtmlPage,
  detectUnloadAction
};
  
export default utils;