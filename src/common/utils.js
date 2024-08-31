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

async function loadRangeAddressHandler(){
  console.log("Start of loadRangeAddressHandle:");
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();     
    worksheet.onSelectionChanged.add(rangeSelectionHandler);
    await context.sync();
  }); 
  console.log("End of loadRangeAddressHandle:");
}

async function rangeSelectionHandler(event){
await Excel.run(async (context) => {

  let range = context.workbook.getSelectedRange();
  range.load("address");
  await context.sync();
  document.getElementById("range_address_id").value = range.address;
  document.getElementById("submit_button_id").disabled = false;
  

  console.log(`The address of the selected range is "${range.address}"`);

});
}



//from stackoverflow - converted to javascript by chatgpt
const promptForRangeBindingId = () => {
  return new Promise((resolve, reject) => {
      const handleResult = (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value.id);
          } else {
              reject(result.error.message);
          }
      };
      console.log("1 promptForRangeBindingId: before bindings: " + Office);
      console.log("2 promptForRangeBindingId: before bindings: " + Office.context);
      console.log("3 promptForRangeBindingId: before bindings: " + Office.context.document);
      console.log("4 promptForRangeBindingId: before bindings: " + Office.context.document.bindings);
      Office.context.document.bindings.addFromPromptAsync(
        Office.BindingType.Matrix,
        handleResult
      );
      console.log("promptForRangeBindingId: after bindings");
  });
};

const getAddressesByBindingId = (bindingId) => {
  return new Promise((resolve, reject) => {
      Excel.run((context) => {
          const binding = context.workbook.bindings.getItem(bindingId);
          const range = binding.getRange();
          range.load('address');
          context.sync();
          return resolve(range.address);
          
      }).catch((error) => {
          reject(error);
      }).finally(() => {
          Office.context.document.bindings.releaseByIdAsync(bindingId);
      });
  });
};

const promptForAddressRange = async () => {
  try {
      const bindingId = await promptForRangeBindingId();
      console.log("After promptForRangeBindingId: " + bindingId);
      const address = await getAddressesByBindingId(bindingId);
      return address;
  } catch (error) {
      console.log("Error from promptForAddressRange: " + error.message);
      // throw new Error(error.message);
  }
};

// export {
//   promptForAddressRange,
//   promptForRangeBindingId,
//   getAddressesByBindingId,
// };

const utils = {
  loadHtmlPage,
  detectUnloadAction,
  loadRangeAddressHandler,
  rangeSelectionHandler,
  promptForAddressRange,
  promptForRangeBindingId,
  getAddressesByBindingId,
};
  
export default utils;