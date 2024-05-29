/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import { loadHtmlPage } from "./ui";

export function SelectIntervalData() {
    
    setGlobal ("strNrmlzBillingData", "No");
    SelectData();
    
    return "SelectIntervalData";  
    
}
function SelectData() {

    // getglobal strmrlz
    // based on the value ex
    // ui.loadHtmlPage(name of the fragment);
    loadHtmlPage('UserForm3InputDataRng');
    console.log("views.SelectData !!!");
  
    return "";  
  
}

