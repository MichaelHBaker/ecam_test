/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

import * as ribbon from './M_JP_RibbonX.js';
//import named functions from specific legacy like folders that are one to one with button clicks


Office.onReady((info) => {
  //info can be used to customize UI
  console.log(info.host.toString());
  console.log(info.platform.toString());

});


function SelectIntervalData() {
  return "SelectIntervalData";  
}


// Associate the function with Office actions
Office.actions.associate("OnAction_ECAM", ribbon.OnAction_ECAM);

