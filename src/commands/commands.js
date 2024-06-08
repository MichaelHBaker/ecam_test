/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */


Office.onReady((info) => {
  //info can be used to customize UI
  console.log("Office.onready in command.js");
  console.log(info.host.toString());
  console.log(info.platform.toString());

});


console.log("end of commands.js");