/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */


const state = (() => {
    let _state = {
      strNrmlzBillingData: null,
      iTimeCols: null,

      // Add other global fields here
    };
  
    return {
      get: (key) => _state[key],
      set: (key, value) => {
        if (key in _state) {
          _state[key] = value;
          console.log(`${key} set to ${value}`);
        } else {
          console.warn(`Key ${key} is not defined in state`);
        }
      },
      getState: () => _state,
    };
  })();
  
  export default state;
  