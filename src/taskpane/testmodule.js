import * as test from "./testdependency";
export function helloworld() {
    console.log("hello from test modulue.js");
    console.log(test.multiply2numbers(3,5));
}