// open terminal and run with this command
// node testFunction.js

function removePrefixSuffix(str) {
    const regex = /^[a-z]+|\d+$/g;
    return str.replace(regex, '');
}

console.log(removePrefixSuffix("abcDef123")); // Expected output: "Def"
console.log(removePrefixSuffix("xyzHelloWorld4567")); // Expected output: "HelloWorld"
console.log(removePrefixSuffix("xyzHello1234World4567")); // Expected output: "HelloWorld"
// Add more test cases if necessary