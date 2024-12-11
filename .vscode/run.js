const readline = require('readline');
const fs = require('fs');

const filePath = process.argv[2];
const lineNumber = process.argv[3];

const line = readline.createInterface({
  input: fs.createReadStream(filePath),
  output: process.stdout,
  terminal: false,
});

let lineCount = 0;
let allTestLines = [];

line.on('line', (line) => {
  // Start of the line
  lineCount++;

  if (line.match(/test\(\"(.*?)\", async/g)) {
    allTestLines.push(lineCount);
  } else if (line.match(/test\(\'(.*?)\', async/g)) {
    allTestLines.push(lineCount);
  }
});

line.on('close', () => {
  if (allTestLines.length == 0) {
    console.error('This file is not supported for playwright testing');
    return;
  }

  let lineToRun = allTestLines.reverse().find((l) => Number(lineNumber) >= l);
  if (lineToRun == undefined) {
    lineToRun = allTestLines[0];
  }

  console.log(lineToRun);
});
