const fs = require("fs");

function delay(fileName, timeMs) {
  let fileDeleted = false;
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      console.log("inside set timeout...");
      if (!fs.existsSync(fileName)) {
        console.log(`file: ${fileName} is deleted successfully!`);
        fileDeleted = true;
      }
      resolve(fileDeleted);
    }, timeMs);
  });
}

async function checkForLockFile(fileName, timeoutMs, iterations) {
  // returns true or false
  // true -- waited and file deletion is confirmed
  // false -- timeout, file still exits after timeout
  let fileDeleted = false;

  console.log("inside funct");
  while (iterations > 0) {
    console.log("looop...");
    fileDeleted = await delay(fileName, timeoutMs);
    if (fileDeleted) {
      break;
    } else {
      console.log("checking delay function- waiting for file to be deleted");
    }
    iterations = iterations - 1;
    console.log("iterations");
  }
  console.log(`file deleted value: ${fileDeleted}`);
  return fileDeleted;
}
module.exports = checkForLockFile;
