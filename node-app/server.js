const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const app = express();
const shell = require("shelljs");
const fs = require("fs");
let cors = require("cors");
const waitForFileToBeDeleted = require("./src/checkForLockFile");
const config = require("./src/config/index");
const webSocketPort = config().websocketPort;
const serverPort = config().serverPort;

const projFolderPath = config().projFolderPath;
const powershellPath = `${projFolderPath}\\ps\\`;

const pathPsScript = path.resolve(projFolderPath + "dotnet-app\\publish");
console.log("pathPsScript: ", pathPsScript);
let WebSocketServer = require("ws").Server,
  wss = new WebSocketServer({ port: webSocketPort });

console.log("webSocketPort ", webSocketPort);
console.log("serverPort ", serverPort);
let wsHandle;
const dbFunctions = require("./src/mongoDB/lib/dbFunctions");
const dbDetails = require("./src/mongoDB/dbDetails");

app.use(bodyParser());
app.use(cors());
app.use(express.static(path.join(__dirname, "build")));

app.use(bodyParser.urlencoded({ extended: false }));

//establish socket connection

wss.on("connection", function(ws) {
  console.log("Connectedd!");
  wsHandle = ws;
});

app.get("/home", (req, res) => {
  res.sendFile(path.join(__dirname, "build", "index.html"));
});
app.post("/checkResultStatus", async (req, res) => {
  let option = req.body.option;
  console.log(`check the status of the run for ${option}`);
  let records = await dbFunctions.getRecords("PdfRec", "PdfRecResults", {
    option: option
  });

  let message = {
    category: "result",
    type: option
  };
  let status;
  if (records == null) {
    status = "NOT AVAILABLE";
  } else {
    status = records[0].status;
  }
  message.value = status;

  console.log(message);
  wsHandle.send(JSON.stringify(message), { binary: false });
  res.redirect("/");
});
app.post("/check", (req, res) => {
  console.log("checking if monthly/quarterly powershell script is running");

  let type = req.body.type;

  let checkLock = shell.exec(
    `powershell ${powershellPath}checkPowershellRunStatus.ps1 -folder ${type}`
  );
  console.log("checklock: ", checkLock.stdout.toString());
  let message = {
    category: "status",
    type,
    value: checkLock.stdout.toString().trim()
  };

  console.log(`socket message for ${type} : ${message}`);
  wsHandle.send(JSON.stringify(message), { binary: false });
  res.redirect("/");
});
app.post("/submit", (req, res) => {
  let symFileLocPart2 = req.body.symphony;
  let rdbFileLocPart2 = req.body.rdb;
  let runType = req.body.runType;
  let option = req.body.frequency;
  let monthlyDayPart = req.body.day;
  let termType;
  let fileName;
  let fileDeleted;
  console.log(`symFileLocPart2: ${symFileLocPart2}`);
  console.log(`rdbFileLocPart2: ${rdbFileLocPart2}`);
  let message = {
    category: "result",
    type: option,
    value: "RUNNING..."
  };

  if (option === "monthly") {
    termType = "monthly";
    fileName = "\\\\ruffer.local\\dfs\\Shared\\PDFRec\\monthlyPdfLock.txt";
  } else if (option === "quarterly") {
    termType = "quarterly";
    fileName = "\\\\ruffer.local\\dfs\\Shared\\PDFRec\\quarterlyPdfLock.txt";
  }

  const asyncExec = new Promise(async (resolve, reject) => {
    if (termType === "monthly") {
      result = shell.exec(
        `powershell ${powershellPath}monthly-pdf-rec.ps1 -projFolder '${pathPsScript}' -symFileLoc '${symFileLocPart2}' -rdbFileLoc '${rdbFileLocPart2}' -runType ${runType} -day ${monthlyDayPart}`,
        { async: true }
      );
    } else if (termType === "quarterly") {
      console.log("running the powershell script for quarterly");
      result = shell.exec(
        `powershell ${powershellPath}quarterly-pdf-rec.ps1 -projFolder '${pathPsScript}' -symFileLoc '${symFileLocPart2}' -rdbFileLoc '${rdbFileLocPart2}' -runType ${runType}`,
        { async: true }
      );
    }
    console.log("printing result *******");

    fileDeleted = await waitForFileToBeDeleted(fileName, 5000, 100);
    console.log(`within server, filedeleted: ${fileDeleted}`);
    resolve(fileDeleted);
  });

  asyncExec.then(async () => {
    console.log(`inside asyncExec`);
    //write results to the DB
    console.log("writing results to mongo Db");

    if (result !== null) {
      //resolve the promise

      console.log(`inside result loop`);
      //enter code here to determine when the powershell prog has finished running... socket connection
      let message1 = {
        category: "result",
        type: option
      };
      if (fileDeleted) {
        console.log(`inside if, filedeleted value: ${fileDeleted}`);
        message1.value = "COMPLETED";
        console.log(message1);
        wsHandle.send(JSON.stringify(message1), { binary: false });
        //send run status
        let message = {
          category: "status",
          type: option,
          value: "NOT RUNNING..."
        };

        wsHandle.send(JSON.stringify(message), { binary: false });
      } else {
        console.log(`inside else, filedeleted value: ${fileDeleted}`);
        if (fs.existsSync(fileName)) {
          console.log("inside else");
          fs.unlinkSync(fileName);
        }

        message1.value = "FAILED";
        console.log(message1);
        let message = {
          category: "status",
          type: option,
          value: "NOT RUNNING..."
        };

        wsHandle.send(JSON.stringify(message), { binary: false });
      }
      await dbFunctions.insertRecords(
        // dbDetails.dbName,
        // dbDetails.dbCollection,
        "PdfRec",
        "PdfRecResults",
        {
          status: message1.value,
          option: option,
          symFileLocation: symFileLocPart2,
          rdbFileLocation: rdbFileLocPart2,
          monthlyDayPart: monthlyDayPart,
          result: result
        }
      );
    } else {
      console.log(`inside else loop`);
      let message = {
        category: "status",
        type: option,
        value: "NOT RUNNING..."
      };

      wsHandle.send(JSON.stringify(message), { binary: false });
      await dbFunctions.insertRecords(
        // dbDetails.dbName,
        // dbDetails.dbCollection,
        "PdfRec",
        "PdfRecResults",
        {
          status: "FAILED",
          option: option,
          symFileLocation: symFileLocPart2,
          rdbFileLocation: rdbFileLocPart2,
          monthlyDayPart: monthlyDayPart,
          result: result
        }
      );
    }
  });

  res.redirect("/");
});
app.listen(process.env.PORT || serverPort, () => {
  console.log("process.env.PORT ", process.env.PORT);
  console.log("serverPort ", serverPort);
  console.log("Node server is running");
});
