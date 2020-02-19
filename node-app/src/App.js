import React from "react";
import logo from "./logo.svg";
import "./App.css";
import Status from "./components/Status";
import Result from "./components/Result";
import Submit from "./components/Submit";
import ReportFrequency from "./ReportFrequency";
import { useState } from "react";
import config from "./config/index";

let websocketHost = config().websocketHost;
var ws = new WebSocket(websocketHost);

console.log("websocketHost: ", websocketHost);
ws.binaryType = "arraybuffer";
let isMonthlyRunning = false;
let isQuarterlyRunning = false;
let monthlyStatus;
let quarterlyStatus;
let buttonDisabled;
function App() {
  let [MONTHLY_STATUS, setMonthlyStatus] = useState("CHECKING...");
  let [QUARTERLY_STATUS, setQuarterlyStatus] = useState("CHECKING...");
  let [MONTHLY_RES_STATUS, setMonthlyResStatus] = useState("NOT RUNNING...");
  let [QUARTERLY_RES_STATUS, setQuarterlyResStatus] = useState(
    "NOT RUNNING..."
  );

  ws.onmessage = function(ev) {
    let socketMessage = JSON.parse(ev.data);
    console.log("socket message inside app.js", socketMessage);

    if (socketMessage.category === "status") {
      if (socketMessage.type === "monthly") {
        if (socketMessage.value === "RUNNING...") {
          isMonthlyRunning = true;
          monthlyStatus = socketMessage.value;
        } else {
          isMonthlyRunning = false;
          // console.log(
          //   `isMonthlyRunning monthly inside else: ${isMonthlyRunning}`
          // );
        }
        setMonthlyStatus(socketMessage.value);
      } else if (socketMessage.type === "quarterly") {
        if (socketMessage.value === "RUNNING...") {
          isQuarterlyRunning = true;
          quarterlyStatus = socketMessage.value;
        } else {
          isQuarterlyRunning = false;
          // console.log(
          //   `isQuarterlyRunning qtly inside else: ${isQuarterlyRunning}`
          // );
        }
        setQuarterlyStatus(socketMessage.value);
      }
    } else if (socketMessage.category === "result") {
      if (isMonthlyRunning) {
        // console.log(
        //   `Result: socketMessage.value  MONTHLY=RUNNING : ${socketMessage.value}`
        // );
        setMonthlyResStatus(monthlyStatus);
      } else if (isQuarterlyRunning) {
        // console.log(
        //   `Result: socketMessage.value   Qtrly=RUNNING : ${socketMessage.value}`
        // );
        setQuarterlyResStatus(quarterlyStatus);
      } else {
        if (socketMessage.type === "monthly") {
          // console.log(
          //   `Result: socketMessage.value  inside MONTHLY : ${socketMessage.value}`
          // );
          setMonthlyResStatus(socketMessage.value);
        } else if (socketMessage.type === "quarterly") {
          // console.log(
          //   `Result: socketMessage.value  inside Qtrly : ${socketMessage.value}`
          // );
          setQuarterlyResStatus(socketMessage.value);
        }
      }
    }
  };
  buttonDisabled = isMonthlyRunning || isQuarterlyRunning;

  return (
    <div className="App">
      <div className="main-container">
        {/* <div className="dialogContainer">
          <p className="dialog">Hello.............</p>
        </div> */}
        <h1 class="header">PDF Reconcilliation Process</h1>
        <Status
          monthlyStatus={MONTHLY_STATUS}
          quarterlyStatus={QUARTERLY_STATUS}
        />
        <Result
          monthlyResStatus={MONTHLY_RES_STATUS}
          quarterlyResStatus={QUARTERLY_RES_STATUS}
        />
        <form action="/submit" method="post">
          <div className="file-container">
            <h3>Enter symphony and RDB file location</h3>
            <div className="file-item">
              <label className="file-item-label">
                Symphony Files Location:
              </label>
              <input
                className="file-item-input"
                type="text"
                name="symphony"
              ></input>
            </div>
            <div className="file-item">
              <label className="file-item-label">RDB Files Location: </label>
              <input className="file-item-input" type="text" name="rdb"></input>
            </div>
          </div>
          <div class="pdf-rec-type">
            <ReportFrequency />
            <div className="distribution-selection-container">
              <h4>Please select distribute or archive PDF Rec option</h4>
              <div className="distribution-selection">
                <label>Distribute</label>
                <input
                  type="radio"
                  name="runType"
                  value="distribute"
                  checked
                ></input>
                <label>Archive</label>
                <input type="radio" name="runType" value="archive"></input>
              </div>
            </div>
          </div>
          <Submit btnValue={buttonDisabled} />
        </form>
      </div>
    </div>
  );
}

export default App;
