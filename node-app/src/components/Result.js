import React, { useState, useEffect } from "react";
import config from "../config/index";

let host = config().hostName;
let port = config().serverPort;
const axios = require("axios");

function Result(props) {
  useEffect(() => {
    axios.post(`http://${host}:${port}/checkResultStatus`, {
      option: "monthly"
    });
    axios.post(`http://${host}:${port}/checkResultStatus`, {
      option: "quarterly"
    });
  });
  return (
    <div className="result-container">
      <div className="item">
        <h3>Report Result Status (Running/Pass/Fail)</h3>
        <h4> MONTHLY PDFREC Result Status is {props.monthlyResStatus} </h4>
        <h4> QUARTERLY PDFREC Result Status is {props.quarterlyResStatus} </h4>
      </div>
    </div>
  );
}

export default Result;
