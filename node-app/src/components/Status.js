import React, { useEffect, useState } from "react";
import config from "../config/index";
const axios = require("axios");

let host = config().hostName;
let port = config().serverPort;

function Status(props) {
  useEffect(() => {
    axios.post(`http://${host}:${port}/check`, { type: "monthly" });
    axios.post(`http://${host}:${port}/check`, { type: "quarterly" });
  });

  return (
    <div className="status-container">
      <div className="item">
        <h3>Report execution status (running / not running)</h3>
        <h4> MONTHLY PDFREC Status is {props.monthlyStatus} </h4>
        <h4> QUARTERLY PDFREC Status is {props.quarterlyStatus} </h4>
      </div>
    </div>
  );
}

export default Status;
