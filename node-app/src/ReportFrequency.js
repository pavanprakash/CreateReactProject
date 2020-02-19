import React, { useState } from "react";

function ReportFrequency() {
  let [showRadios, setShowRadios] = useState();
  function onChangeHandler(event) {
    console.log("on change handler fired...");
    console.log(event.target.value);
    if (event.target.value === "monthly") {
      setShowRadios(true);
    } else {
      setShowRadios(false);
    }
  }
  let radiosHtml = (
    <div>
      <label>BD2</label>
      <input type="radio" name="day" value="BD2" checked></input>
      <label>BD5</label>
      <input type="radio" name="day" value="BD5"></input>
    </div>
  );
  return (
    <div className="report-selection-container">
      <h4>Please select monthly or quarterly PDF Rec option</h4>

      <div onChange={onChangeHandler}>
        <select
          className="report-selection"
          name="frequency"
          id="option-select"
        >
          <option value="">--- Please choose an option  ---</option>
          <option value="monthly">Monthly</option>
          <option value="quarterly">Quarterly</option>
        </select>
      </div>
      <div className="report-selection-radios">
        {showRadios ? radiosHtml : null}
      </div>
    </div>
  );
}

export default ReportFrequency;
