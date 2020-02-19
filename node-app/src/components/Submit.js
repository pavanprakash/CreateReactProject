import React from "react";

function Submit(props) {
  let buttonHTML;
  console.log(`props.buttonDisabled: ${props.btnValue}`);

  let confirmBox = (
    <div id="confirmBox">
      <div class="message"></div>
      <button class="yes">Yes</button>
      <button class="no">No</button>
    </div>
  );

  if (props.btnValue) {
    buttonHTML = (
      <button class="button" disabled>
        Submit
      </button>
    );
  } else {
    buttonHTML = (
      <div>
        <button
          // onclick={window.confirm("Are you sure you want to submit")}
          // onClick={window.dialog.show()}
          // onClick={dialogForm}
          class="button"
        >
          Submit
        </button>
      </div>
    );
  }

  return <div class="button-container">{buttonHTML}</div>;
}
export default Submit;
