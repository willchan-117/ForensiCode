Office.onReady(() => {
  console.log("Commands ready");
});

// Function triggered by the ribbon button
function testButton(event) {
  Office.context.document.setSelectedDataAsync(
    "Hello from ForensiCode!",
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      }
      event.completed();
    }
  );
}

// Make the function globally accessible to manifest
window.testButton = testButton;
