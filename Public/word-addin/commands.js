// commands.js
Office.onReady(() => {
  console.log("Commands ready");
});

function testButton(event) {
  Office.context.document.setSelectedDataAsync("Hello from ForensiCode!", () => {
    event.completed();
  });
}

// Make the function globally accessible
window.testButton = testButton;
