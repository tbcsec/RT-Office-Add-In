Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;

    // Attach new event handlers
    document.getElementById("downloadFile").onclick = downloadFileFromGitHub;
    document.getElementById("openExploitDb").onclick = openExploitDb;

    updateRibbon();
    updateTaskPaneUI();
  }
});

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}

// Function to download a file from a GitHub repository
async function downloadFileFromGitHub() {
  const url = 'https://raw.githubusercontent.com/username/repository/branch/filename'; // Replace with actual URL

  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error('Network response was not ok.');
    }
    const fileContent = await response.text();
    console.log("Downloaded file content:", fileContent);
    // You can further process the file content or display it in Excel
    Office.context.ui.displayDialogAsync(fileContent);
  } catch (error) {
    console.error('Error downloading file:', error);
  }
}

// Function to open ExploitDB in a dialog window
async function openExploitDb() {
  const exploitDbUrl = 'https://www.exploit-db.com';
  Office.context.ui.displayDialogAsync(exploitDbUrl, { height: 50, width: 50 });
}
