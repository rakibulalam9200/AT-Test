export async function run() {
  try {
    await Excel.run(async (context) => {
      // Access the workbook and get or create "Sheet 2"
      const sheets = context.workbook.worksheets;
      let sheet2 = sheets.getItemOrNullObject("Sheet2");
      sheet2.load("name");

      await context.sync();

      if (sheet2.isNullObject) {
        // Create Sheet2 if it doesn't exist
        sheet2 = sheets.add("Sheet2");
      }

      // Activate Sheet2
      sheet2.activate();

      // Clear existing content (optional)
      sheet2.getRange().clear();

      // Add login page UI
      const usernameCell = sheet2.getRange("A1");
      const passwordCell = sheet2.getRange("A2");
      const loginButtonCell = sheet2.getRange("A3");

      usernameCell.values = [["Username:"]];
      usernameCell.format.font.bold = true;

      passwordCell.values = [["Password:"]];
      passwordCell.format.font.bold = true;

      loginButtonCell.values = [["[Login]"]];
      loginButtonCell.format.fill.color = "lightblue";
      loginButtonCell.format.font.color = "white";
      loginButtonCell.format.font.bold = true;
      loginButtonCell.format.horizontalAlignment = "Center";

      // Adjust the column widths for better visibility
      sheet2.getRange("A:A").format.columnWidth = 20;

      // Sync context changes
      await context.sync();

      console.log("Login page added on Sheet 2.");
    });
  } catch (error) {
    console.error(error);
  }
}
