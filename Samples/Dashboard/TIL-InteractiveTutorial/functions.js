const functions = {
  addSheetToDashboard: async (args, api) => {

    const sheets = await api.workbook.executeCommandAsync("tabdoc", 'get-sheets-info', {});
    console.log(sheets);
    const command = await api.workbook.executeCommandAsync("tabdoc", 'add-sheet-to-dashboard', {
      "add-as-floating": "false",
      "dashboard": args.targetDashboard,
      "worksheet": args.sourceSheet
    })
    console.log(command);

  }
}