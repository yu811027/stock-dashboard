function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("0050正二工具")
    .addItem("建立/更新近三個月資料", "updateDashboard")
    .addToUi();
}
