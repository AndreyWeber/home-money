// Put test functions here
function test() {
  // Put test code here
  // const date = new Date(Date.UTC(2024, 2, 8));

  // const formatter = Intl.DateTimeFormat("en-GB", { style: 'short' });
  // const str = formatter.format(date);

  // const userLocale = Session.getActiveUserLocale();

  // const DateTime = luxon.DateTime; // Access Luxon's DateTime object
  // const local = DateTime.local();

  // const zoneName = local.zoneName;

  const test = getRowsData(getRawDataSheet());

  const scriptTimeZone = Session.getScriptTimeZone();

  var DateTime = luxon.DateTime;
  // var dt = DateTime.fromJSDate(cellData);
  // var scriptTimeZone = Session.getScriptTimeZone();
  // var dtRezoned = dt.setZone(scriptTimeZone, { keepLocalTime: true });
  // var a = dtRezoned.toString();

  const str = "08.03.2024";
  const dt = DateTime.fromFormat(str, "dd.MM.yyyy", { zone: scriptTimeZone });
  const dtStr = dt.toString();

  // const rawTransactions = getRawDataTransactionObjects();
  // const str = rawTransactions[0].dateOfTransaction.toString();
  const i = 1;
  // setLatestTransactionDate(date);
  // const metadataDate = getMetadataObject(MetadataKeys.LATEST_TRANSACTION_DATE);
}