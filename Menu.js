function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mark Membership')
      .addItem('Remove Membership', 'unMembership')
      .addItem('Give Membership', 'giveMembership')
      .addToUi();

  ui.createMenu('Input Data')
      .addItem('Input Membership Number', 'setMembershipNumber')
      .addItem('Input BirthYear', 'setBirthYear')
      .addToUi();   

  ui.createMenu('Refresh Matrix')
      .addItem('Refresh All', 'readData')
      .addSeparator()
      .addItem('Refresh BCA-CIM-Proforma', 'refreshBCACIMProforma')
      .addItem('Refresh Members To Process', 'refreshMembersToProcess')
      .addItem('Refresh Processed Members', 'refreshMembersProcessed')
      .addSeparator()
      .addItem('Refresh Cancellation', 'refreshCancellation')
      .addToUi();   
}

function readData() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();

  const reports = [
    reportBCACIMProforma,
    reportMembersToProcess,
    reportMembersProcessed,
    reportCancellation,
  ];

  reports.forEach(report => report(stmt));

  stmt.close();
  conn.close();
}

function refreshBCACIMProforma() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  reportBCACIMProforma(stmt);
  stmt.close();
  conn.close();
}

function refreshMembersToProcess() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  reportMembersToProcess(stmt);
  stmt.close();
  conn.close();
}

function refreshMembersProcessed() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  reportMembersProcessed(stmt);
  stmt.close();
  conn.close();
}

function refreshCancellation() {
  var conn = Jdbc.getConnection(url, username, password);
  var stmt = conn.createStatement();
  reportCancellation(stmt);
  stmt.close();
  conn.close();
}
