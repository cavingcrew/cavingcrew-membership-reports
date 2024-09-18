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