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
      .addItem('Refresh All', 'reportMemberMatrix')
      .addSeparator()
      .addItem('Refresh BCA-CIM-Proforma', 'reportBCACIMProforma')
      .addItem('Refresh Members To Process', 'reportMembersToProcess')
      .addItem('Refresh Processed Members', 'reportMembersProcessed')
      .addSeparator()
      .addItem('Refresh Cancellation', 'reportCancellation')
      .addToUi();   
}
