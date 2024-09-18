function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mark Membership')
      .addItem('Remove Membership', 'unMembership')
      .addItem('Give Membership', 'giveMembership')
      
      .addToUi();

   // create menu for dispatch  functions
  ui.createMenu('Input Data')
      .addItem('Input Membership Number', 'setMembershipNumber')
      .addItem('Input BirthYear', 'setBirthYear')
      .addToUi();   

  ui.createMenu('Refresh Matrix')
      .addItem('Refresh All', 'readData')
     /* .addSeparator()
      .addItem('Refresh Output', 'runOutputTab')
      .addSeparator()
      .addItem('Refresh Volunteering', 'volunteerData')
      .addItem('Refresh Roles', 'runRolesTab')
      //.addItem('Refresh Volunteering-old', 'runVolunteeringTab')
           .addSeparator()
      .addItem('Refresh Lead Training', 'leadBelayTrainingData')
      .addItem('Refresh TR Belay Training', 'topRopeTrainingData')

                 .addSeparator()

      .addItem('Refresh Event Listings Dashboard', 'readEvents')
      .addItem('Refresh Badges', 'badgesData')
*/
      .addToUi();   
/*
        ui.createMenu('Badges')
      .addItem('Mark Given Badge', 'markGivenBadge')
      .addItem('Refresh Badges', 'badgesData')



      .addToUi();  

        ui.createMenu('Give Volunteer Competencies')
      .addItem('Check-In - SignOff', 'giveCheckIn')
      .addItem('Announcements - SignOff', 'giveAnnouncements')
      .addItem('Floorwalker - SignOff', 'giveFloorwalker')
        .addSeparator()
      .addItem('Pairing - Eager To Learn', 'givePairingToLearn')
      .addItem('Pairing - In Training', 'givePairingTraining')
      .addItem('Pairing - SignOff', 'givePairing')
      .addSeparator()
      .addItem('Trip Director - Eager To Learn', 'giveTripDirectorToLearn')
      .addItem('Trip Director - In Training', 'giveTripDirectorTraining')
      .addItem('Trip Director - SignOff', 'giveTripDirector')




      .addToUi();   */

}
