const authorizedUsers = ["saroxgaming@gmail.com", "lunarch.midnightwire@gmail.com"]

function PushProgressions() {
  if (authorizedUsers.includes(Session.getActiveUser().getEmail())){
    // Load in Spreadsheet
    const ss = SpreadsheetApp.openById("1QpYCO2swH7fly6izSHeo6VMaV7ipI6X_MVXuPArdcPQ");
    const playersSheet = ss.getSheetByName("Players");
    const logSheet = ss.getSheetByName("Progression Log");

    // Collect Data - Progressions
    const sourceRange = ss.getRangeByName("TeamProg");
    const sourceVals = sourceRange.getValues();

    // Collect PlayerIDs
    const playerIDs = ss.getRangeByName("PlayerIDs").getValues().flat();

    // Filter Progressions to Players  
    var cleanedData = sourceVals.filter(row => row[0] !== "");
    cleanedData = cleanedData.filter(subArray => !subArray.slice(-7).every(item => item === ""));

    // Loop through players
    for (let i = 0; i < cleanedData.length; i++) {
      // Gets PlayerID and finds row of player
      var searchID = cleanedData[i][0];
      var rowIndex = playerIDs.indexOf(searchID) + 2;

      // Gets player data and adds progression
      playerData = playersSheet.getRange(rowIndex, 23, 1, 7).getValues()[0];
      progressionArray = cleanedData[i].slice(7);
      newPlayerData = playerData.map((stat, index) => stat + progressionArray[index]);

      // Debug Logs
      //console.log(rowIndex);
      //console.log(playerData);
      //console.log(progressionArray);
      //console.log(newPlayerData);

      // Set player data to final stats
      playersSheet.getRange(rowIndex, 23, 1, 7).setValues([newPlayerData]);

      // Mark Player as Progressed
      playersSheet.getRange(rowIndex, 32).setValue(true);

      // Progression Logging
      const date = new (Date);
      const email = Session.getActiveUser().getEmail();
      const logData = [date, email, ...cleanedData[i].slice(0,3), ...progressionArray, ...newPlayerData];
      logSheet.appendRow(logData);
    }

    // Reset
    const resetRange = ss.getRangeByName("ResetProg");
    const resetValues = resetRange.getValues();

    for(let row = 0; row < 38; row++){
      for(let col = 0; col < 7; col++){
        if(typeof resetValues[row][col] === 'number'){
          resetValues[row][col]="";
        }
      }
    }

    resetRange.setValues(resetValues)
  }
  else {
    throw new Error("Unauthorized User");
  }
}
