function ProgAdvance() {
  // Load in Spreadsheet
  const ss = SpreadsheetApp.openById("1QpYCO2swH7fly6izSHeo6VMaV7ipI6X_MVXuPArdcPQ");
  const leagueMeta = ss.getSheetByName("CWSFL Info");

  if(leagueMeta.getRange(2, 6).getValue() == false){
    // Load in specific sheets
    const playersSheet = ss.getSheetByName("Players");
    

    // Determine Length
    const lastRow = playersSheet.getLastRow() - 1;

    // Get column data                                 || Note to self: need to make this select only non-retired players ||
    const playerActivity = playersSheet.getRange(2, 2, lastRow, 1).getValues();
    const playerTeam = playersSheet.getRange(2, 3, lastRow, 1).getValues();
    const progBool = playersSheet.getRange(2, 32, lastRow, 1).getValues();
    const progMiss = playersSheet.getRange(2, 33, lastRow, 1).getValues();

    // Update missed prog counts
    for(let i = 0; i < progBool.length; i++){
      if(playerActivity[i][0] != "RETIRED"){
        if(playerTeam[i][0] != "" && playerTeam[i][0] != "DRAFT" && playerTeam[i][0] != "FREEAGENT"){
          //console.log(progBool[i][0])
          if(progBool[i][0] === false){
            let currentProgMiss = Number(progMiss[i][0]) || 0;
            progMiss[i][0] = currentProgMiss + 1;
            if(progMiss[i][0] >= 3){
              playerActivity[i][0] = "INACTIVE";
            }
          }
          else {
            progMiss[i][0] = 0;
            playerActivity[i][0] = "ACTIVE";
          }
        }
      }
    }
    
    // Update League's Current Progression
    previousProgression = leagueMeta.getRange(3, 2).getValue();
    currentProgression = previousProgression + 1;

    // Update sheets
    playersSheet.getRange(2, 33, lastRow, 1).setValues(progMiss);
    playersSheet.getRange(2, 2, lastRow, 1).setValues(playerActivity)
    playersSheet.getRange(2, 32, lastRow, 1).setValue(false);
    leagueMeta.getRange(3, 2).setValue(currentProgression);

    // Turn on Safety
    leagueMeta.getRange(2, 6).setValue(true);
  }
  else{
    // Notify user of safety lock
    throw new Error("Safety Lock on");
  }
}
