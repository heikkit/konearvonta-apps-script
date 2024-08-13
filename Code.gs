/*
Hello. Hope you like spaghetti...
Noodles cooked by Heikki Tammisalo (HEX)

Changelog:
3.9.2019
-lisätty trimUnusedPins -ajo konearvonnan (korttiPakka) alkuun. Pitäisi vapauttaa varatuiksi jääneet koneet.
4.9.2019
-nyt skripti katsoo ryhmäkoon A1-solusta, jonne lisätty kaava. Pelaajat-välilehdellä väärä ryhmäjako.
26.9.2019
-korjattu bugi, jossa ainoaa vapaana olevaa konetta ei saanut pelata, vaikka olisi alle 2 peliä/pelaaja sillä.
-hiottu tyhmennys-skripti poistamaan kaarisulut tageista

*/

// Global vars ...
var ss = SpreadsheetApp.getActiveSpreadsheet();
var koneSheetti = ss.getSheetByName("Koneet");
var playerSheet = ss.getSheetByName("Pelaajat");
var sheetName;
var groupSheet;
var koneDataRange;
var koneData;
var caller;
var cell;
var uusinta;
var pinCellColors = ["#f9cb9c", "#fff9cb9c", "#fce5cd", "#fffce5cd", "#ffe599", "#ffffe599", "#e69138", "#ffe69138"]; // pudotuspelien konesolujen taustavärit
var scoreCellColors = ["#ffffff", "#ffffffff"];
var pudotuspeli = false;

function onOpen() {  
  
  // Create custom menus
  var ui = SpreadsheetApp.getUi();    
  
  ui.createMenu('Konearvonta 🎲')
      .addItem('Fight! 🤼 \(Ctrl+Alt+Shift+0\)', 'korttiPakka')
    /*  
      .addItem('Debug: tyhjennä ryhmätaulukko', 'clearGroupSheet')
      .addItem('Debug: täytä ryhmätaulukko', 'fillGroupSheet')
      .addItem('Debug: tyhjää pelit koneista', 'clearKoneet')  
      .addItem('Debug: trimUnusedPins', 'trimUnusedPins')
      .addItem('Debug: pinCellProtectOn;', 'pinCellProtectOn')
      .addItem('Debug: pinCellProtectOff;', 'pinCellProtectOff')
      .addItem('Debug: showCoordinates;', 'showCoordinates')
    */  
      .addToUi();  
      
}


/*
Täällä määritellään, mitä tapahtuu, kun pelitaulukoiden soluja sörkitään.
*/
function onEdit() {
  
    //return;
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetName = ss.getActiveSheet().getName();
  var override = ss.getActiveSheet().getRange("A23").getValue();
  
  //Browser.msgBox(override);
  
  if (sheetName.indexOf("RYHMÄ") > -1) {
    if (!override) {
       return;
    }
  }
  
  // Odotetaan max. 30 sekuntia lukitusta funktioon.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
    
  if ((sheetName.indexOf("RYHMÄ ") < 0) && (sheetName.indexOf("Finaali") < 0) && (sheetName.indexOf("Tiebreak") < 0)) {  
    Logger.log("oneEdit() terminating. \'" + sheetName + "\' is not a group sheet");    
    return;
  }
  
  else if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1))
    pudotuspeli = true;
  
  Logger.log("onEdit");
  Logger.log("sheetName: " + sheetName);
  caller = "onEdit";  
  
  groupSheet = SpreadsheetApp.getActiveSheet();
  
  // Refresh koneData from "Koneet" sheet.
  koneDataRange = koneSheetti.getDataRange();
  koneData = koneDataRange.getValues();  
  cell = groupSheet.getActiveCell();
  
  //Logger.log("onEdit - cell: " + cell.getValue());
  
  var rivi = cell.getRow();
  var sarake = cell.getColumn();  
  var sheet = ss.getActiveSheet();  

  var kone;
  var matchRow;  
  var arvo = cell.getValue();  
  var pelaajat = findPlayerPair(rivi, sarake);
  var p1 = pelaajat[0]; //aloittava pelaaja
  var p2 = pelaajat[1];
  var p1Tag = pelaajat[2];
  var p2Tag = pelaajat[3];

  /* Ryhmätaulukko */
  if (sheetName.indexOf("RYHMÄ ") > -1) {
    
Logger.log("RYHMÄSHEETTI");    
    var inTheGreen = isInsideScoringArea(rivi, sarake);
Logger.log("onEdit - inTheGreen: " + inTheGreen);
    var koneSolu = /*inTheGreen &&*/ (sarake % 2 == 0);
Logger.log("onEdit - koneSolu: " + koneSolu);
    var pisteSolu = /*inTheGreen &&*/ (sarake % 2 == 1);
Logger.log("onEdit - pisteSolu: " + pisteSolu);    
    if (inTheGreen == "false") {
      Logger.log("NYT LOPPU");
      return;
    }
  }  
  else if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) { // Pudotuspelit. Katsotaan, että solun väri löytyy konesoluvärien listalta. Ja alta löytyy pistesolun värinen solu.
    Logger.log("PUDOTUSPELI");
    var bg = cell.getBackground();    
    var koneSolu = ((pinCellColors.indexOf(bg) > -1) && (scoreCellColors.indexOf(cell.offset(1, 0).getBackground()) > -1));
  }
  
Logger.log("IN THE GREEN: " + inTheGreen);
Logger.log("KONESOLU: " + koneSolu);
Logger.log("PISTESOLU: " + pisteSolu);  
  
Logger.log("koneSolu == true: " + (koneSolu == true));
Logger.log(typeof koneSolu);
  if (koneSolu) {
Logger.log("koneSolu!!!");
  }
  else {
    Logger.log("EI OO MUKA KONESOLU");
    Logger.log("Solu: " + cell.getA1Notation());
    Logger.log("Solun arvo: " + cell.getValue());
  }
  
  
  // Manuaalinen koneen syöttäminen matsin aloittamiseksi
  if (koneSolu) {
Logger.log("KONESOLU");
    var lyhenne = arvo;
    
    console.log("Konesoluun " + cell.getA1Notation() + " ryhmässä " + sheetName[6] + " kirjoitettu: " + arvo);
     
    if (typeof(lyhenne) == "number") {
      cell.clearContent();
      Browser.msgBox("Nyt meni pelkkiä numeroita konesoluun?");
      return;
    }
      
    kone = lyhenne.toUpperCase()    
    if (kone != "" && findRow (koneSheetti, 3, kone) <= 0) {
      cell.clearContent();
      Browser.msgBox("Ei löydy konetta \'" + kone + "\' - tarkista lyhenne \'Koneet\'-taulukosta.");
      return;
    }      
    cell.setValue (kone);
    matchRow = checkForMatch(p1Tag, p2Tag, kone, koneSheetti); // Gets row in 'Koneet' sheet for ongoing match between p1 + p2; 0 if no match.    
Logger.log("MATCHROW: " + matchRow);
    
    // ei sallita pelin aloittamista päätetyn pelin kohdalla
    //var pisteet = sheet.getRange(rivi, sarake + 1).getValue();
    
    if (sheetName.indexOf("RYHMÄ ") > -1) {
      var pisteet = cell.offset(0, 1).getValue();
      if (pisteet != "") {
        //Browser.msgBox("Peliin on jo merkitty pisteet!");
        console.log("Vaihdettiin peli " + kone + " pelattuun matsiin solussa " + cell.getA1Notation());
        return;
      }
    }

Logger.log("***matchRow: " + matchRow);  
// TÄHÄN JÄÄTIIN    
//DEBUG VAR:    
//if (sheetName.indexOf("RYHMÄ ") < 0)
//  pelaajat = [0, 0, 0, 0];
//</DEBUG>  
Logger.log("pelaajat: " + pelaajat);    
    var busy = checkIfBusy(pelaajat[2], pelaajat[3], kone, koneSheetti);
      //Logger.log("BUSY: " + busy);
      if (busy != 0) {
        //Logger.log("BUSY != 0");
        // Tämä if - else hoitaa koneen deletoinnin ryhmäsheetistä - saattaa rikkoa jotain?
        //Logger.log("MATCHROW: " + matchRow);
        if (matchRow >= 0) {
          Browser.msgBox(busy + " ei ole vapaana?");
          cell.clearContent();
          return;
        }
        else { // Löytyi matsi näillä pelaajilla mutta eri koneella.
          endMatch(-matchRow);
          //cell.clearContent();
        }
      }
      
      var twoPlaysRule = check2PlaysRule(kone);    
      if ((twoPlaysRule == "x" || twoPlaysRule == "y")) {
        Browser.msgBox("Korkeintaan 2 peliä samalla koneella per pelaaja!");
        cell.clearContent();
        return;
      }         
  
      
      else if (matchRow == 0) { 

        //Logger.log("matchRow == 0");
        beginMatch (pelaajat, kone);
      }
      else if (matchRow > 0) { // Full match: p1Tag, p2Tag, kone
        //Logger.log("onEdit found full match - ending match!!");
        endMatch(matchRow);
      }    
      else if (matchRow < 0) { //Match found without machine (since cell is already empty)
        //Logger.log("HEP2!");
        //Logger.log("onEdit uusinta: " + uusinta);
        uusinta = 1;    
        if (kone != "") {
          //Logger.log("HEP3!");
          endMatch(-matchRow, beginMatch (pelaajat, kone));
        }
        else {
          //Logger.log("HEP4");
          endMatch(-matchRow);
        // beginMatch (pelaajat, kone);
        }
      }
    }

//DEBUG  
//return;
    
Logger.log("pisteSolu: " + pisteSolu);
//Logger.log("isInsideScoringArea(rivi, sarake): " + isInsideScoringArea(rivi, sarake));
  
    //Tuloksen syöttäminen pistesoluun
    if (pisteSolu) {
      
      console.log("Pistesoluun " + cell.getA1Notation() + " ryhmässä " + sheetName[6] + " kirjoitettu: " + arvo);

      Logger.log("PISTESOLU");    
      piste = arvo;    
      Logger.log("piste: " + piste);    
      // var koneArvo = sheet.getRange(rivi, sarake - 1).getValue();
      var koneArvo = cell.offset(0, -1).getValue();    
      if (koneArvo == "" && typeof(piste) == "number") {
        Browser.msgBox("Jospa ei merkata pisteitä pelaamattomaan matsiin?")
        cell.clearContent();
        return;
      }   
      if (typeof(piste) != "number" && piste != "") {
        Browser.msgBox("Vain numeroita pistesoluun, kiitos.");
        cell.clearContent();
        return;
      }    
      else if (!pudotuspeli && (piste > 1 || piste < 0)) {
        Browser.msgBox("Pisteeksi vain 1 tai 0, kiitos..");
        cell.clearContent();
        return;
      }
      kone = groupSheet.getRange(cell.getRow(), (cell.getColumn() - 1)).getValue();
      matchRow = checkForMatch(p1Tag, p2Tag, kone, koneSheetti); // Gets row in 'Koneet' sheet for ongoing match between p1 + p2; 0 if no match.        
      if (piste == 0 || piste == 1)
        endMatch(matchRow, function() { Logger.log ("endMatch tehty"); });
      //var testink = koneSheetti.getRange(matchRow, 4, 1, 5);
      //Logger.log("tyhjenikö? " + testink.getValues());
      //pinCellProtectOff();
    }
  
  
  /* Pudotuspelit */
  if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) { 
Logger.log("Pudotuspelit jee");
    // Pisteen syöttäminen
    var kone;
    //var bg = cell.getBackground();
    var pisteSolu = (scoreCellColors.indexOf(bg) > -1);
//    var koneSolu = (pinCellColors.indexOf(bg) > -1);
Logger.log("pisteSolu: " + pisteSolu);    
Logger.log ("bg: " + bg);
Logger.log ("scoreCellColors: " + scoreCellColors);

    if (pisteSolu) {      
      var counter = 1;
      while (pisteSolu) {
        bg = cell.offset(-counter, 0).getBackground();
        pisteSolu = (scoreCellColors.indexOf(bg) > -1);
        if (!pisteSolu) {
          kone = cell.offset(-counter, 0).getValue();
Logger.log("kone: " + kone);
          var rivi = findRow(koneSheetti, 3, kone)
Logger.log("rivi: " + rivi);
          endMatch(rivi);
        }
Logger.log(bg);
        counter++;
      }      
    }
    if (koneSolu) {
      kone = cell.getValue().toUpperCase();
      cell.setValue(kone);
      beginMatch(0, kone);
    }
  }
//Logger.log("onEdit - releasing LOCK");
  lock.releaseLock();
}


/* Yrittää aloittaa matsin valituilla pelaajilla ja koneella */
function beginMatch(pelaajat, kone) {
  
  
  Logger.log("beginMatch");
  Logger.log("beginMatch pelaajat: " + pelaajat, " kone: " + kone);
  
  
  if (uusinta == 1) {
    //Logger.log("beginMatch uusinta: " + uusinta);
    // Refresh koneData from "Koneet" sheet.
    
    //koneDataRange = koneSheetti.getDataRange();
    //koneData = koneDataRange.getValues();
  }

  if (pelaajat != 0) {
    var starter = pelaajat[0];
    var follower = pelaajat[1];
    var p1Tag = pelaajat[2];
    var p2Tag = pelaajat[3];  
    var sheet = playerSheet; 
    
    //Browser.msgBox(pelaajat);
    
    
    if (p1Tag == ",") //...mistä tuohon edes tulee pilkku...
      throw("Virhe: ei löytynyt pelaajia :( Onko tagit merkitty?");
  
    // Tarkistaa, onko pelaajia jo pelissä.
    if ((findRow(koneSheetti, 6, p1Tag) > 0) || (findRow(koneSheetti, 6, p2Tag) > 0) || (findRow(koneSheetti, 7, p1Tag) > 0) || (findRow(koneSheetti, 7, p2Tag) > 0)) {
      if (!(cell.isBlank()) && uusinta == 0) { // ignoorataan, jos kyseessä koneen vaihto
        Browser.msgBox("Pelaaja merkitty jo peliin?");
        cell.clearContent();
        return;
      }    
    }
  }
  
  if (kone != "") {
    
Logger.log("kone: " + kone);  
Logger.log("koneSheetti:" +  koneSheetti.getName());
  
    var koneRivi = findRow (koneSheetti, 3, kone) - 1;

Logger.log("koneRivi: " + koneRivi);    

  
    //Logger.log(koneRivi);
    
    /*
    if (koneRivi <= 0) {
        //koneen tyhjyyden tarkistaminen järkevämmäksi...?
        Browser.msgBox("Konetta ei löydy. Tarkista lyhenne \"Koneet\"-taulukosta.");
        cell.clearContent();
        return;    
    } 
    */
    
    if (koneData[koneRivi][4] == false) {
      Browser.msgBox("Kone ei käytössä? Tarkista \"Koneet\"-taulukko.");
      //var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
      cell.clearContent();
      return;
    }    
  
    if (koneData[koneRivi][5] == "EI") {
      Browser.msgBox("Kone on jo varattu?");
      //var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
      cell.clearContent();
      return;
    }
  }  
  
  var rivi = findRow(koneSheetti, 3, kone);  
  
  // Tullaanko tähän ehtoon edes koskaan...
  if ((rivi == 0) && (cell.getValue() != "")) {
    Browser.msgBox("Konetta ei löytynyt. Tarkista lyhenne \"Koneet\"-taulukosta.")
    cell.clearContent();
    return;
  }

  try {
  var riviData = koneSheetti.getRange(rivi, 3, 1, 6).getValues();
  }
  catch (err) {
    throw err;
  }
    
  var group = getGroupLetter();
  if (group == "")
    throw("Virhe: ei löytynyt ryhmän nimeä :(");
  
  Logger.log("!!!PELAAJAT: " + pelaajat);  
  
  
  // Merkataan matsi koneet-taulukkoon    
  if ((riviData[0][2] == true) && (riviData[0][3] != "EI") && (riviData[0][4].length == 0) && (riviData[0][5].length == 0)) { // Kone merkattu käyttöön, ei varattu, jne.
    var matchData = [];
    if (pelaajat != 0)
      matchData.push(["EI", p1Tag, p2Tag, group]);    
    else
      matchData.push(["EI", "", "", group()]);
    
    var r = koneSheetti.getRange(rivi, 6, 1, 4);
    r.setValues(matchData);
    if (caller != "korttiPakka")
      SpreadsheetApp.flush(); // Kirjoitetaan datat sheetille ennen kuin dialogi pausettaa skriptin.
  }
  else
    throw("Nyt on jotain vikaa matsin parametreissa :(");
    
  
  
  var koneNimi = koneData[rivi -1][2];
  var koneLyhenne = koneData[rivi -1][3];
  
  if (pelaajat != 0) {
    announceMatch(koneNimi, koneLyhenne, starter, follower);
    console.log("Matsi aloitettu: " + koneNimi + " " + p1Tag + " " + p2Tag + " " + group);
  }
}



function endMatch(row, callback) {
  
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    cell = ss.getActiveCell();
  }
  catch(err) {
    Logger.log("endMatch - ei aktiivista solua");
  }
    
  Logger.log("endMatch");
    
  if (row <= 0) {
    return;
  }  
  
  try {
    koneSheetti.getRange(row, 7, 1, 3).clearContent();
    koneSheetti.getRange(row, 6).setValue("KYLLÄ");
  }
  catch(err) {
    throw("Matsin poistaminen epäonnistui, kutsu apua? :( - " + err);
  }
  var score;
  try {
    score = cell.getValue(); //koitetaan napata piste talteen
    console.log("Matsi lopetettu: " + koneData[row - 1][2] + " " +  koneData[row - 1][6] + " " + score + " " + koneData[row - 1][7] + " " + (1 - score) + " " + koneData[row - 1][8]);
  }
  catch (err) {
    //console.log("Rivin " + row + " matsi loppui, mutta loggaus epäonnistui. Peruutus?");
    //console.log("rivin " + row + " sisältö: " + koneSheetti.getRange(row, 3, 1, 7).getValues());
    console.log("Koneen " + koneSheetti.getRange(row, 3, 1, 1).getValues() + " matsi loppui, mutta loggaus epäonnistui. Peruutus?");
  }
  
  //SpreadsheetApp.flush(); // ei auta kun onEdit() ei triggeröidy :(
  
  if (typeof(callback) != "undefined")
    callback();
}


function checkForMatch(p1Tag, p2Tag, kone, koneSheetti) {
   
  Logger.log("checkForMatch");
    
  var p1Row = findRow(koneSheetti, 6, p1Tag); //(...muistetaan ne offsetit...)
  var p2Row = findRow(koneSheetti, 7, p2Tag);
  var pinRow = findRow(koneSheetti, 3, kone);
  
  
  if (p1Row == p2Row && p2Row == pinRow) { //Found match with p1, p2, pin
    return pinRow; // complete match found, return row in koneet sheet
  }
  else if (p1Row == p2Row) {
    return -p1Row; // match found with players but different pin, return negative row nr.
  }
  else {
    return 0;
  }
}

/*
Checks if any of the match components, i.e. p1, p2 or pin are busy.
Returns the busy party.
*/

function checkIfBusy(p1Tag, p2Tag, kone, koneSheetti) {
  
  Logger.log("checkIfBusy");  
  
  var p1p1 = findRow(koneSheetti, 6, p1Tag);
  var p1p2 = findRow(koneSheetti, 7, p1Tag);
  var p2p1 = findRow(koneSheetti, 6, p2Tag);
  var p2p2 = findRow(koneSheetti, 7, p2Tag);
  var pinRow = findRow(koneSheetti, 3, kone);
  var pinState;
  
//  Logger.log("checkIfBusy - koneSheetti: " + koneSheetti.getName());
 // Logger.log("checkIfBusy - pinRow: " + pinRow);
  
  if (pinRow > 0)
    pinState = koneSheetti.getRange(pinRow, 6).getValue();

  if ((pinRow > 0) && (pinState == "1")) {
    return kone; 
  }
  else if (p1p1 > 0 || p1p2 > 0) {
    return p1Tag;
  }
  else if (p2p1 > 0 || p2p2 > 0) {
    return p2Tag;
  }
  else
    return 0;
}


//Etsii rivin taulukosta sarakkeen ja stringin perusteella
function findRow(sheet, column, string){
  
  Logger.log("findRow");
  
  var data;  
  if (sheet.getName() == "Koneet" && typeof(koneData) != "undefined") { //Konetaulukko on jo arrayssa, niin käytetään sitä
    data = koneData;
  Logger.log("findRow DATA LENGTH: " + data.length);
  }
  
  
  else data = sheet.getDataRange().getValues();  
  if (typeof(data) == "undefined") {
    data = sheet.getDataRange().getValues();
  }  
  var found = 0;  
  for(var i = 0; i < data.length; i++){
    if(data[i][column] == string){
      found = 1;
      return i+1;
    } 
  }
  if (found == 0)
    return 0;
}

//Hoitaa uusinta-arvonnat
function arvoUusi() {
  cell = ss.getActiveCell();
  Logger.log("arvoUusi")
  Logger.log(ss.getActiveCell().getA1Notation());
  //jos ollaan pistesolussa, otetaan yksi ruutu vasempaan päin.
  if (cell.getColumn() % 2 == 1)
    moveLeft();
  peruutaMatsi(function() { Logger.log ("peruutaMatsiCALLBACK"); });
  emptyCell(function() { Logger.log ("emptyCellCALLBACK"); });
  korttiPakka();
}

function emptyCell(callback) {
  solu = ss.getActiveCell();
  var kone = solu.getValue();
  solu.clearContent();
  var koneRivi = findRow (koneSheetti, 3, kone);
//  Logger.log("emptyCell konerivi: " + koneRivi);
  callback();
}


/*
 Konearvonnan pääfunktio. Käynnistyy valikosta.
*/
function korttiPakka() {
  
  
  
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  //Kokeillaan, lisääkö trimUnusedPins merkittävästi suoritusaikaa...
  trimUnusedPins();
     
  Logger.log("korttiPakka"); 
  var sheet = SpreadsheetApp.getActiveSheet();
  sheetName = sheet.getName();
  cell = sheet.getActiveCell();
  //cell = sheet.getCurrentCell();
  var rivi = cell.getRow();
  var sarake = cell.getColumn();
  var groupLetter;
  uusinta = 0;
  
  if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1))
    pudotuspeli = true;
  
  if (caller != "onEdit")
    caller = "korttiPakka";
  
  // Check if selected cell is in correct area for machine marking.  
  //var inTheGreen = isInsideScoringArea(rivi, sarake);  

  if(sheetName.indexOf("RYHMÄ ")>-1) {
    if (isInsideScoringArea(rivi, sarake) == "false") {
      errorDialog("Nyt meni kokonaan ohi merkkausalueelta...");
      return;
    }    
    if ((sarake % 2 == 1) || (sarake < 3) || (rivi < 3)) {
      errorDialog("Konearvonnat vasempaan soluun, kiitos.");
      return;
    }
  
    else if (cell.getValue() != "") {
      var pisteet = sheet.getRange(rivi, sarake + 1).getValue();
      if (typeof(pisteet) != 'number') {
        uusinta = 1;
        arvoUusi();
        return;
      }
      else {
        errorDialog("Peliin on jo merkitty tulos.");
        return;
      }
    }  
    groupLetter = sheet.getName()[6];
  }
  
  if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) { 
    Logger.log("Sheet is Pudotuspelit or tiebreak");
    var bg = cell.getBackground();

    if (pinCellColors.indexOf(bg) < 0) {
      throw("Väärän värinen solu!");
      return;
    }
    else if (scoreCellColors.indexOf(cell.offset(1, 0).getBackground()) < 0) {
      throw("Ei siihen!");
      return;
    }
  }
  
  // Refresh machine data from "Koneet" sheet.
  koneDataRange = koneSheetti.getDataRange();
  koneData = koneDataRange.getValues();

  //if(sheetName.indexOf("RYHMÄ ")>-1) {
    var pelaajat = findPlayerPair(rivi, sarake);
    var starter = pelaajat[0];
    var follower = pelaajat[1];
    var p1Tag = pelaajat[2];
    var p2Tag = pelaajat[3];
  //}

Logger.log("korttiPakka - pelaajat: " + pelaajat);

  var kone; 
  var koneNimi;
  var koneLyhenne;
  var vapaaLkm;  // vapaat koneet
  var cellValue = cell.getValue();
  var tooManyPlays;
  var kokeiltuLkm;
  
  // Arvotaan sellainen kone, jota kumpikaan ei ole pelannut vielä 2 kertaa.  
  kone = arvoKone(); // returns [koneNimi, koneLyhenne, vapaaLkm], or 0 if no machines free.  
  if (kone == 0) // ei koneita vapaana, dialogi näytetty
    return;
  koneLyhenne = kone[1];
  vapaaLkm = kone[2];  
  
  
  if (vapaaLkm >= 2)
    tooManyPlays = check2PlaysRule(koneLyhenne, 1); // returns 'x', 'y', or 0 (if neither player has too many plays)
  else if (vapaaLkm == 1)
    tooManyPlays = check2PlaysRule(koneLyhenne, 2); // returns 'x', 'y', or 0 (if neither player has too many plays)
  
  
Logger.log("korttiPakka - koneLyhenne: " + koneLyhenne);
Logger.log("korttiPakka - vapaaLkm: " + vapaaLkm);
Logger.log("korttiPakka - tooManyPlays: " + tooManyPlays);
  if (vapaaLkm == 1 && (tooManyPlays == 'x' || tooManyPlays == 'y')) { // vapaana 1 kone, jota joku pelannut jo 2 kertaa.
    errorDialog("Ei sopivia koneita vapaana.");  
    //Browser.msgBox("Ei sopivia koneita vapaana.");
    return;
  }
  else if (vapaaLkm > 1) { // vapaana useampi kuin 1 kone    
    
    
    var kokeillutKoneet = [];
    kokeiltuLkm = kokeillutKoneet.length;    
    
Logger.log("tooManyPlays: " + tooManyPlays + ", kokeiltuLkm:" + kokeiltuLkm + ", vapaaLkm: " + vapaaLkm);     
    // Katsotaan, löytyykö pelaamatonta konetta:
    while (tooManyPlays != 0 && kokeiltuLkm < vapaaLkm) {
Logger.log("TSEKATAAN ONKO PELAAMATTOMIA");      
      
      kone = arvoKone();
      koneLyhenne = kone[1];
      if (kokeillutKoneet.indexOf(koneLyhenne) < 0) {
        kokeillutKoneet.push(koneLyhenne);
        kokeiltuLkm = kokeillutKoneet.length;
      }
      tooManyPlays = check2PlaysRule(koneLyhenne, 1);
Logger.log("kutsuttu check2PlaysRule(koneLyhenne, 1)");
    }
    
//Browser.msgBox("TÄÄLLÄ OLLAAN, " + tooManyPlays + " " + kokeillutKoneet);
    
//Browser.msgBox("tooManyPlays: " + tooManyPlays + ", koneLyhenne: " + koneLyhenne);
    
    if (tooManyPlays != 0) {    
//Browser.msgBox("TÄNNE MENNÄÄN");      
      kokeillutKoneet = [];
      kokeiltuLkm = kokeillutKoneet.length;
    // Katsotaan, löytyykö max. kerran pelattua konetta:
      while (tooManyPlays != 0 && kokeiltuLkm < vapaaLkm) {
        kone = arvoKone();
        koneLyhenne = kone[1];
        if (kokeillutKoneet.indexOf(koneLyhenne) < 0) {
          kokeillutKoneet.push(koneLyhenne);
          kokeiltuLkm = kokeillutKoneet.length;
        }
        tooManyPlays = check2PlaysRule(koneLyhenne, 2);
      }
    }    
    
    
  }
  
  if (tooManyPlays != 0) {
    if (caller = "korttiPakka")
      errorDialog("Ei sopivia koneita vapaana.");      
    else      
      Browser.msgBox("Ei sopivia koneita vapaana.");
    return;
  }
  
  if (cellValue != "")
    uusinta = 1;
  
  //if(sheetName.indexOf("RYHMÄ ")>-1) {
  
  // Tarkistetaan, onko kyseessä uusinta-arvonta (matsi jo käynnissä näillä tageilla ja solun koneella)
  var matchRow = 0; // Rivi, jolta matsi ehkä löytyy
  if (cellValue.length > 0) // Vain, jos solu ei tyhjä.
     var matchRow = checkForMatch(p1Tag, p2Tag, cellValue, koneSheetti);  
  if (matchRow > 0) {
    endMatch(matchRow); // Lopetetaan vanha matsi ennen uuden koneen arvontaa.
  }     
  // Tarkistetaan, löytyykö pelaajatageja tai konetta jo jostakin matsista.
  var busy = checkIfBusy(pelaajat[2], pelaajat[3], koneLyhenne, koneSheetti);
  if ((busy != 0) && (uusinta == 0)) {
    //Logger.log("******I'M HEEEEEEEEEEEEEERE*");
    //Browser.msgBox(busy + " ei ole vapaana?");
    errorDialog(busy + " ei ole vapaana?");
    cell.clearContent();
    return;
  }  
  moveRight(); //liikutetaan valintaa oikealle. workaround,
                 //jos käyttäjä ehtii painaa enteriä arvonnan aloittamisen jälkeen.  
  cell.setValue(koneLyhenne); 
  
  beginMatch (pelaajat, koneLyhenne);

  Logger.log("korttiPakka - cell:" + cell.getValue());
  lock.releaseLock();
}

// Näyttää dialogin arvonnan tuloksesta: kone ja pelaajat.
function announceMatch(koneNimi, koneLyhenne, starter, follower) {
   
  Logger.log("announceMatch");
  
//DEBUG
  if (pudotuspeli) {
    Logger.log("Pudotuspeli, ei näytetä aloitusdialogia.");
    return;
  }  
  
  var arvontaTulos = [];

  if (caller == "korttiPakka") {
    arvontaTulos = [koneNimi + " (" + koneLyhenne + ")", starter, follower];
    dialogi = arvontaTulos; // Olikohan tälle syytä... global var?
    matchDialog(dialogi);
  }
  else {
    arvontaTulos = ["🎲 " + koneNimi + " (" + koneLyhenne + ")", " ️1️⃣️ " + starter, " 2️⃣ " + follower];
    dialogi = arvontaTulos; // Olikohan tälle syytä... global var?
    SpreadsheetApp.getUi().alert(arvontaTulos);
  }
  
  //afterCheck(koneLyhenne, starter);

  //console.log("Matsi aloitettu: " + arvontaTulos);
  
}

// Find player pair based on currently selected cell.
function findPlayerPair(rivi, sarake) {

Logger.log("findPlayerPair");
Logger.log("pudotuspeli: " + pudotuspeli);
  
  if (pudotuspeli) {
    
    Logger.log("findPlayerPair - pudotuspeli");
    
    var bg = cell.offset(1, 0).getBackground();
    var playerCount = 1;
    
    while (scoreCellColors.indexOf(bg) > -1) {
      bg = cell.offset(playerCount, 0).getBackground();
      playerCount++;
    }
    playerCount = playerCount - 2;
    
    //var playerNames = [];
    var players = [];
    var tags = [];

    

    
    var columnsToPlayers = getColumnsToPlayers();
//Logger.log("columnsToPlayers: " + columnsToPlayers);
    
    /*
    var offsetToPlayers = 0;
    bg = cell.offset(1, 0).getBackground();
    while (scoreCellColors.indexOf(bg) > -1) {
      bg = cell.offset(1, -i).getBackground();
      i++;
      Logger.log("BG: " + bg);
      Logger.log("CELL: " + cell.offset(1, -i).getA1Notation());
    }
    */
    
    for (i = 0; i < playerCount; i++) {
      players.push(cell.offset(i+1, columnsToPlayers).getValue());
//Logger.log("players[i]: " + players[i]);
      tags.push(players[i].substring(
      players[i].lastIndexOf("(") + 1, 
      players[i].lastIndexOf(")")
      ));
    }
    
    var p1Tags;
    var p2Tags;
        
    switch(tags.length) {
      case 4:
        p1Tags = tags[0] + "," + tags[2];
        p2Tags = tags[1] + "," + tags[3];
        break;
      case 3:
        p1Tags = tags[0] + "," + tags[2];
        p2Tags = tags[1];
        break;
      default:
        p1Tags = tags[0];
        p2Tags = tags[1];
    }    
    return ["Pudotuspelaaja(t) 1", "Pudotuspelaaja(t) 2", p1Tags, p2Tags];
  }
  
  else { // karsintapelit...
      // If selected cell is score cell, count one column to the left.
    if (sarake % 2 == 1)
      sarake = sarake - 1;
    
    // Find player coordinates based on x = 2nd row player tags, y = row in 2nd column player list
    var x_p2 = ((sarake - 4) / 2) + 1;
    var y_p2 = x_p2 + 2;
    
    // Fetch names of players from 2nd column
    var sheet = SpreadsheetApp.getActiveSheet();
    var yPlayer = sheet.getRange(rivi, 2).getValue(); // Player from left column 
    var xPlayer = sheet.getRange((y_p2), 2).getValue(); // Player from top row  
    
    // Get tags from 2nd row  
    var yTagRange = sheet.getRange(2, (2 + ((rivi - 2) * 2))); //Debuggaillaan tätä vielä...
    var yTag = yTagRange.getValue();  
    var xTagRange = sheet.getRange(2, (2 + x_p2 * 2));
    var xTag = xTagRange.getValue();  
    
    // Check if both players are available
    var p1Status = sheet.getRange(rivi, 3).getValue();
    var p2Status = sheet.getRange((y_p2), 3).getValue();  
    var starter = whoStarts(xPlayer, yPlayer, x_p2, y_p2, rivi);
    var starterTag;
    var follower;
    var followerTag;
    
    if (starter == yPlayer) {
      starterTag = yTag;
      follower = xPlayer;    
      followerTag = xTag;    
    }  
    else {
      followerTag = yTag;
      follower = yPlayer;
      starterTag = xTag;
    }
  return [starter, follower, starterTag, followerTag];
  }
}

function getColumnsToPlayers () {
  var i = 1;
  bg = cell.offset(1, 0).getBackground();
  while (scoreCellColors.indexOf(bg) > -1) {
    bg = cell.offset(1, -i).getBackground();
    if (scoreCellColors.indexOf(bg) < 0) {
      Logger.log("BG: " + bg);
      Logger.log("CELL: " + cell.offset(1, -i).getA1Notation());
    }
    i++;
  }
  return -(i - 1);
}

/* Palauttaa aloittavan pelaajan tagin */
function whoStarts(xPlayer, yPlayer, x_p2, y_p2, rivi) {  
  Logger.log("whoStarts");
    
  if (x_p2 % 2 == 0) { // Parillinen pelaajasarake
    if (rivi % 2 == 1) { // Pariton pelaajarivi
      return xPlayer; // aloittaja on 2. rivin pelaaja
    }
    else { // Parillinen pelaajarivi, pariton pelaajasarake
      return yPlayer;  // aloittaja on B-sarakkeen pelaaja
    }
  }
  else if ((x_p2 % 2 == 1) && (rivi % 2 == 1)) { // Pariton pelaajarivi, pariton pelaajasarake 
    return yPlayer; // aloittaja on 2. rivin pelaaja
  }
  else { // Parillinen pelaajarivi, pariton pelaajasarake
    return xPlayer; // aloittaja on B-sarakkeen pelaaja
  }
}


function getRandomInt(max) {  //Gives a random number. Replace with random.org API version later?
  Logger.log("getRandomInt");
  return Math.floor(Math.random() * Math.floor(max));
}



function arvoKone() {  // Picks a random machine from "Koneet" sheet. Returns pin name, abbreviation.
    
  Logger.log("arvoKone");
    
  var vapaatKoneet = [];
  
  for (i = 0; i < koneData.length; i++) {
    if ((koneData[i][4] == true) && (koneData[i][5] != "EI")) {
//      Logger.log("push it");
      vapaatKoneet.push(koneData[i]);
    }
  }
  
  var vapaaLkm = vapaatKoneet.length;
  var random = getRandomInt(vapaaLkm);


  if (vapaaLkm < 1) {
    errorDialog("Ei löytynyt vapaita koneita. 😔");
    //Browser.msgBox("Ei taida olla koneita vapaana...");
    return 0;
  }
  
  var koneNimi = vapaatKoneet[random][2];
  var koneLyhenne = vapaatKoneet[random][3];
  
//  Logger.log(koneNimi + " " + koneLyhenne);
    
  return [koneNimi, koneLyhenne, vapaaLkm];
}

/*
Tarkistaa, onko jompi kumpi pelaaja jo pelannut konetta 2 kertaa.
*/
function check2PlaysRule (lyhenne, countMax) {
  
  Logger.log("check2PlaysRule");

  Logger.log("countMax paramterina saatu:" + countMax);
  
  if (countMax === undefined)
    countMax = 2;
  
  

  if (lyhenne == "undefined") {
    throw("check2PlaysRule ei saanut lyhennettä :(");
  }
  
  var playedMachines = listMachinesByCell(lyhenne);
//  Logger.log("check2PlaysRule() played by x: " + playedMachines[0]);
//  Logger.log("check2PlaysRule() played by y: " + playedMachines[1]);
//  Logger.log(playedMachines[0][0]);
  
  
  Logger.log("playedMachines: " + playedMachines);
  
  
  // Pudotuspelien ja tiebreakien käsittely:
  sheetName = ss.getActiveSheet().getName();
  if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) {  
    var dupes = findDupes(playedMachines);    
    Logger.log("dupes: " + dupes);    
    var count = 0;
      for(var i = 0; i < dupes.length; ++i){
        if(dupes[i] == lyhenne)
          count++;
      }
    Logger.log("check2PlaysRule - count: " + count);
      if (count > 0)
        return "x"; // mennään nyt x:llä tässä sitten...
      else
        return 0;
  }
  
  var dupesX = findDupes(playedMachines[0]);
Logger.log("dupesX: " + dupesX);
  var dupesY = findDupes(playedMachines[1]);
Logger.log("dupesY: " + dupesY);
  
  var countX = 0;
    for(var i = 0; i < dupesX.length; ++i){
    if(dupesX[i] == lyhenne)
      countX++;
  }

  
//DEBUG
Logger.log("***COUNT FOR X: " + countX); //Debug log
/*  
  if (countX > 0)
    Browser.msgBox("***COUNT FOR X: " + countX);
*/  
  var countY = 0;
    for(var i = 0; i < dupesY.length; ++i){
    if(dupesY[i] == lyhenne)
      countY++;
  }  
//DEBUG
Logger.log("***COUNT FOR Y: " + countY); //Debug log
/*  
if (countY > 0)
  Browser.msgBox("***COUNT FOR Y: " + countY);
    */
  
  
//  Browser.msgBox("countX: " + countX + ", countY: " + countY + ", countMax: " + countMax);
  
  // Nyt palauttaa vain ekan osuman (x/y)... pitää muuttaa, jos haluaa pelaajatagin dialogiin mukaan.
  if (countX >= countMax) {
//    Logger.log("dupe in X!"); //Debug log
    return "x";
  }
  else if (countY >= countMax) {
//    Logger.log("dupe in Y!"); //Debug log
    return "y";
  }
  else {
//    Logger.log("no dupes."); //Debug log
    return 0;
  }  
}

  
/*  
Etsii ryhmätaulukosta kummankin pelaajan jo pelaamat koneet.
Palauttaa koneet 2D-arrayssa/taulukossa/mikäliesuomeksi
...paitsi jos on pudotuspeleistä kyse, niin sitten tekee jotain muuta.
*/
function listMachinesByCell(lyhenne) { //Tarvitaanko argumenttia? 
  
  Logger.log("listMachinesByCell");
  
  //var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var cell = ss.getActiveCell();
  var rivi = cell.getRow();
  var sarake = cell.getColumn();  
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var groupLetter; 
  
  
  
// Tehdään, jos ollaan pudotuspelitaulukossa.  
  if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) { 
    
    Logger.log("PUDOTUSPELIT");
    var cellColor = cell.getBackground();
    var machines = [];
    
//Logger.log(cellColor);
//Logger.log(pinCellColors);
    
    if (pinCellColors.indexOf(cellColor) >-1) {
      machines.push(cell.getValue());
      var bg = cellColor;
      var counter = 1;
      while (bg == cellColor) {
        bg = cell.offset(0, -counter).getBackground();
        if (bg == cellColor)
          machines.push(cell.offset(0, -counter).getValue()); 
        Logger.log(bg);
        counter++;
      }
      counter = 1;
      bg = cellColor;
      while (bg == cellColor) {
        bg = cell.offset(0, counter).getBackground();
        if (bg == cellColor)
          machines.push(cell.offset(0, counter).getValue()); 
        Logger.log(bg);
        counter++;
      }      
    } 
  //var dupes = findDupes(machines);  
  Logger.log("machines: " + machines);
  //Logger.log("dupes: " + dupes);
  //throw("joo, kesken on");
  return machines;
  
  }
  

  
  
  
// Tehdään, jos ollaan ryhmätaulukossa...  
  
  if(sheetName.indexOf("RYHMÄ ")>-1) { // Check if we are in a group sheet
    groupSheet = sheet;
    groupLetter = sheetName[6]; // Get group letter (A/B/C/etc. from sheet title. Has to be 7th char. This should be its own function. With more than one letter in group name. And blackjack. And hookers. Eh, forget the function.
    
  }
  else {
    Browser.msgBox("Tämä ei ole ryhmätaulukko...");
    return;
  }
  
  var scoringArea = findScoringArea();  
  var numPlayers = numPlayersInGroup(groupLetter);

  var xPlayerRange = groupSheet.getRange(3, sarake, numPlayers, 1);
  var xPlayerMachines = xPlayerRange.getValues();
  var yPlayerColumn = rivi * 2 - 2;
  var yPlayerRange = groupSheet.getRange(3, yPlayerColumn, numPlayers, 1); // should probably do global vars for sheet margins... or constants even?
  var yPlayerMachines = yPlayerRange.getValues();
  
  //Logger.log("listMachinesByCell - cell: " + cell.getValue()); 
    
  var machines = [];  
  machines[0] = xPlayerMachines;
  machines[1] = yPlayerMachines;
  
  /*
  onEdit() eli manuaalinen koneen syöttö ottaa mukaan myös solussa olevan arvon, korttiPakka() eli konearvonta ei.
  Workaroundina lisätään käsiteltävä kone tässä mukaan, jotta voidaan käsitellä kumpikin tapaus samoin.
  */
  if (caller == "korttiPakka") { // i.e. if caller was korttiPakka()...
    if (lyhenne != "") {
      machines[0].push(lyhenne);
      machines[1].push(lyhenne);
    }
  }  
  return machines; // Back to check2PlaysRule
}


function arrayTo1D (arrToConvert) { // Converts 2D array to 1D array
  
  Logger.log("arrayTo1D");

  /* Trying out conversion from 2D array to 1D array... */
  //var arrToConvert = rowData; // [[0,0,1],[2,3,3],[4,4,5]];
  var newArr = [];


  for(var i = 0; i < arrToConvert.length; i++)
  {
    newArr = newArr.concat(arrToConvert[i]);
  }
  return newArr;
}


// Checks array for 2+ pin entries.
function findDupes(arr) {
     
  Logger.log("findDupes");
  
  // Filter out empties from array.
  /*
  var arrFiltered = arr.filter(function(string) {
    return (string != "");
  });  
  */
  var arrFiltered = filterEmpties(arr);
  
  var arrSorted = arrFiltered.slice().sort(); // Clone and sort filtered array.
  var arr1D = arrayTo1D(arrSorted);  
  var duplicates = [];
  
  // Push duplicates into duplicates[];  
  for (var i = 0; i < arr1D.length - 1; i++) {
    if (arr1D[i + 1] == arr1D[i]) {
      duplicates.push(arr1D[i]);
//    Logger.log("pusheeeeeen"); 
    }
  }
  
//  Logger.log("findDupes - duplicates: " + duplicates);  
  return duplicates; // Back to check2PlaysRule  
}

function filterEmpties(arr) {
    var arrFiltered = arr.filter(function(string) {
    return (string != "");
  });  
  return arrFiltered;
}

/* Laskee rivit sheetistä. */
function countRows(sheetName) {
  Logger.log("countRows");
  row = koneSheetti.getLastRow();
  return row;
}


// Returns array of acceptable row, column coordinates for score marking
function findScoringArea() {
  
  Logger.log("findScoringArea");
  
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var groupLetter;

  var playersInGroup;
  
  if(sheetName.indexOf("RYHMÄ ")>-1) { // Check if we are in a group sheet
    groupLetter = sheetName[6]; // Get group letter (A/B/C/etc. from sheet title)
  }
  else {
    throw("Tämä ei ole ryhmätaulukko.");
  }
  
  playersInGroup = numPlayersInGroup(groupLetter);  

  var greenHillZone = []; // Array of allowable pin/score marking cells.  
  var countteri = -1;    
  
  for (var i = 0; i < playersInGroup; i++) { //row
  countteri = countteri + 1;  
  for (var j = countteri + i; j < (playersInGroup * 2 - 2); j++) { //column
    greenHillZone.push({ row: (i + 3), column: (j + 6) });
  }
 }
//  Logger.log(greenHillZone);
  
  //Browser.msgBox("sheet: " + sheet + " sheetName:" + sheetName + " groupLetter: " + groupLetter + " playersInGroup: " + playersInGroup );
  return greenHillZone;
}



function getGroupLetter() {
  
  Logger.log("getGroupLetter");
  
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var groupLetter;
  
  
  //getColumnsToPlayers ()
  
  if (sheetName.indexOf("RYHMÄ ")>-1) { // Check if we are in a group sheet
    groupLetter = sheetName[6]; // Get group letter (A/B/C/etc. from sheet title)
  }
  else   if ((sheetName.indexOf("Finaali") > -1) || (sheetName.indexOf("Tiebreak") > -1)) { 
    var i = 1;
    var bg = cell.getBackground();
    while (pinCellColors.indexOf(bg) >-1) {
      bg = cell.offset(0, -i).getBackground();
      i++;
    }
    groupLetter = cell.offset(0, getColumnsToPlayers()).getValue();
  }
  //  groupLetter = "getGroupLetter";
  else
    throw("Tämä ei ole ryhmätaulukko... -_-");
Logger.log("getGroupLetter luulee ryhmän nimeksi: " + groupLetter);  
  return groupLetter;

}

function isInsideScoringArea(row, column) {
  
  Logger.log("isInsideScoringArea");
  //throw(row + " " + column);

  area = findScoringArea();    
  for (var i = 0; i < area.length; i++) {
    if ((area[i].row == row) && (area[i].column == column))
      return "true";
  }
  return "false";
}

/*
Returns number of players in a group by letter, e.g. group "A" has n players.
...saattaa hajota, jos on ryhmätoiveita samassa solussa...
*/

function numPlayersInGroup(groupLetter) {
  Logger.log("numPlayersInGroup");  
  
//DEBUG:
//groupLetter = "b";
    
  
  /* 
  //Vanha metodi, joka laski ryhmäkirjaimen lkm:n Pelaajat-välilehden B-sarakkeesta.
   
  var totalNumPlayers = firstEmptyRowInColumn("Pelaajat", 2);
  var allPlayers = playerSheet.getRange(3, 2, totalNumPlayers).getValues();
  var playersInGroup = 0;
  
  for (i = 0; i < totalNumPlayers; i++) {
    if (allPlayers[i].toString().indexOf(groupLetter)>-1)
      playersInGroup++;
  }  
  return playersInGroup;
  */
  
//Browser.msgBox(groupLetter);
  
  
  
  var groupSheet
  
  if (groupLetter == "X")
    groupSheet = ss.getSheetByName("RYHMÄ X aka Jatkopelit");
  else  
    groupSheet = ss.getSheetByName("RYHMÄ " + groupLetter.toUpperCase());
  
  
//Browser.msgBox(groupSheet.getName());
  
  var playersInGroup = groupSheet.getRange("A1").getValue();
  return playersInGroup;

//Browser.msgBox(groupSheet.getName());
  
}

function firstEmptyRowInColumn(sheetName, column) { // Takes name of sheet as string and column as number, returns first empty row in column.
  
  Logger.log("firstEmptyRowInColumn");
  
  //Logger.log("TÄMÄ VIE PIRUSTI AIKAA?");
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var lr=playerSheet.getLastRow()
  var columntest = playerSheet.getRange(1, column, lr).getValues();  
  
    for(var i=0; i<columntest.length ; i++){
      ////Logger.log (columntest[i]);
      if (columntest[i] ==""){
        var rowOfEmptyCell = playerSheet.getRange(i + 1, column).getRow();
        return rowOfEmptyCell;
    }
  }
}

/* Arvontadialogin (html-version) "Peruuta"-nappi */
function peruutaMatsi(callback) {
  
Logger.log("peruutaMatsi");
  
  var solu = ss.getActiveCell();
  if (solu.getColumn() % 2 == 1) {
    solu = ss.getActiveCell().offset(0, -1)
    solu.activateAsCurrentCell();
  }
  cell = solu; // global var
  
    
  var rivi = solu.getRow();
  var sarake = solu.getColumn();
  var kone = solu.getValue();
  
  var pelaajat = findPlayerPair(rivi, sarake);
  var p1Tag = pelaajat[2];
  var p2Tag = pelaajat[3];  
  
  var matsinRivi = checkForMatch(p1Tag, p2Tag, kone, koneSheetti);
  
  endMatch(matsinRivi, function() { Logger.log ("peruutaMatsi CALLBACK"); });
  
//  Logger.log("peruuta pelaajat: "  + pelaajat);
//  Logger.log("peruuta kone: "  + kone);
//  Logger.log("peruuta solu: " + solu.getA1Notation());
  solu.clearContent();
//  Logger.log("peruutus tehty");  
    //pinCellProtectOff();
  
  if (typeof(callback) != "undefined")
    callback();
  
} 


// wait, how do i pass variables to html here...
function errorDialog(dialogi, callback) {
    Logger.log("errorDialog");
  var htmlNimi = 'Error';
  var html = doGet(htmlNimi, dialogi); //HtmlService.createHtmlOutputFromFile('Error');
//  Logger.log("openDialog() about to do call ss.getUi  with dialogi: " + dialogi);
  html.setWidth(360);
  html.setHeight(160);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, "Oho...");
//  Logger.log("openDialog() finished.");
  if (typeof(callback) != "undefined")
    callback();
  
}



/*Pops up a nice html modal dialog, which doesn't pause the script like Browser.msgBox()...*/
//tätä geneerisemmäksi...
function matchDialog(dialogi) { 
  Logger.log("openDialog");
  var htmlNimi = 'Matsi';
  var html = doGet(htmlNimi, dialogi); //HtmlService.createHtmlOutputFromFile('Index');
  html.setWidth(480);
  html.setHeight(240);
//  Logger.log("openDialog() about to do call ss.getUi  with dialogi: " + dialogi);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, "Konearvonta...");
//  Logger.log("openDialog() finished.");
}

function doGet(htmlNimi, dialogi) { // used by openDialog
//  Logger.log("doGet htmlNimi: " + htmlNimi);
  var t = HtmlService.createTemplateFromFile(htmlNimi);
//  Logger.log("doGet t: " + t);
  if (typeof(dialogi) != "undefined")
    t.data = dialogi;
  else
    t.data = "heippa moi.";
  return t.evaluate();
}

/* //old version before passing variables...
function doGet(htmlNimi) { // used by openDialog
  Logger.log("doGet");
  return HtmlService
      .createTemplateFromFile(htmlNimi)
      .evaluate();
//  Logger.log("doGet is finishing.");
}
*/

function testLog() {
  var d = new Date();
  var timeStamp = d.getTime();
  Logger.log("testLog - time: " + timeStamp);
}

function testiDialogi() {
     var sivu =  HtmlService
      .createTemplateFromFile('Testi')
      .evaluate();
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(sivu, 'Herja');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function trimUnusedPins() {  
  
  Logger.log("trimUnusedPins");
  var matchesKoneet = getMatchesFromKoneet();
  var matchesGroup = [];
  var groupLetters = [];
  var openGames = [];
  var trimmablePins = [];
  var myRegxp = /^([A-Z]){1}$/; // Rajoitetaan ryhmän nimi 1 isoon kirjaimeen. Melkein ymmärrän RegExiä.
    
  for (var i = 0; i < matchesKoneet.length; i++) {
    if(myRegxp.test(matchesKoneet[i][5]) == true)
      groupLetters.push(matchesKoneet[i][5]);
  }
  groupLetters = filterEmpties(groupLetters);
  groupLetters = groupLetters.filter( onlyUnique );
  groupLetters.sort();

  var letter;
  // Luuppaa ryhmäkirjaimet
  for (var i = 0; i < groupLetters.length; i++) {
    letter = groupLetters[i];
    matchesGroup = getMatchesFromGroup(letter);
    openGames = findOpenGames(matchesGroup);
    // Luuppaa koneet
    for (var j = 0; j < matchesKoneet.length; j++) {
      if (matchesKoneet[j][5] == letter) {
        if (openGames.indexOf(matchesKoneet[j][0]) < 0) {
          trimmablePins.push(matchesKoneet[j][0]);      
      }      
     }
    }    
   }
  
  for (var i = 0; i < trimmablePins.length; i++) {
    var rivi = findRow(koneSheetti, 3, trimmablePins[i]);
    console.log("trimUnusedPins vapautti varatuksi jääneen koneen riviltä: " + rivi);
    console.log("rivin " + rivi +" sisältö: " + koneSheetti.getRange(rivi, 3, 1, 7).getValues());
    endMatch(rivi);
  }
}

function getMatchesFromKoneet() {
  Logger.log("getMatchesFromKoneet");
  var koneLkm = countRows("Koneet") - 1;
//  Logger.log("koneLkm: " + koneLkm);
  var range = koneSheetti.getRange(2, 4, koneLkm, 6);
  var matches = range.getValues();
/*
  Logger.log("getMatchesFromKoneet - matches: " + matches);
  Logger.log("getMatchesFromKoneet - matches.length: " + matches.length);
  Logger.log("getMatchesFromKoneet - matches[0]: " + matches[0]);  
  Logger.log("getMatchesFromKoneet - matches[0][3]: " + matches[0][3]);  
*/  
  return matches;
}

function getMatchesFromGroup(letter) {
  Logger.log("getMatchesFromGroup");
  Logger.log("letter: " + letter);
  
Browser.msgBox(letter);
  
  //DEBUG VAR
  //letter = 'A';
  var numPlayers = numPlayersInGroup(letter);
  var sheetName = "RYHMÄ " + letter;
  var sheet;
  if (letter == "X")
    sheet = ss.getSheetByName("RYHMÄ X aka Jatkopelit");
  else
    sheet = ss.getSheetByName(sheetName);
  var groupData = sheet.getRange(2, 4, numPlayers + 1, numPlayers * 2).getValues();
  //Logger.log(groupData);
  //Logger.log(sheetName + " " + numPlayers);
  //return matches;
  return groupData;  
}


//returns open games from group sheet
function findOpenGames(groupData) {  
  var openGames = [];
  
  var i;
  var j;  
  for (i = 1; i < groupData.length; i++) {
    for (j = 0; j < (groupData.length - 1) * 2; j++) {
    //Logger.log(groupData[i][j]);
      
      if ((typeof(groupData[i][j]) == "string") && (typeof(groupData[i][j+1]) != "number"))
        openGames.push(groupData[i][j]);      
    }    
  }  
  openGames = filterEmpties(openGames);
  openGames = openGames.filter( onlyUnique );
  openGames.sort();  
//  Logger.log("findOpenGames - openGames: " + openGames);  
  return openGames;  
}

function onlyUnique(value, index, self) { 
    //usage example: var uniqueDupes = dupes.filter(onlyUnique);
    return self.indexOf(value) === index;
}

function pinCellProtectOn(sheet, cell) {  
  
  Logger.log("pinCellProtectOn");
  
  //var ss = SpreadsheetApp.getActive();
  //cell = ss.getActiveCell();    

  if (typeof(range) == "undefined")
    range = sheet.getRange(cell.getA1Notation());
  var protection = range.protect().setDescription('konesolu');
  protection.setWarningOnly(true);
  //protection.setDomainEdit(false);
}

function pinCellProtectOff() {
  
  Logger.log("pinCellProtectOff");
  
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  
  for (var i = 0; i < protections.length; i++) {
//    Logger.log(protections[i].getDescription());
    if (protections[i].getDescription() == 'konesolu') {
      protections[i].remove();
    }
  }
}

function clearGroupSheet() {
  
  Logger.log("clearGroupSheet");
  
  var greenHillZone = findScoringArea();
  sheet = ss.getActiveSheet();
  
  /*Logger.log(sheetName.indexOf("RYHMÄ "));
  
  if(sheetName.indexOf("RYHMÄ ") < 1) {
    Logger.log("is not group sheet");
    return;
  }
  */
  
  for (var i = 0; i < greenHillZone.length; i++) {
    var row = greenHillZone[i].row;
    var column = greenHillZone[i].column;
    
    sheet.getRange(row, column).clearContent();
  }
  trimUnusedPins();
}

function fillGroupSheet() {
  
  koneDataRange = koneSheetti.getDataRange();
  koneData = koneDataRange.getValues();
  var kone;
  var koneLyh;
  
  var greenHillZone = findScoringArea();
  
  //Logger.log(greenHillZone);
  
  sheet = ss.getActiveSheet();
  
  for (var i = 0; i < greenHillZone.length; i++) {
    var row = greenHillZone[i].row;
    var column = greenHillZone[i].column;
    
    if (column % 2 == 0) {
      kone = arvoKone();
      koneLyh = kone[1];
      sheet.getRange(row, column).setValue(koneLyh);
    }
    else
      sheet.getRange(row, column).setValue(getRandomInt(2));
  }
  trimUnusedPins();  
}

function clearKoneet() {
  //DEBUG-funktio
  var range = "F2:I35";
  koneSheetti.getRange(range).clearContent();
}

function moveRight () {
  cell = ss.getActiveCell();
  Logger.log("WORKAROUND");
  Logger.log("CELL: " + cell.getA1Notation());
  cell.offset(0, 1).activate();
}

function moveLeft () {
  cell = ss.getActiveCell();
  Logger.log("WORKAROUND");
  Logger.log("CELL: " + cell.getA1Notation());
  cell.offset(0, -1).activate();
}

function showCoordinates(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var x = sheet.getActiveCell().getColumn();
  var y = sheet.getActiveCell().getRow();
  Browser.msgBox("col: " + x + ", row: " + y);
}

function dumben() {
  /*
  "Tyhmentää" ryhmäkaaviot, eli poistaa kaavat pelaajien nimistä ja tageista ja korvaa ne pelkillä arvoilla.
  Tämä siksi, että IFPAn ranking-palvelin ei aina jaksa pysyä pystyssä, jolloin kaaviosta tyhjeneekin kaikki nimet...
  */
  
  /*
  var cell = ss.getActiveCell();
  var val = cell.getValue();
  cell.setValue(val);
  */
  
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
     'Oletko varma?',
     'Onko tullut tyhmennyksen aika?',
      ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    
    var groupLetters = ["A", "B", "C", "D"];
  
  //var sheet = ss.getActiveSheet();
  //var sheetName = sheet.getName();
  
  //Browser.msgBox(groupLetters[1]);  
  
    for (i = 0; i < groupLetters.length; i++) {
  
      sheet = ss.getSheetByName("RYHMÄ " + groupLetters[i]);  
      var nimiAlue = sheet.getRange("B3:B18");
      var nimet = nimiAlue.getValues();
      nimiAlue.setValues(nimet);
      var tagiAlue = sheet.getRange("D2:AH2");
      var tagit = tagiAlue.getValues();
                  
      //Poistetaan ylimääräiset merkit, käytännössä kaarisulut lyhyistä tageista, esim. "TV)" -> "TV"
      for (j = 0; j < tagit[0].length; j++) {
        tagit[0][j] = tagit[0][j].replace(/[^a-zA-Z0-9\s_-]/g,'');
      }
            
      //tagit[0][14] = tagit[0][4].replace(/[^a-zA-Z0-9_-]/g,'');
      
      
      //Browser.msgBox(tagit[0]);
      
      tagiAlue.setValues(tagit);
      
      
      
    
  }    
    
    ui.alert('"Would anybody tell me if I was gettin\' stupider?"');
  } else {
    return;
  }  
  //Browser.msgBox("Would anybody tell me if I was getting stupider?");
}

function playerVsPlayer() {
  var sheet = ss.getActiveSheet();
  var group = sheet.getRange("B3").getValue().toUpperCase();
  var groupSheet
  
  if (group == "X")
    groupSheet = ss.getSheetByName("RYHMÄ X aka Jatkopelit");
  else
    groupSheet = ss.getSheetByName("RYHMÄ " + group);
  var ruudukko = ss.getRangeByName("ruudukko");
  //Browser.msgBox(ruudukko.getA1Notation());
  var tagit = ss.getRangeByName("tagit").getValues();
  tagit = tagit.filter(String);
  var numTags = tagit.length;  
  var ekaRivi = ruudukko.getRow();
  var ekaSarake = ruudukko.getColumn();
  var vikaRivi = ekaRivi + numTags - 1;
  var vikaSarake = ekaSarake + numTags - 1;
  
  
  var groupTags = groupSheet.getRange(2, 4, 1, 32).getValues();
  var groupData = groupSheet.getRange("A1:AI18").getValues();
  
  var group1D = arrayTo1D(groupTags);
  
  var p1Tag;
  var p2Tag;
  var indexP1;
  var indexP2;
  var offset;
  var piste;
  
  ruudukko.clearContent();
  
  
  //Browser.msgBox("numTags: " + numTags + ", vikaRivi: " + vikaRivi + ", vikaSarake: " + vikaSarake);  
  //var tagit2 = tagit.filter(String);
  //Browser.msgBox(tagit2.length);
  
  for (i = ekaRivi; i <= vikaRivi; i++) {
Logger.log("i: " + i);
    p1Tag = tagit[i-ekaRivi].toString();
    p1Tag = p1Tag.toUpperCase();

    
    
    for (j = ekaSarake; j <= vikaSarake; j++) {
Logger.log("j: " + j);
      p2Tag = tagit[j-ekaSarake].toString();
      p2Tag = p2Tag.toUpperCase();

      
Logger.log(sheet.getRange(i, j).getA1Notation());
Logger.log("p1Tag: " + p1Tag);
Logger.log("p2Tag: " + p2Tag);
      
      if (p1Tag != p2Tag) {
        //sheet.getRange(i, j).setValue("LOL");
        
        //Browser.msgBox(group1D);
        
        indexP1 = group1D.indexOf(p1Tag) + 4;
        indexP2 = group1D.indexOf(p2Tag) + 4;
        //offset = (indexP2 - indexP1) / 2;
        offset = (indexP2 - 2) / 2;

        
//Logger.log("OFFSET: " + offset);
/*
        if (offset < 0) {
          
          Logger.log("CONVERTING OFFSET");
          offset = (16 + offset);
        }
        Logger.log("OFFSET AFTER CONVERSION: " + offset);
*/                      
        
        //Logger.log("offset: " + (offset) + " indexP1: " + indexP1);
        
      try {
        piste = groupData[offset+1][indexP1];
        if (piste === "") {
          //Browser.msgBox(piste);
          piste = 0;
        }
        else
          piste = 1-piste;
      }
      catch(err) {
        throw("Virhe - typo TAGissa?" + err);
      }
        
        

        
        //Logger.log("indexP1: " + indexP1 + ", indexP2: " + indexP2 + ", offset: " + offset);
        //Logger.log("piste: " + groupData[offset+1][indexP1]);
        sheet.getRange(i, j).setValue(piste);
        
        
      }      
    }
  }
}



function resolveTies(score) {
  
}


/*This is the end... */