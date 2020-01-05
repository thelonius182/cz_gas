/* 
------------------------------------------------------------------------------------------
Tools voor Roosters 4.0
versie: 26 nov 2018 19:20, LvdA

Legenda: (g)o = (global) object
         (g)a = (global) array
         s = string, r = range, i = integer
         d = date, b = boolean, sh = sheet
         c = constant, p = parameter
celverwijzingen:
   in sheets: 1-based (row en col)
   in arrays: 0-based (rowIdx en colIdx)
------------------------------------------------------------------------------------------
*/

// Global variables
var goSshRoosters = SpreadsheetApp.getActiveSpreadsheet();
var gsMailTo;
var gdToday;
var gdStartDtm;
var gdStartDtm_ymd;
var gdStopDtm;
var gc_EEN_DAG_IN_MS;
var gaMtrType_k;
var gaMtrType_v;
var gaHostReport;
var giEditKol_min = goSshRoosters.getRangeByName('mtr_colhdr').getValues()[0].indexOf('Type') + 1;
var giEditKol_max = goSshRoosters.getRangeByName('mtr_colhdr').getValues()[0].indexOf('Datum') + 1;
var giEditKol_psr = goSshRoosters.getRangeByName('psr_colhdr').getValues()[0].indexOf('Def') + 1;
var giKolBijz = giEditKol_max + 1;
var gsTypeHU_regex = /^(h|u)$/i;
var gaDockRefs_m2p = goSshRoosters.getRangeByName("mtr_dockrefs").getValues();
var gaDockRefs_p2m = goSshRoosters.getRangeByName("psr_dockrefs").getValues();
var grPres_hum = goSshRoosters.getRangeByName("psr_hum"); // vlag: er is herh/upl/mont in montagerooster
  
/* ==========================================================================================
 * Bij openen spreadsheet: plaats CZ-menu
 */
function onOpen() {
  init();
  SpreadsheetApp.getUi().createMenu('CZ-menu')
      .addItem('Sync Mt-rooster vanaf ' + gdStartDtm_ymd, 'czSync')
      .addItem('Roosters indexeren', 'czDockSchedules')
      .addToUi();
}

/* ==========================================================================================
 * Global init
 */
function init() {
  gc_EEN_UUR_IN_MS = 60 * 60 * 1000;
  gc_EEN_DAG_IN_MS = 60 * 60 * 1000 * 24;
  gdStartDtm = new Date();
  gdStartDtm.setHours(0, 0, 0);
  gdStartDtm_ymd = Utilities.formatDate(gdStartDtm, "Europe/Amsterdam", "yyyy-MM-dd");
  var iStartDtm_ms = gdStartDtm.getTime(); 
  gdStartDtm = new Date(iStartDtm_ms - gc_EEN_UUR_IN_MS);

}

/* ==========================================================================================
 * Voeg een rij toe aan het montagerooster
 *
 */
function czAddRowsToMontagerooster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('montage');
  sheet.insertRowsAfter(sheet.getLastRow(), 65);
}


/* ==========================================================================================
 * Create date
 */
function czDtm(piDtm) { // bv '20181102'
  var sDtm = piDtm.toString();
  var czDateParser = /(\d{4})(\d{2})(\d{2})/;
  var match = sDtm.match(czDateParser);
  var dResult = new Date(
    match[1],     // year
    match[2] - 1, // monthIndex
    match[3],     // day
    0,            // hours
    0,            // minutes
    0             // seconds
  );
  return dResult;
}

/* ==========================================================================================
 * Zorg dat uitzendingtype in het montagerooster (Mtr) overeenkomt met het presentatierooster
 * en plaats de feedback uit het montagerooster in het presentatierooster
 */
function czSync() {
  init();

  var rMtrDtms = goSshRoosters.getRangeByName('mtr_datum');
  var aMtrDtms = rMtrDtms.getValues()
                         .filter(function(v1) {
                           return v1 > "";
                         });
  
  // laatste sync_datum = laatste Mtr_datum
  gdStopDtm = czDtm(aMtrDtms[aMtrDtms.length - 1]);

  var aMtrSlots = goSshRoosters
  .getRangeByName('Mtr_uren').getDisplayValues()
  .map(function(v1, i1) {
    return aMtrDtms[i1] + "_" + v1[0].slice(0, 2);
  });

  var aMtrType = goSshRoosters.getRangeByName('mtr_type').getValues();
  var aPresDagDtmUren = goSshRoosters.getRangeByName("dagDtmUren").getValues();
  var aPresTypes = goSshRoosters.getRangeByName("Live_SemiLive").getValues();
  var aPresSlots = [];
  var aPresSlotTypes = [];
  var aZondagen = [];
  
  
  for (var presNr = 1; presNr < aPresDagDtmUren.length; presNr += 1) {
    var dPresDtm = aPresDagDtmUren[presNr][1];
    
    // neem alleen dat deel vh presentatierooster dat strookt met het montagerooster
    // vanaf vandaag. 
    if (dPresDtm < gdStartDtm) {
      continue;
    }
    
    if (dPresDtm > gdStopDtm) { 
      break;
    }
    
    var sPresDtm_ymd = Utilities.formatDate(dPresDtm, "Europe/Amsterdam", "yyyyMMdd");
    
    var sPresUren_van = aPresDagDtmUren[presNr][5].slice(0, 2);
    var sPresUren_tem = aPresDagDtmUren[presNr][5].slice(8, 10);
    var iPresUren_van = parseInt(sPresUren_van, 10);
    var iPresUren_tem = parseInt(sPresUren_tem, 10);
    
    for (var u1 = iPresUren_van; u1 < iPresUren_tem; u1 += 1) {
      var sSlot = sPresDtm_ymd + "_" + Utilities.formatString("%02d", u1);
      aPresSlots.push(sSlot);
      aPresSlotTypes.push(aPresTypes[presNr]);
    }
  }
  
  // Sensenta en X-rated erbij
  var caLive = ["Live"];
  var iStartDtm_ms = gdStartDtm.getTime(); 
  var iStopDtm_ms = gdStopDtm.getTime(); 
    
  for (var z1 = iStartDtm_ms; z1 < iStopDtm_ms; z1 += gc_EEN_DAG_IN_MS) {
    var dDag1 = new Date(z1);
    
    if (dDag1.getDay() === 0) { // zondag
      
      var sPresDtm_ymd = Utilities.formatDate(dDag1, "Europe/Amsterdam", "yyyyMMdd");
      
      var sSlot = sPresDtm_ymd + "_19"; // Sensenta
      aPresSlots.push(sSlot);
      aPresSlotTypes.push(caLive);
      
      sSlot = sPresDtm_ymd + "_21"; // X-rated
      aPresSlots.push(sSlot);
      aPresSlotTypes.push(caLive);
    }
  }

  aPresSlots.forEach(function(pres_slot, s1) {
    var sPresType = aPresSlotTypes[s1][0]; 
    var sMtrType_soll;
    
    if (sPresType === 'Live') {
      sMtrType_soll = 'Live'; 
    } else if (sPresType === 'Semi-live') {
      sMtrType_soll = 'm'; 
    } else { // leeg
      sMtrType_soll = '?'; 
    }
    
    var iMtrSlotIdx = aMtrSlots.indexOf(pres_slot);
    
    if (iMtrSlotIdx !== -1) { // corresponderend mtr-slot gevonden
      var sMtrType = aMtrType[iMtrSlotIdx][0];
      
      if (gsTypeHU_regex.test(sMtrType)) { // type h or u
        aMtrType[iMtrSlotIdx] = [sMtrType];
      } else {
        aMtrType[iMtrSlotIdx] = [sMtrType_soll];  
      }
    }
  });

  // montagerooster bijwerken
  goSshRoosters.getRangeByName('mtr_type').setValues(aMtrType);
  
  // niet-live slots terugmelden aan presentatierooster.
  putFeedback();
}


/* ==========================================================================================
 * feedback MR > PR in batch
 * 
 * In het presentatierooster is te zien of het montagerooster voor dat blok slots heeft die 
 * niet live gaan. De achtergrond van de kolom "Def" in dat blok is dan grijs. De besturing
 * van die achtergrondkleur is een vlag: de variabele aPres_hum. 1 = grijs, 0 = wit.
 * De vlag wordt gezet als minstens 1 vd slots die tot het blok behoren een herhaling of
 * upload is, of minstens 1 van de variabelen Pres/Tech/Dtm is ingevuld.
 *
 * Het script verwerkt alleen de blokken die stroken met het montagerooster vanaf vandaag.
 * Verwerking van het hele rooster duurt te lang: de Google-server staat max. 6 minuten toe.
 *
 */
function putFeedback() {
  
  init();
  
  var rMtrTypes = goSshRoosters.getRangeByName('mtr_type');
  var rMtrDtms = goSshRoosters.getRangeByName('mtr_datum');
  var aPsrDockRefs = goSshRoosters.getRangeByName('psr_dockrefs').getValues();
  var aPresDagDtmUren = goSshRoosters.getRangeByName("dagDtmUren").getValues();
  
  // laatste feedbackdatum = laatste Mtr_datum
  var aMtrDtms = rMtrDtms.getValues()
                         .filter(function(v1) {
                           return v1 > "";
                         });
  
  // laatste sync_datum = laatste Mtr_datum
  // chg 2019-12-14, nav moro 20200102: beethoven op do-ochtend niet meenemen
  // gdStopDtm = czDtm(aMtrDtms[aMtrDtms.length - 1]);
  // chg 2020-01-05, operatie Missa: "beethoven" niet meer in MT-rooster
  // gdStopDtm = czDtm(aMtrDtms[aMtrDtms.length - 2]);
  gdStopDtm = czDtm(aMtrDtms[aMtrDtms.length - 1]);
  
  // column header rPres_hum 
  var aPres_hum = [["hum"]];
  
  
  // doorloop de presentatieblokken
  for (var presNr = 1; presNr < aPresDagDtmUren.length; presNr += 1) {
  
    // neem alleen die blokken die stroken met het montagerooster vanaf vandaag. 
    var dPresDtm = aPresDagDtmUren[presNr][1];
    
    if (dPresDtm < gdStartDtm || dPresDtm > gdStopDtm) {
      aPres_hum.push([0]);
      continue;
    }
    
    var psr_ref = aPsrDockRefs[presNr];
    Logger.log('psr_ref = ' + psr_ref);
    var sRef = psr_ref[0].slice(1).split("|")[0];
    var iRef = parseInt(sRef, 10);
    var mtrCell = rMtrTypes.getCell(iRef, 1);
    var iMtrColumn = mtrCell.getColumn();
    var iMtrDelta = giEditKol_min - iMtrColumn + 1;
    var iDockRef_m2p = mtrCell.offset(0, iMtrDelta - 8).getValue();
    
    var bSemilive = getMtrState(mtrCell, iDockRef_m2p, iMtrDelta);
    
    if (bSemilive) {
      aPres_hum.push([1]);
    } else {
      aPres_hum.push([0]);
    }
  }
  
  // presentatierooster bijwerken
  goSshRoosters.getRangeByName('psr_hum').setValues(aPres_hum);
}

/* ==========================================================================================
 * Zitten er semi-live onderdelen in de slots van dit blok?
 */
function getMtrState(piCell, piDockRef_m2p, piDelta) {
  var aMtrRowsFromPsr = gaDockRefs_p2m[piDockRef_m2p - 1][0]
  .slice(1)
  .split("|")
  .map(function(v1){
    return parseInt(v1, 10);
  });
  var iCellRow = piCell.getRow();
  var iRowOffset = aMtrRowsFromPsr[0] - iCellRow;
  var bHUPTD = false; // init: er is geen sprake van Herh of Upl; Pres/Tech/Datum leeg
  
  for (var c1 = 0; c1 < aMtrRowsFromPsr.length; c1 += 1) {
    var sPTD = piCell.offset(iRowOffset + c1, piDelta).getValue() +
               piCell.offset(iRowOffset + c1, piDelta + 1).getValue() +
               piCell.offset(iRowOffset + c1, piDelta + 2).getValue();
    var sH_of_U = piCell.offset(iRowOffset + c1, piDelta - 1).getValue();
    
    if (sPTD || gsTypeHU_regex.test(sH_of_U)) {
      bHUPTD = true;
      break;
    }
  }
  return bHUPTD;
}

/* ==========================================================================================
 * Koppel montageslots en presentatieblokken
 * en maak de presentatieblokken zichtbaar in het montagerooster
 * door de uren vet te maken.
 */
function czDockSchedules() {
  init();

  var rMtrDtms = goSshRoosters.getRangeByName('mtr_datum');
  var aMtrDtms = rMtrDtms.getValues();

  var aMtrSlots = goSshRoosters
  .getRangeByName('Mtr_uren').getDisplayValues()
  .map(function(v1, i1) {
    return aMtrDtms[i1] + "_" + v1[0].slice(0, 2);
  });

  grDockRefs_m2p = goSshRoosters.getRangeByName("mtr_dockrefs");
  grDockRefs_p2m = goSshRoosters.getRangeByName("psr_dockrefs");

  // reset dockrefs in montagerooster
  gaDockRefs_m2p = grDockRefs_m2p.getValues()
  .map(function(v1, i1) {
    var result;
    if(i1 === 0) { // column header laten staan
      result = [v1];
    } else {
      result = [0];
    }
    return result;
  });
  
  // reset dockrefs in presentatierooster
  gaDockRefs_p2m = grDockRefs_p2m.getValues()
  .map(function(v1, i1) {
    var result;
    if(i1 === 0) { // column header laten staan
      result = [v1];
    } else {
      result = [""];
    }
    return result;
  });
  
  var aPresSlots = [];
  var aPresRows = [];
  var aPresDagDtmUren = goSshRoosters.getRangeByName("dagDtmUren").getValues();

  
  for (var presNr = 1; presNr < aPresDagDtmUren.length; presNr += 1) {
    var dPresDtm = aPresDagDtmUren[presNr][1];
    var sPresDtm_ymd = Utilities.formatDate(dPresDtm, "Europe/Amsterdam", "yyyyMMdd");
    var sPresUren_van = aPresDagDtmUren[presNr][5].slice(0, 2);
    var sPresUren_tem = aPresDagDtmUren[presNr][5].slice(8, 10);
    var iPresUren_van = parseInt(sPresUren_van, 10);
    var iPresUren_tem = parseInt(sPresUren_tem, 10);
    
    for (var u1 = iPresUren_van; u1 < iPresUren_tem; u1 += 1) {
      var sSlot = sPresDtm_ymd + "_" + Utilities.formatString("%02d", u1);
      aPresSlots.push(sSlot);
      aPresRows.push(presNr + 1);
    }
  }
  
  aMtrSlots.forEach(function(mtr_slot, m1) {
    var iPresSlotsIdx = aPresSlots.indexOf(mtr_slot);
    
    if (iPresSlotsIdx !== -1) { // slot in blok
      var iPresRow = aPresRows[iPresSlotsIdx];
      gaDockRefs_m2p[m1] = [iPresRow];
      var m2 = m1 + 1;
      gaDockRefs_p2m[iPresRow - 1] = [gaDockRefs_p2m[iPresRow - 1][0] + "|" + m2];
    }
  });

  // spreadsheet bijwerken
  grDockRefs_m2p.setValues(gaDockRefs_m2p);
  grDockRefs_p2m.setValues(gaDockRefs_p2m);
  
  // PR-uren in MR vet maken
  var aFontWeights = [];
  
  gaDockRefs_m2p.forEach(function(ref, r1) {
    if (ref[0] === 0 || r1 === 0) {
      aFontWeights.push(["normal"]);
    } else {
      aFontWeights.push(["bold"]);
    }
  });
  
  goSshRoosters.getRangeByName('Mtr_uren').setFontWeights(aFontWeights);
}

