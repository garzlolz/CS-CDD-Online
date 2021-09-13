var ss = SpreadsheetApp.getActiveSpreadsheet();

function UpdateNewVPO2() {
  UpdateNewVPO3("Vendor","Work1",8,18);
  UpdateNewVPO3("Vendor_Ext","Work2",8,18);
  UpdateNewVPO3("FC_RM_WH","Work3",8,18, 1);  
  //Maintenance("Vendor");
  //Maintenance("Vendor_Ext");
}

function UpdateMin() {
  //UpdateNewVPO3("Vendor","Work1",9,18);
  //UpdateNewVPO3("Vendor_Ext","Work2",9,18);
  var timenow = new Date();
  if ( timenow.getHours()==6 && timenow.getMinutes()==0 ) {
      ss.getSheetByName("Config").getRange("G2").setValue("1");
  }
  
  Maintenance("Vendor");
  Maintenance("Vendor_Ext");
  Maintenance("FC_RM_WH");

  
}

function test() {
  Protect('Vendor_Ext','M:M', 2);
}
function Maintenance(sheet_Vendor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName(sheet_Vendor);
  var sheet1 = ss.getSheetByName("Config");
  var Lockdown = sheet1.getRange("H2").getValue().split(":");
  var Locked = sheet1.getRange("I2").getValue();
  //Logger.log(Lockdown[0]);
  
  var timenow = new Date();
  var timehour = timenow.getHours();
  var starttime = new Date( timenow.getFullYear(), timenow.getMonth(), timenow.getDate(), Lockdown[0], Lockdown[1],0);
  var endtime = new Date(starttime.getTime()+1200000); // 20mins delay
  
  
  switch (true) {
    case ( Lockdown == "" && Locked == ""):
      break;
      
    case (Lockdown != "" && Locked == "" && timenow <=endtime && timenow >=starttime ):
      LockDown();
      sheet1.getRange("I2").setValue("1");    // set Locked flag
      break;
    case ((Lockdown != "" && Locked == "1") && ( timenow >endtime || timenow < starttime)): // lockdown expired
      UnLockDown()
      sheet1.getRange("G2").setValue("1");
      sheet1.getRange("H2").setValue("");
      sheet1.getRange("I2").setValue("");
      break;
    case ( Lockdown == "" && Locked == "1" ): // invalid lockdown
      UnLockDown()
      sheet1.getRange("G2").setValue("");
      sheet1.getRange("H2").setValue("");
      sheet1.getRange("I2").setValue("");
      break;
  }
  
}

function LockDown() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Vendor'); // or whatever is the name of the sheet 
  var range = sheet.getRange('X:Y');
  var rule = SpreadsheetApp.newDataValidation()
     .requireTextEqualTo('Biggio')
     .setAllowInvalid(false)
     .setHelpText('We are working on importing new CDD change, Page locked down. Please allow us some mins. Thanks')
     .build();
  range.setDataValidation(rule);
  Protect('Vendor_Ext','M:M', 1);
  Protect('FC_RM_WH','M:M', 1);
}


function UnLockDown()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Vendor'); 
  sheet.getRange('X:AA').clearDataValidations(); 
  
  var range = sheet.getRange('X:X');
  range.setNumberFormat('d"-"mmm"-"yyyy');
  var yearnow = new Date();
  var yearnow = yearnow.getFullYear();
  //Logger.log(yearnow);
  
  var rule = SpreadsheetApp.newDataValidation()
     .requireDateBetween(new Date('1/1/' + yearnow), new Date('12/31/' + (yearnow+2)))
     .setAllowInvalid(false)
     .setHelpText('Wrong Date format, please try again.')
     .build();
  range.setDataValidation(rule);
  Protect('Vendor_Ext','M:M', 2);
  Protect('FC_RM_WH','M:M', 2);
  
}


function UpdateNewVPO3(sheet_Vendor, sheet_Work, starttime, endtime, zero_reset ) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName(sheet_Vendor);
  var sheet1 = ss.getSheetByName(sheet_Work);
  var sheet2 = ss.getSheetByName("Config");
  var Reset = sheet2.getRange("G2").getValue(); // Reset field
  //Logger.log(Reset);
  if (Reset == "1" ) {
    if (sheet0.getLastRow() == 1) {
      var lastRow = 1;
    } else {
      var lastRow = sheet0.getLastRow() - 1;
    }
    var range0 = sheet0.getRange(2, 1, lastRow, sheet0.getLastColumn());
    Logger.log('last'+sheet0.getLastColumn());
    range0.clearContent();
    
    sheet0 = ss.getSheetByName("Vendor_Ext");     
    if (sheet0.getLastRow() == 1) {
      var lastRow = 1;
    } else {
      var lastRow = sheet0.getLastRow() - 1;
    }
    var range0 = sheet0.getRange(2, 1, lastRow, sheet0.getLastColumn());
    
    range0.clearContent();
    
    
    if (zero_reset == 1) { sheet2.getRange("G2").setValue("0"); }
    starttime = 0;
    endtime = 24;
    Utilities.sleep(10000);
    var sheet0 = ss.getSheetByName(sheet_Vendor);
  }

    if (starttime == undefined) var starttime = 0;
    if (endtime == undefined) var endtime = 24;
    var nowtime = new Date();
    var nowhour = nowtime.getHours();
    if (nowhour <= endtime && nowhour >= starttime) {

        //var sheet0 = ss.getSheets()[0];
        //var sheet1 = ss.getSheets()[2];

        var lastRow0 = sheet0.getLastRow();
        var lastCol0 = sheet0.getLastColumn();
        var maxRow = sheet0.getMaxRows();

        if (maxRow < lastRow0 + 1) {
            sheet0.insertRowsAfter(lastRow0, 1);
            Logger.log('Blank Row inserted');

        }

        var lastRow1 = sheet1.getLastRow();
        var lastCol1 = sheet1.getLastColumn();
        //var lastRow1= 30;
        //var lastCol1 = 2;
        //Logger.log(lastRow0);
        //Logger.log(lastCol0); 
        //Logger.log(maxRow);
        //Logger.log(lastRow1);
        //Logger.log(lastCol1);
        var A2Value = sheet1.getRange("A2").getValue();
      if (lastRow1>1 && A2Value >0 ) {
            var range1 = sheet1.getRange(2, 1, lastRow1 - 1, lastCol1);
            range1.copyTo(sheet0.getRange(lastRow0+1,1), {contentsOnly:true});
            //range1.copyValuesToRange(sheet0, 1, lastCol1, lastRow0 + 1, lastRow0 + lastRow1 - 2);
            //sheet0.getRange(2, 1, sheet0.getLastRow() - 1, sheet0.getLastColumn()).sort([1,2]);
        if (sheet_Vendor == 'Vendor_Ext') {
          sheet0.getRange("M:M").setBackground("yellow");
        } else {
          sheet0.getRange("X:Y").setBackground("yellow");
        }
        
        }

    }
}

function Protect(distSheet, distRange, Type) {   // Lockdown Sheet , 1=ENABLE, 2=DISABLE
  
  if (distSheet == null) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(distSheet);
  
  switch (Type) {
    case 2:
      var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < protections.length; i++) {
        var protection = protections[i];
        if (protection.canEdit() && protection.getDescription()=='Import_Protecting') {
          protection.remove();
        }
      }
      break;      
      
    case 1:
    default:
      
      var protection = sheet.getRange(distRange).protect().setDescription('Import_Protecting');
      
      // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
      // permission comes from a group, the script throws an exception upon removing the group.
      var me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
      break;
  }
  
   
}

/**********************************************************************************   [GarZ][Code]  *************************************************************/

                                /** Paste to FC_RM_WH function start at[2021/4/12] **/

//定義執行欄位 SheetName:[Vendor_KeyIn],[FC_RM_WH]
//   function Vendor_Paste(){
//     let OriginSheet = 'Vendor_KeyIn';
//     let dest_col = 29;
//     let src_col = 25;
//     FillNote('Vendor',src_col,dest_col,OriginSheet)
//   }
//   function FC_RM_WH_Paste(){
//   let OriginSheet = 'FC_RM_KeyIn';
//   var destrange = 14;                                     //欄位:將被format的
//   var srcRange = 13;                                      //欄位:將被paste的
//   FillNote('FC_RM_WH', srcRange,destrange,OriginSheet);               //('sheetName',paste,format)
// }
// ///////////////////////////////////////////////////////////////////////////////////

// //取得note並插入至指定欄位
//   function FillNote(sheetname, src_col, dest_col,OriginSheet) {
//   SpreadsheetApp.getActiveSpreadsheet().toast(`KeyIn From ${OriginSheet} To ${sheetname},it will take a few second`,`Status`);
//   let sheet = ss.getSheetByName(sheetname);
//   let NoteContent = sheet.getRange(1,src_col).getNote();  //取得note
//   sheet.getRange(1, dest_col).setFormula(NoteContent);    //置入方程式                       
//   exchange(sheetname, src_col, dest_col,OriginSheet);                 //執行交換
//   }
// //取得值並交換
//   function exchange(sheetname, src_col, dest_col,OriginSheet){
//     let sheet =ss.getSheetByName(sheetname)
//     let rangeLast = sheet.getLastRow();                   //取得最後一行
//     let destrange = sheet.getRange(1,dest_col,rangeLast); //定義copy的range

//     let destCopy = destrange.getValues();                 //range內的值
//     sheet.getRange(1,src_col,rangeLast).setValues(destCopy);//定義要paste的range
//     sheet.getRange(1,dest_col).clearContent();              //清除原欄位
//     ss.getSheetByName(OriginSheet).getRange('B2:E').clear();

//      SpreadsheetApp.getActiveSpreadsheet().toast(`KeyIn From ${OriginSheet} To ${sheetname},Done`,`Status`);
//      var ui = SpreadsheetApp.getUi();
//      ui.alert('KeyIn Has Done Completely!')
     
//   }


////////////工具列////////////
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('頁面工具')
          .addItem('FC_RM_KeyIn -> FC_RM_WH ', 'FC_RM_WH_Paste')
          .addItem('Vendor_KeyIn -> Vendor ', 'Vendor_Paste')
      .addToUi();
}
                             /** paste function Paste function done at[2021/4/13] off **/
 

/*********************************************************************************************************************************************/ 
