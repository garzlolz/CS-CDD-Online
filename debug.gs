
  /****************************** 取得保護範圍並重設 ******************************************/
    function RemoveAndSet(){
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Sheet33');                                     //選擇Sheet33
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE); //取得以保護的範圍 
                                               
      var backup=[];
      var users = []
      for (var i = 0; i < protections.length; i++) {
          backup.push(protections[i]);
          users.push(protections[i].getEditors());
          Logger.log(users);
          protections[i].remove();
        }

      for(var i=0;i<backup.length;i++){
       var Range = backup[i].getRange();
       var protection = Range.protect().setDescription(backup[i].getDescription())
       protection.removeEditors(protection.getEditors())
       protection.addEditors(users[i])
      }
    }
   /******************************************************************************************/   
  /****************************** 取得保護範圍並重設 ******************************************/
    function NOTRemoveAndSet(){
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Sheet33');                                     //選擇Sheet33
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE); //取得以保護的範圍 
                                               
      var backup=[];
      var users = []
      for (var i = 0; i < protections.length; i++) {
          backup.push(protections[i]);
          users.push(protections[i].getEditors());
          protections[i].remove();
        }

      for(var i=0;i<backup.length;i++){
       var Range = backup[i].getRange();
       var protection = Range.protect().setDescription(backup[i].getDescription());
       Logger.log(protection.getEditors())
       protection.removeEditors(users[i])
       //Logger.log(users[i])
      }
    }
   /******************************************************************************************/   





