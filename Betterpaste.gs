/*
  此處分為兩個function 分別為FC_RM_WH,Vendor 
  在兩個function 裏面先定義各個欄位的名稱
  再將參數帶入GO function 執行
*/
function FC_RM_WH(){
     const Keyin_Sheet_Name = 'FC_RM_KeyIn'
     const KeyIn_Document_Number='document #';     //FC_RM_KeyIn [document #] = 1
     const KeyIn_Item='Item';                      //FC_RM_KeyIn [Item] = 2
     const KeyIn_Quantity= 'Qty'                   //FC_RM_KeyIn [Qty] = 3
     const KeyIn_Date = 'Date'                 //FC_RM_KeyIn [Date] = 4
     const Paste_Sheet_Name = 'FC_RM_WH'
     const Paste_Document_Number='Document Number';    //FC_RM_WH [Document Number] = 3
     const Paste_item = 'Item';                        //FC_RM_WH [Item] = 5
     const Paste_Quantity='Quantity'                   //FC_RM_WH [Quantity] = 6
     const Paste_NoteChange = 'Note Change';

     GoDeal(Keyin_Sheet_Name,Paste_Sheet_Name,
            KeyIn_Document_Number,
            KeyIn_Item,KeyIn_Quantity,
            KeyIn_Date,Paste_Document_Number,
            Paste_item,Paste_Quantity,Paste_NoteChange);
}
function Vendor(){
     const Keyin_Sheet_Name = 'Vendor_KeyIn'
     const KeyIn_Document_Number='document #';   
     const KeyIn_Item='Item';                     
     const KeyIn_Quantity= 'Qty'                  
     const KeyIn_Date = 'Date'                    
     const Paste_Sheet_Name = 'Vendor'
     const Paste_Document_Number='Document Number';    //Vendor [Document Number] = 3
     const Paste_item = 'WO Assembly Item';            //Vendor [WO Assembly Item] = 17
     const Paste_Quantity='Quantity'                   //Vendor [Quantity] =14
     const Paste_NoteChange = 'Vendor Note'

     GoDeal(Keyin_Sheet_Name,Paste_Sheet_Name,
            KeyIn_Document_Number,
            KeyIn_Item,KeyIn_Quantity,
            KeyIn_Date,Paste_Document_Number,
            Paste_item,Paste_Quantity,Paste_NoteChange);
}

/*
  GoDeal是主要處理 執行貼上,清除,顯示錯誤訊息 的主要程式
  在一開始將所有的欄位名稱的位置index出來
  接下來跑兩個迴圈判斷單號,item,數量 吻合就貼上進行
  清除key in 欄位 最後將未貼上的資料推入至錯誤訊息
 */
function GoDeal(sheetName_K,sheetName_P,
                documentNumber_K,
                item_K,quantity_K,
                date_K,documentNumber_P,
                item_P,quantity_P,note_P)
  {
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        var Sheet_K = ss.getSheetByName(sheetName_K);
        var Sheet_P = ss.getSheetByName(sheetName_P);

        var SheetValues_K = Sheet_K.getDataRange().getValues();
        var SheetValues_P = Sheet_P.getDataRange().getValues();

        var Doc_K = SheetValues_K[0].indexOf(documentNumber_K);
        var Doc_P = SheetValues_P[0].indexOf(documentNumber_P);

        var Item_K = SheetValues_K[0].indexOf(item_K);
        var Item_P = SheetValues_P[0].indexOf(item_P);

        var Qty_K  = SheetValues_K[0].indexOf(quantity_K);
        var Qty_P = SheetValues_P[0].indexOf(quantity_P);

        var Date = SheetValues_K[0].indexOf(date_K);
        var NoteChange = SheetValues_P[0].indexOf(note_P);

        var ERRORS_MISS=[`✗錯誤 「${sheetName_K}」:\n`]; 
        var ERRORS_repeat=[`⚠警告:\n`];

        if(sheetName_P == 'Vendor')
        {
          var Ratio_K = SheetValues_K[0].indexOf('RATIO 勿動')
          var Description_Col_I = SheetValues_P[0].indexOf('Description');
          var Item_Col_H = SheetValues_P[0].indexOf('Item');
        }

        /*
         兩個 for迴圈分別跑 keyin表單以及paste表單
        首先最一開始判斷Keyin的Doc是否為零以及是否與paste的Doc相等
        若相等 先判斷是 Vendor還是FC 
        接下來判斷Item是否吻合以及是否是空值
        再來判斷Ratio或Quantity是否超過 
          如果沒有 則判斷是否以及輸入過
        若超過 則判斷該為超量輸入
        */

        for(let k=0;k<SheetValues_K.length;k++)
        {
          for(let p=0;p<SheetValues_P.length;p++)
          {
            var string = '';
                
           if(SheetValues_K[k][Doc_K]!=''&&  
              SheetValues_K[k][Doc_K] == SheetValues_P[p][Doc_P])
              {
                if(sheetName_P == 'Vendor')
                {
                  if((SheetValues_K[k][Item_K] == SheetValues_P[p][Item_P])||
                     (SheetValues_P[p][Item_P] == ''&&
                     (SheetValues_K[k][Item_K] == SheetValues_P[p][Item_Col_H]||
                      SheetValues_K[k][Item_K]== SheetValues_P[p][Description_Col_I])))
                     {
                        if(SheetValues_K[k][Ratio_K]<=1)
                        {
                          if(SheetValues_P[p][NoteChange]=='')
                          {
                            string = 
                           `SHIP@${formatDate(SheetValues_K[k][Date])}@${SheetValues_P[p][Qty_P]*SheetValues_K[k][Ratio_K]}`

                            Sheet_P.getRange(p+1,NoteChange+1).setValue(string);
                            Sheet_K.getRange(k+1,1,1,Sheet_K.getLastColumn()).clearContent();
                          }
                          else if(SheetValues_P[p][NoteChange]!=''&& ERRORS_repeat.indexOf(`第${k+1}列 重複輸入`)<0)
                          {  
                            ERRORS_repeat.push(`第${k+1}列 重複輸入`);
                          }  
                        }
                        else if(SheetValues_K[k][Ratio_K]>1 && ERRORS_MISS.indexOf(`第 ${k+1}列\n 欄位[Qty] 數量有誤\n`)<0)
                        {
                          ERRORS_MISS.push(`第 ${k+1}列\n 欄位[Qty] 數量有誤\n`);
                        }
                     }
                     
                }   

                else if(sheetName_P == 'FC_RM_WH' && SheetValues_K[k][Item_K]==SheetValues_P[p][Item_P])
                {
                
                      if(SheetValues_K[k][Qty_K]<=SheetValues_P[p][Qty_P])
                      {
                        if(SheetValues_P[p][NoteChange]=='')
                          {
                              string = 
                             `SHIP@${formatDate(SheetValues_K[k][Date])}@${SheetValues_K[k][Qty_K]}`

                              Sheet_P.getRange(p+1,NoteChange+1).setValue(string);
                              Sheet_K.getRange(k+1,1,1,Sheet_K.getLastColumn()).clearContent();  
                          }
                          else if(SheetValues_P[p][NoteChange]!='' && ERRORS_repeat.indexOf(`第${k+1}列 重複輸入`)<0)
                          {
                               ERRORS_repeat.push(`第${k+1}列 重複輸入`);
                          }
                      }
                      else if(SheetValues_K[k][Qty_K]>SheetValues_P[p][Qty_P] && ERRORS_MISS.indexOf(`第 ${k+1}列\n 欄位[Qty] 數量有誤\n`)<0)
                      {
                          ERRORS_MISS.push(`第 ${k+1}列\n 欄位[Qty] 數量有誤\n`);
                      }
                }
              }

            //判斷是否有空值
            if(SheetValues_K[k][Doc_K]!='' && SheetValues_K[k][Item_K] == ''&& 
               ERRORS_MISS.indexOf(`第 ${k+1}列\n 欄位 ${item_K} 為空值\n`)<0)

                ERRORS_MISS.push(`第 ${k+1}列\n 欄位 ${item_K} 為空值\n`);

            if(SheetValues_K[k][Doc_K]=='' && SheetValues_K[k][Item_K] != ''&& 
               ERRORS_MISS.indexOf(`第 ${k+1}列\n 欄位 ${documentNumber_K} 為空值\n`)<0)

                ERRORS_MISS.push(`第 ${k+1}列\n 欄位 ${documentNumber_K} 為空值\n`);
          }
        }
        
       
       
  
    /*
    在此處重新獲取一次keyin的資料 並對paste 
    進行每筆的index 若沒有一筆符合 將推送至
    錯誤清單，並顯示在表單上。
    */
      var P_CheckValue = ss.getSheetByName(sheetName_P).getDataRange().getValues();
      var K_CheckValue = ss.getSheetByName(sheetName_K).getDataRange().getValues();

      for(let k=1;k<K_CheckValue.length;k++)
      {
        var times = 0;
        if(K_CheckValue[k][Doc_K]!=''&& K_CheckValue[k][Item_K]!='')
        {
          for(let p=0;p<P_CheckValue.length;p++)
          {
            if((P_CheckValue[p].indexOf(K_CheckValue[k][Item_K])>0)&&
               (P_CheckValue[p].indexOf(K_CheckValue[k][Doc_K])>0))
            {
              times++;
              console.log(times)
            }
          }

          if(times == 0)
          {
              ERRORS_MISS.push(`第 ${k+1}列\n 欄位[Document]或[Item] 有誤\n`);
          }
        }
      }


       if(ERRORS_MISS.length>1 && ERRORS_repeat.length>1)
        {
          SpreadsheetApp.getUi().alert(ERRORS_MISS.join('\n')+'\n'+ERRORS_repeat.join('\n'));
        }
       else if(ERRORS_MISS.length>1 && ERRORS_repeat.length==1)
        {
          SpreadsheetApp.getUi().alert(ERRORS_MISS.join('\n'));
        }
       else if(ERRORS_repeat.length>1 && ERRORS_MISS.length==1)
        {
          SpreadsheetApp.getUi().alert(ERRORS_repeat.join('\n'));
        }
       else
        {
          SpreadsheetApp.getActiveSpreadsheet()
          .toast(`從 ${sheetName_K} 輸出到 ${sheetName_P},完成`,`Status`);
        } 
  }



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('頁面工具')
          .addItem('FC_RM_KeyIn -> FC_RM_WH ', 'FC_RM_WH')
          .addItem('Vendor_KeyIn -> Vendor ', 'Vendor')
      .addToUi();
}


function formatDate(date) {
      var d = new Date(date),
          month = '' + (d.getMonth() + 1),
          day = '' + d.getDate();
          //day = '' + d.getDate(),
          // year = d.getFullYear();
      if (month.length < 2) 
          month = '0' + month;
      if (day.length < 2) 
          day = '0' + day;
      return [month, day].join('/');
    }