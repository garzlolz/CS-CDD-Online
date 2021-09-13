
// //此為 FC_RM_KeyIn => FC_RM_WH 
// function FC_RM_WH_Paste(){
//      const Keyin_Sheet_Name = 'FC_RM_KeyIn'
//      const KeyIn_Document_Number='document #';     //FC_RM_KeyIn [document #] = 1
//      const KeyIn_Item='Item';                      //FC_RM_KeyIn [Item] = 2
//      const KeyIn_Quantity= 'Qty'                   //FC_RM_KeyIn [Qty] = 3
//      const KeyIn_Date = 'Date'                 //FC_RM_KeyIn [Date] = 4
//      const Paste_Sheet_Name = 'FC_RM_WH'
//      const Paste_Document_Number='Document Number';    //FC_RM_WH [Document Number] = 3
//      const Paste_item = 'Item';                        //FC_RM_WH [Item] = 5
//      const Paste_Quantity='Quantity'                   //FC_RM_WH [Quantity] = 6

//           GoExchange(Keyin_Sheet_Name,Paste_Sheet_Name,
//                      KeyIn_Document_Number,
//                      KeyIn_Item,KeyIn_Quantity,
//                      KeyIn_Date,Paste_Document_Number,
//                      Paste_item,Paste_Quantity);
// }
// //此為 Vendor_KeyIn => Vendor
// function Vendor_Paste(){
//      const Keyin_Sheet_Name = 'Vendor_KeyIn'
//      const KeyIn_Document_Number='document #';   
//      const KeyIn_Item='Item';                     
//      const KeyIn_Quantity= 'Qty'                  
//      const KeyIn_Date = 'Date'                    
//      const Paste_Sheet_Name = 'Vendor'
//      const Paste_Document_Number='Document Number';    //Vendor [Document Number] = 3
//      const Paste_item = 'WO Assembly Item';            //Vendor [WO Assembly Item] = 17
//      const Paste_Quantity='Quantity'                   //Vendor [Quantity] =14

//            GoExchange(Keyin_Sheet_Name,Paste_Sheet_Name,
//                       KeyIn_Document_Number,
//                       KeyIn_Item,KeyIn_Quantity,
//                       KeyIn_Date,Paste_Document_Number,
//                       Paste_item,Paste_Quantity);
// }
// //此為執行 貼上的函式
//   function GoExchange(Keyin_Sheet_Name,Paste_Sheet_Name,
//                       KeyIn_Document_Number,
//                       KeyIn_Item,KeyIn_Quantity,
//                       KeyIn_Date,Paste_Document_Number,
//                       Paste_item,Paste_Quantity)
//   { 
//     var  Keyin_Sheet_Title = ss.getSheetByName(Keyin_Sheet_Name).getDataRange().getValues()[0];    
//     //取  Key in sheet 的 Header 分別為Document_Number,Item,Quantity,Date 
//     var  Col_KeyIn_Document_Number=Keyin_Sheet_Title.indexOf(KeyIn_Document_Number);   
//     var  Col_KeyIn_Item=Keyin_Sheet_Title.indexOf(KeyIn_Item);                     
//     var  Col_KeyIn_Quantity= Keyin_Sheet_Title.indexOf(KeyIn_Quantity)                  
//     var  Col_KeyIn_Date = Keyin_Sheet_Title.indexOf(KeyIn_Date)

//     var  Paste_Sheet_Title = ss.getSheetByName(Paste_Sheet_Name).getDataRange().getValues()[0];   
//     //取  Paste in sheet 的 Header 分別為Document_Number,Item,Quantity
//     var  Col_Paste_Document_Number=Paste_Sheet_Title.indexOf(Paste_Document_Number);   
//     var  Col_Paste_item = Paste_Sheet_Title.indexOf(Paste_item);            
//     var  Col_Paste_Quantity=Paste_Sheet_Title.indexOf(Paste_Quantity)                   

//     //若表單 Paste_Sheet 名稱為 Vendor 先將H欄,I欄 定義 以及Note_Change不知為何是25,13 index出來是24,14
//     if(Paste_Sheet_Name == 'Vendor'){
//       var Col_Paste_Note_Change = Paste_Sheet_Title.indexOf('Vendor Note')                                             
//       var Col_Item_H=Paste_Sheet_Title.indexOf('Item');
//       var Col_Description_I=Paste_Sheet_Title.indexOf('Description');
//       var Col_keyIn_Ratio = Keyin_Sheet_Title.indexOf('RATIO 勿動');
//     }
//     else{
//        var Col_Paste_Note_Change = Keyin_Sheet_Title.indexOf('Note Change');
//     }

//     //取得表單,範圍的值
//     var KeyIn_Sheet = ss.getSheetByName(Keyin_Sheet_Name);
//     var KeyIn_DataRange = KeyIn_Sheet.getRange(1,1,KeyIn_Sheet.getLastRow(),KeyIn_Sheet.getLastColumn());
//     var KeyIn_SheetValues = KeyIn_DataRange.getValues();
//     var Paste_Sheet = ss.getSheetByName(Paste_Sheet_Name);
//     var Paste_DataRange = Paste_Sheet.getRange(1,1,Paste_Sheet.getLastRow(),Paste_Sheet.getLastColumn())
//     var Paste_SheetValues = Paste_DataRange.getValues(); 


//     //出兩個錯誤警告的 Header
//     var errors=[`✗表單 「${Keyin_Sheet_Name}」有誤:\n`];
//     var P_errors=[`⚠警告:\n`]

//     //Main code 處理 判斷貼上
//     for(let i=0;i<Paste_SheetValues.length;i++){
//       let string = '';
//       for(let j=0;j<KeyIn_SheetValues.length;j++){
//         /******* 若表單為Vendor 先判斷item的位置 *******/

//         //如果 表單Paste_Sheet_Name名稱是Vendor 及Vendor的 Document_Number與Keyin相等，則
//         if( Paste_Sheet_Name == 'Vendor' && 
//             Paste_SheetValues[i][Col_Paste_Document_Number] == KeyIn_SheetValues[j][Col_KeyIn_Document_Number])
//           {
//             //若相等 先將數量做處理 把Vendor Keyin 的數量 變成 Vendor的的樹樣乘上 Ratio
//             KeyIn_SheetValues[j][Col_KeyIn_Quantity] = Paste_SheetValues[i][Col_Paste_Quantity]*KeyIn_SheetValues[j][Col_keyIn_Ratio]
//             //若Vedor中的 WO Assembly Item 為空值的話，則
//             if(Paste_SheetValues[i][Col_Paste_item]== ''){
//               //若 欄位H 的item  與keyin的item為相同
//               if(Paste_SheetValues[i][Col_Item_H] == KeyIn_SheetValues[j][Col_KeyIn_Item]){
//                 Paste_SheetValues[i][Col_Paste_item] = Paste_SheetValues[i][Col_Item_H];
//               }
//               //但若 欄位H 與keyin item不相同 最後將item的位置指定給 欄位I
//               else if(Paste_SheetValues[i][Col_Item_H] != KeyIn_SheetValues[j][Col_KeyIn_Item]){
//                 Paste_SheetValues[i][Col_Paste_item] = Paste_SheetValues[i][Col_Description_I];
//               }
//             }
//           }
//         //若document,item 與 Keyin相等,且數量符合 則set到Note Change 若不是 就將該列push到錯誤的陣列 
//         if(((Paste_SheetValues[i][Col_Paste_Document_Number]==KeyIn_SheetValues[j][Col_KeyIn_Document_Number]) && 
//             (Paste_SheetValues[i][Col_Paste_item]==KeyIn_SheetValues[j][Col_KeyIn_Item]) &&
//             (Paste_SheetValues[i][Col_Paste_Quantity]>=KeyIn_SheetValues[j][Col_KeyIn_Quantity])&&
//             (KeyIn_SheetValues[j][Col_KeyIn_Document_Number]!='')))
//           {
//             string = 
//                 `SHIP@${formatDate(KeyIn_SheetValues[j][Col_KeyIn_Date])}@${KeyIn_SheetValues[j][Col_KeyIn_Quantity]}`
//             let range =Paste_Sheet.getRange(i+1,Col_Paste_Note_Change+1);
//               if(range.getValue()==''){
//                 range.setValue(string)
//                 KeyIn_Sheet.getRange(j+1,1,1,KeyIn_Sheet.getLastColumn()).clearContent(); 
//               }
//               else{
//                 P_errors.push(`第${j+1}列 重複輸入`)
//               }
//           }
//       }
//     }
      
//       //檢查沒有貼上的資料,重新取得Key in sheet 中的值
//       var CheckValue = ss.getSheetByName(Keyin_Sheet_Name).getDataRange().getValues();
//       var documents = [];
//       for(let i=1;i<Paste_SheetValues.length;i++){
//         documents.push(`${Paste_SheetValues[i][Col_Paste_Document_Number]}${Paste_SheetValues[i][Col_Paste_item]}`)
//       };

//       for(let i=1;i<CheckValue.length;i++){
//         let check = CheckValue[i][0];
//         if(check !=''){
//           if(documents.indexOf(check)<0){
//             errors.push(`第 ${i+1}列\n 欄位[documents#]或[item] 有誤或不存在.\n`);
//           }
//           else{
//             for(j=0;j<Paste_SheetValues.length;j++){
//               if((CheckValue[i][0] == Paste_SheetValues[j][Col_Paste_Document_Number]+Paste_SheetValues[j][Col_Paste_item])&&
//                   CheckValue[i][Col_KeyIn_Quantity]>Paste_SheetValues[j][Col_Paste_Quantity]){
//                       errors.push((`第 ${i+1}列\n 欄位[Qty] 數量有誤\n`));
//               }
//             }
//           }
//         }
//       };

//         if(errors.length>1 && P_errors.length>1){
//           SpreadsheetApp.getUi().alert(errors.join('\n')+'\n'+P_errors.join('\n'));
//         }
//         else if(errors.length>1 && P_errors.length==1){
//           SpreadsheetApp.getUi().alert(errors.join('\n'));
//         }
//         else if(P_errors.length>1 && errors.length==1){
//           SpreadsheetApp.getUi().alert(P_errors.join('\n'));
//         }
//         else{
//           SpreadsheetApp.getActiveSpreadsheet()
//           .toast(`從 ${Keyin_Sheet_Name} 輸出到 ${Paste_Sheet_Name},完成`,`Status`);
//         } 
//   }


// function formatDate(date) {
//       var d = new Date(date),
//           month = '' + (d.getMonth() + 1),
//           day = '' + d.getDate();
//           //day = '' + d.getDate(),
//           // year = d.getFullYear();
//       if (month.length < 2) 
//           month = '0' + month;
//       if (day.length < 2) 
//           day = '0' + day;
//       return [month, day].join('/');
//     }
// ////////////工具列////////////
// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('頁面工具')
//           .addItem('FC_RM_KeyIn -> FC_RM_WH ', 'FC_RM_WH_Paste')
//           .addItem('Vendor_KeyIn -> Vendor ', 'Vendor_Paste')
//       .addToUi();
// }
//                              /** paste function Paste function done at[2021/9/2] off **/
 

// /*********************************************************************************************************************************************/