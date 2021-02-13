// Отладка команда debugger
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = 
      [
        {name : "Сохранить в истории",functionName : "save_history"},
        {name : "Добавить водителей",functionName : "add_list_drivers"},
        {name : "Расчитать",functionName : "count"},
        {name : "Обновить диаграмму", functionName : "updateData"}
      ]
  sheet.addMenu("Скрипты", entries);
}


function save_history(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_data = ss.getSheetByName('Данные');
  var sheet_history = ss.getSheetByName('История');

  var rows_data = sheet_data.getLastRow(); // кол-во строк
  var cols_data = sheet_data.getLastColumn(); // кол-во ячеек в строке
  
  var rows_history = sheet_history.getLastRow(); // кол-во строк
  var cols_history = sheet_history.getLastColumn(); // кол-во ячеек в строке
    
  var range = sheet_data.getRange(2,1, rows_data-1, cols_data );
  
  range.copyValuesToRange(sheet_history,1,cols_data,rows_history+1,rows_history+rows_data);
  
  range.clear()
  sheet_history.sort(1,false)

}

function add_list_drivers() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet_data = ss.getSheetByName('Данные');
  var sheet_list_drivers = ss.getSheetByName('Водители');
  var rows_list_drivers = sheet_list_drivers.getLastRow(); // кол-во строк

  var range_date = sheet_data.getRange(2,1, rows_list_drivers-1,1 );  
  range_date.setFormula('=TODAY()')

  var range = sheet_list_drivers.getRange(2,1, rows_list_drivers-1,3 );
  
  range.copyValuesToRange(sheet_data,5,7,2,rows_list_drivers+1);
   
}

function count(){

  var arr = [];
  var str_condition = "";
  var str = "";
  var int_travel_profit;
  var int_bonus;
  var int_nal;
  var int_beznal;
  var my_profit;
  var col_expense;
  var branding;
  var value;
  var sym;
  var min_sum = 8000;
  var max_sum = 10000;
  var col_my_dohod;
  var col_send_to_driver;
  var col_partners;
  var col_prostoy_auto;
  var first_value;
  var first_sym;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  debugger
  var sheet_data = ss.getSheetByName('Данные');
  var rows_sheet_data = sheet_data.getLastRow(); // кол-во строк
  var send_to_driver_column_range = sheet_data.getRange(2, 12, rows_sheet_data-1)
  var my_profit_column_range = sheet_data.getRange(2, 14, rows_sheet_data-1)
  
  
  for(var i = 2; i < rows_sheet_data+1; i++ ) {
    
    col_prostoy_auto = sheet_data.getRange(i, 3 ).getValues();// столбец: Простой автомобиля
    col_expense = sheet_data.getRange(i, 4 ).getValues();// столбец: Расход
    col_my_dohod = sheet_data.getRange(i, 14 ); // столбец: Мой доход 
    col_send_to_driver = sheet_data.getRange(i, 12 ); //столбец: Отправить водителю
    col_partners = sheet_data.getRange(i, 15 ); // столбец: Партнеры
    
    branding = Number(sheet_data.getRange(i, 16 ).getValues());// столбец: Брендирование
    arr = sheet_data.getRange(i, 7, 1, 5).getValues(); //Получить массив данных: Условие работы, Доход от поездок,Бонус,Наличный,Безналичный
    
    str_condition = arr[0][0];
    int_travel_profit = arr[0][1];
    int_bonus = arr[0][2];
    int_nal = arr[0][3];
    int_beznal = arr [0][4];
    
    if (int_travel_profit){
      
      str = str_condition.split(" ")
      
      if(str.length == 2){
      
        sym = str[1];
        value = Number(str[0]);
      
        switch(sym) {
          case '%':
            my_profit = int_travel_profit/100*value;
            
            // Условие для столбца Условие работы (50 %, 5050 %)            
            if(value == 50 || value == 5050){
              
              if(value == 5050){
                col_my_dohod.setValue(int_travel_profit / 2 -col_expense - col_prostoy_auto + + branding);// столбец: Мой доход 
                col_send_to_driver.setValue(int_beznal-(int_travel_profit / 2)) //столбец: Отправить водителю
                break;
              } //if(value == 5050)
              
              if(int_travel_profit < min_sum){
                col_my_dohod.setValue(min_sum / 2 -col_expense - col_prostoy_auto + branding);// столбец: Мой доход 
                col_send_to_driver.setValue(int_beznal-(min_sum / 2)) //столбец: Отправить водителю
                break;
              }
              if(int_travel_profit > max_sum){
                col_my_dohod.setValue(max_sum / 2 - col_expense - col_prostoy_auto + branding);// столбец: Мой доход 
                col_send_to_driver.setValue(int_beznal-(max_sum / 2)) //столбец: Отправить водителю
                break;
              }
              else{
                col_my_dohod.setValue(my_profit-col_expense - col_prostoy_auto + branding);// столбец: Мой доход 
                col_send_to_driver.setValue(int_beznal-my_profit) //столбец: Отправить водителю
                break;
              }
            } //(value == 50 || value == 5050)
            else{
              col_my_dohod.setValue(my_profit-col_expense - col_prostoy_auto + branding);// столбец: Мой доход 
              col_send_to_driver.setValue(int_beznal-my_profit) //столбец: Отправить водителю
              break;
            }
          case 'UAH':
            col_my_dohod.setValue(value-col_expense - col_prostoy_auto + branding); // столбец: Мой доход
            col_send_to_driver.setValue(int_beznal-value) //столбец: Отправить водителю
            break;
          
          case 'BRAND':
            if(value != 0) {
              if(value < int_beznal){
                col_my_dohod.setValue(branding); // столбец: Мой доход
                col_partners.setValue(value); // столбец: Партнеры
                col_send_to_driver.setValue(int_beznal-value); //столбец: Отправить водителю
                break;
              } else{
                col_my_dohod.setValue(branding); // столбец: Мой доход
                col_partners.setValue(int_beznal); // столбец: Партнеры
                col_send_to_driver.setValue(int_beznal-value); //столбец: Отправить водителю
                break;
              }  
            } else{
              col_my_dohod.setValue(branding); // столбец: Мой доход
              col_partners.setValue(int_beznal); // столбец: Партнеры
              break;
            }
        } //switch
      } //if(str.length == 2) 
    
      // Условие для столбца Условие работы (например 5 % 400 UAH, 5% 3300 BRAND, 5 % 50 %, 400 UAH 3300 BRAND)
      if(str.length == 4){
      
        first_value = Number(str[0]);
        first_sym = str[1]      
        value = Number(str[2]);
        sym = str[3]
      
        switch(first_sym){
          case '%':
            my_profit = int_travel_profit/100*Number(str[0]);
            break;
          case 'UAH':
            my_profit = first_value;
            break;
        } 
            
        switch(sym){
            
          case 'UAH':    
            if(my_profit >= value){
              col_my_dohod.setValue(value-col_expense - col_prostoy_auto +branding); // столбец: Мой доход
              col_send_to_driver.setValue(int_beznal-value); //столбец: Отправить водителю
              break;  
            } else{
              col_my_dohod.setValue(my_profit-col_expense - col_prostoy_auto +branding); // столбец: Мой доход
              col_send_to_driver.setValue(int_beznal-my_profit); //столбец: Отправить водителю
              break;
            }
          case 'BRAND':
            col_my_dohod.setValue(my_profit+branding); // столбец: Мой доход
            
            if(value != 0){
              if(value < int_beznal){
                if (int_beznal-(value + my_profit) >= 0) {
                  col_partners.setValue(value); // столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(value + my_profit)); //столбец: Отправить водителю
                  break;
                 }
                 else {
                  col_partners.setValue(int_beznal - my_profit); // столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(value + my_profit)); //столбец: Отправить водителю
                  break;                  
                 }
              }
              if(value == int_beznal){
                col_partners.setValue(value-my_profit); // столбец: Партнеры
                col_send_to_driver.setValue(0-my_profit); //столбец: Отправить водителю
                break;
              }
              if(value > int_beznal){
                if(int_beznal <= my_profit){
                  col_partners.setValue(0); // столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(value + my_profit)); //столбец: Отправить водителю
                  break;                  
                 } else {
                  col_partners.setValue(int_beznal-my_profit); // столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(value + my_profit)); //столбец: Отправить водителю
                  break;
                }
              }  
            } else {
              col_partners.setValue(int_beznal); // столбец: Партнеры
              break;
            }
            
          case '%':
            
            if (my_profit < int_beznal) {
            
              if(int_travel_profit < min_sum){
                
                if (int_beznal - my_profit < min_sum / 2 ) {
                  col_my_dohod.setValue(my_profit);// столбец: Мой доход
                  col_partners.setValue(int_beznal - my_profit);// столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(min_sum / 2)-my_profit) //столбец: Отправить водителю
                }  
                else {
                  col_my_dohod.setValue(my_profit);// столбец: Мой доход
                  col_partners.setValue(min_sum / 2);// столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(min_sum / 2)-my_profit) //столбец: Отправить водителю
                }                               
                break;
              }
               else if(int_travel_profit > max_sum){
                
                 if (int_beznal - my_profit < max_sum / 2 ) {
                   col_my_dohod.setValue(my_profit);// столбец: Мой доход
                   col_partners.setValue(int_beznal - my_profit);// столбец: Партнеры
                   col_send_to_driver.setValue(int_beznal-(max_sum / 2)-my_profit) //столбец: Отправить водителю
                  }  
                  else {
                   col_my_dohod.setValue(my_profit);// столбец: Мой доход
                   col_partners.setValue(max_sum / 2);// столбец: Партнеры
                   col_send_to_driver.setValue(int_beznal-(max_sum / 2)-my_profit) //столбец: Отправить водителю
                  }                               
                  break;
              }
             else{
             
                if (int_beznal - my_profit < int_travel_profit / 100 * value) {
                  col_my_dohod.setValue(my_profit);// столбец: Мой доход
                  col_partners.setValue(int_beznal - my_profit);// столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(int_travel_profit / 100 * value)-my_profit) //столбец: Отправить водителю
                  
                }  
                else {
                  col_my_dohod.setValue(my_profit);// столбец: Мой доход
                  col_partners.setValue( int_travel_profit / 100 * value);// столбец: Партнеры
                  col_send_to_driver.setValue(int_beznal-(int_travel_profit / 100 * value)-my_profit) //столбец: Отправить водителю
                }                               
                break;
              }
              }
              else {
                col_my_dohod.setValue(my_profit);// столбец: Мой доход
                col_send_to_driver.setValue(int_beznal-my_profit) //столбец: Отправить водителю            
              }
          
            
        }// switch   
    
      } //if(str.length == 4)
      
   } //if (int_travel_profit)
    
 } //for
 
 send_to_driver_column_range.setBorder(false, true, false, true, true, false, "red", SpreadsheetApp.BorderStyle.SOLID);
 my_profit_column_range.setBorder(false, true, false, true, true, false, "blue", SpreadsheetApp.BorderStyle.SOLID); 

} //func count

  
function updateData() {
  
var list = [];  
var sheetName = "Сводная таблица"; 
var ss = SpreadsheetApp.getActiveSpreadsheet();
var a = 1

var sheet = ss.getSheetByName(sheetName);
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();

var copySheet = ss.getSheetByName('CopyPivotTable');

if (!copySheet.isSheetHidden()) { 
   copySheet.hideSheet();
}
  
copySheet.clear()


var range = sheet.getRange(4, 1, lastRow-3, lastColumn);
var dataValues = range.getValues();
  
var dataValuesLength = dataValues.length

for (var i = 0;  i < dataValuesLength-1; i++){
  
    copySheet.getRange(a, 1).setValue(dataValues[i][0]);
    copySheet.getRange(a, 2).setValue(dataValues[i][lastColumn-1]);
    a++
  }

var copySheetRange = copySheet.getDataRange();
var copySheetRangeSort = copySheetRange.sort({column: 2, ascending: false}); 

}