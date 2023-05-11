var XLSX = require("xlsx");
var workbook = XLSX.readFile("tsc.xlsx");
const fs = require("fs")
var sheet_name_list = workbook.SheetNames
// console.log(sheet_name_list); // getting as Sheet1

sheet_name_list.forEach(function (y) {
    var worksheet = workbook.Sheets[y];
    //getting the complete sheet
    // console.log(worksheet);
  
    var headers = {};
    var data = [];
    var data1 = [];

    for (z in worksheet) {
      if (z[0] === "!") continue;
      //parse out the column, row, and value
      var col = z.substring(0, 1);
      console.log("this is col",col);

      console.log("this is z",z);
  
      var row = parseInt(z.substring(1));
      console.log("row",row);
      // if(){

      // }
  
      var value = worksheet[z].v;
      // console.log(value);
      console.log("this is  value",value)
      //store header names
      if (row == 1) {
        headers[col] = value;
        // storing the header names
        continue;
      }
      console.log("ta inja")
      if( row==35){
        // console.log("rund")
        // console.log(value)
        // data1[row] = {};
        // data1[row][headers[col]] = value;
        // console.log("data1",data1);
        
        // console.log("data1[row]",data1[row][headers[col]])
        // data[row] = {};
        // const datamain=data1[row][headers[col]]=value;
        // console.log("value",data1[row][headers[col]]=value)
        // console.log("json",JSON.stringify(datamain))
       
        data1[row][headers[col]] = value;
        data1.shift(); 
        data1.shift();
        fs.writeFile("./data1.json" , JSON.stringify(data1), (error) => {
          if (error) {
            console.log('An error has occurred ', error);
            return;
          }
          console.log('Data written successfully to disk');
          return
        });
      }
      if (!data[row]) data[row] = {};
      data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift(); 
    data.shift();
    console.log();

    fs.writeFile("./data.json" , JSON.stringify(data), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
  });