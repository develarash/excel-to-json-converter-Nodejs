var XLSX = require("xlsx");
var workbook = XLSX.readFile("maintsc.xlsx");
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
      // console.log("this is col",col);

      // console.log("this is z",z);
  
      var row = parseInt(z.substring(1));
      // console.log("row",row);
      // if(){

      // }
  
      var value = worksheet[z].v;
      // console.log(value);
      // console.log("this is  value",value)
      //store header names
      if (row == 1) {
        headers[col] = value;
        // storing the header names
        continue;
      }
      // console.log("ta inja")
         
      if (!data[row]) data[row] = {};
      data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift(); 
    data.shift();
    
    const datamain=data.shift.length;
    const operation =data.length/6;
    const half1=data.slice(0,146902);
    const half2=data.slice(146903,293806)



    const sec1 = data.slice(0,48966);
    const sec2 = data.slice(48966,97932);
    const sec3 = data.slice(97932,146898);
    const sec4 = data.slice(146898,195864);
    const sec5 = data.slice(195864,244830);
    const sec6 = data.slice(244830,293804);



    console.log("this is type ",typeof data);
    console.log("this is data ",data.shift.length);
    console.log("ba shift ",data.shift.length);
    console.log("bedon shift  ",data.length);

    console.log("this is shift / 2 ",operation);

    console.log("this is half1 ",half1.length);
    console.log("this is half2 ",half2.length);

    // console.log("this is half ",half.length);

    console.log("this is type ",typeof data);

    // fs.writeFile("./jhalf1.json" , JSON.stringify(half1), (error) => {
    //     if (error) {
    //       console.log('An error has occurred ', error);
    //       return;
    //     }
    //     console.log('Data written successfully to disk');
    //   });


    //   fs.writeFile("./jhalf2.json" , JSON.stringify(half2), (error) => {
    //     if (error) {
    //       console.log('An error has occurred ', error);
    //       return;
    //     }
    //     console.log('Data written successfully to disk');
    //   });
      fs.writeFile("./out/part1.json" , JSON.stringify(sec1), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      fs.writeFile("./out/part2.json" , JSON.stringify(sec2), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      fs.writeFile("./out/part3.json" , JSON.stringify(sec3), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      fs.writeFile("./out/part4.json" , JSON.stringify(sec4), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      fs.writeFile("./out/part5.json" , JSON.stringify(sec5), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      fs.writeFile("./out/part6.json" , JSON.stringify(sec6), (error) => {
        if (error) {
          console.log('An error has occurred ', error);
          return;
        }
        console.log('Data written successfully to disk');
      });
      
  });