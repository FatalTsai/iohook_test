'use strict';

var correspond =[];
var keycode = require('keycode'); //npm i keycode
const ioHook = require('iohook');//npm i iohook
const xlsx=require('xlsx'); //npm i xlsx 

var barcode="";

var workbook = xlsx.readFile('BYOC.csv'); //write BYOC.csv(filename) 
var sheetNames=workbook.SheetNames;
//console.log(sheetNames)
var worksheet = workbook.Sheets[sheetNames[0]] //write this worksheet
let counter=2;

//console.log("[!ref] = "+worksheet['!ref'])

ioHook.on('keydown', event => {

    var element = keycode.names[  event['rawcode']  ];//save the char
    
    //console.log(keycode.names[  event['rawcode']  ]  ); 
    //console.log(event);
    if(element!='enter') //if you don't not PRESS enter, then append in the var barcode
    {
        barcode+=element;
    }
    else{//once you press enter, save in A column

        while(worksheet['A'+counter]!=undefined)//write the very first null A column 
            counter++


        console.log(barcode);
        worksheet['A'+counter++]={v:barcode,t:'s'} //append the barcode in workbook
      
        //worksheet['!ref']='A1:E'+counter
        //console.log(worksheet)

        let wb = xlsx.utils.book_new() //save the workbook

        xlsx.utils.book_append_sheet(wb, worksheet, 'BYOC')
        xlsx.writeFile(workbook, 'BYOC.csv')


        barcode =''
    }

   

});

// Register and start hook
ioHook.start();

// Alternatively, pass true to start in DEBUG mode.
ioHook.start(true);

