let xlsx = require('xlsx');
let wb = xlsx.readFile('data.xlsx');
let ws = wb.Sheets['Sheet1'];
var rowCount =0;
var colCount =0;
var startAscii = 65;
var tempString = 'A1';
var theString = "A1";
var iteratorString = "A1";
var tempDate,tempTime;
while(ws[tempString]!=null){
    ++rowCount;
    tempString = 'A' + rowCount.toString();
}

tempString = 'A1';
while(ws[tempString]!=null){
    ++colCount;++startAscii;
    tempString = String.fromCharCode(startAscii) + "1";

}
console.log(colCount);
rowCount = rowCount-1;
for(var i=1;i<=rowCount;++i){
 var temp1,temp2,temp1Type,temp2Type;
 for(var j=1;j<=colCount;++j){
     if(i==1)
     {
         if(j==1)
         continue;
         else if(j==2){
             temp1 = ws[formString(i,j)].v;
             temp1Type = ws[formString(i,j)].t;
             ws[formString(i,j)].t = "s";
             ws[formString(i,j)].v = "Time";
             
         }
         
        if(ws[formString(i,j+1)]!=null){
            temp2Type = ws[formString(i,j+1)].t
            temp2 = ws[formString(i,j+1)].v; 
            ws[formString(i,j+1)].t = temp1Type;
            ws[formString(i,j+1)].v = temp1;
            temp1 = temp2;
            temp1Type = temp2Type;
             }
        else
             {   ws[formString(i,j+1)] = {};
                 ws[formString(i,j+1)].t = temp1Type;
                 ws[formString(i,j+1)].v = temp1;
                
             }

         
       }
       else{
        
           if(j==1){
               splitData(ws[formString(i,j)]);
               ws[formString(i,j)].t = "d";
               ws[formString(i,j)].v = tempDate;
               temp1Type = ws[formString(i,j+1)].t;
               temp1 = ws[formString(i,j+1)].v;
               ws[formString(i,j+1)].t = "s";
               ws[formString(i,j+1)].v = tempTime;
               
               
               
           }
           else{
               if(ws[formString(i,j+1)]!=null){
                temp2Type = ws[formString(i,j+1)].t;
               temp2 = ws[formString(i,j+1)].v;
               ws[formString(i,j+1)].t = temp1Type;
               ws[formString(i,j+1)].v = temp1;
               temp1=temp2;
               temp1Type = temp2Type;
               }
               else{
                   ws[formString(i,j+1)]={};
                   ws[formString(i,j+1)].t = temp1Type;
                   ws[formString(i,j+1)].v = temp1;
                   
                   
               }
           }
       }


 }




} 

console.log(ws["I2"].v);
/*console.log(colCount);
let cell = ws['A3'].v.split(',');
var cellParts = cell[0].split('-');
var cellPartsTime = cell[1].split(':');
//console.log(cellPartsTime[0]);
var dateObj = new Date('2019','11' -1,'04',cellPartsTime[0]-6 -5 + 12,cellPartsTime[1]-30 -30,cellPartsTime[2] -10);
var timeObj = cellPartsTime[0] + ':' +  cellPartsTime[1] + ":" + cellPartsTime[2];
ws["A3"] = {
    t:'d',
    v:dateObj
};
ws["B2"]={
    t:'s',
    v:timeObj
};
//console.log(ws['A2']);
//console.log(ws['A3']);
*/

let data = xlsx.utils.sheet_to_json(ws);
var newWb = xlsx.utils.book_new();
var newWs = xlsx.utils.json_to_sheet(data);
xlsx.utils.book_append_sheet(newWb,ws,"New Data");
xlsx.writeFile(newWb,"new DataFIle.xlsx");

function formString( i, j){
    var startAscii = 64;
    var str = String.fromCharCode(startAscii+j) + i.toString();
    return str;

}

function splitData(data){
    
    let cell = data.v.split(','); 
    var cellParts = cell[0].split('-');
    var cellPartsTime = cell[1].split(':');
    var dateObj = new Date(cellParts[0],cellParts[1] -1,cellParts[2],cellPartsTime[0]-6 -5 + 12,cellPartsTime[1]-30 -30,cellPartsTime[2] -10);
    var timeObj = cellPartsTime[0] + ':' +  cellPartsTime[1] + ":" + cellPartsTime[2];

    tempDate = dateObj;
    tempTime = timeObj;
    
}