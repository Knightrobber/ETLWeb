const express = require('express');
const app = express();
const morgan = require('morgan');
const bodyParser = require('body-parser')
let xlsx = require('xlsx');
var rowCount =0;
var colCount =0;
var startAscii = 65;
var tempString = 'A1';
var theString = "A1";
var iteratorString = "A1";
var formidable = require('formidable');
var tempDate,tempTime;


app.use(morgan('short'));
app.use(bodyParser.urlencoded({extended:false}))
app.use(express.static('./public'));
app.listen(8000,()=>{
    console.log('listening on port 8000');
})

app.post("/processFile",(req,res)=>{ //  The file is processed here 
    console.log("starting file procession");
    var dataFile = req.body.dataFileFromHTML;
    
   



let wb = xlsx.readFile(dataFile);
let ws = wb.Sheets[wb.SheetNames[0]];
var rowCount =0;
var colCount =0;
var startAscii = 65;
var tempString = 'A1';
var theString = "A1";
var iteratorString = "A1";


/*splitData(ws["A2"]);
console.log("hey")
console.log(tempTime)
console.log(ws["A2"]);
ws["A2"].t = "d"
ws["A2"].v = tempDate
ws["B2"].t = "s"
ws["B2"].v = tempTime
*/


while(ws[tempString]!=null){ // counting no of rows
    ++rowCount;
    tempString = 'A' + rowCount.toString();
}

tempString = 'A1';
while(ws[tempString]!=null){ // coiunting no of columns 
    ++colCount;++startAscii;
    tempString = String.fromCharCode(startAscii) + "1";

}
rowCount = rowCount-1; 


for(var i=1;i<=rowCount;++i){
    
 var temp1,temp2,temp1Type,temp2Type;
 for(var j=1;j<=colCount;++j){
     if(i==1) // In the first row "Time column has to be inserted"
     {
         if(j==1)
         continue;
         else if(j==2){ // 2nd column becomes time
             temp1 = ws[formString(i,j)].v; // to access a cell - ws["A1"], form string forms a string based on the value of the row and column given to it.
             temp1Type = ws[formString(i,j)].t;
             
             ws[formString(i,j)].t = "s"; // "s" specifies that its a string
             ws[formString(i,j)].v = "Time";
             
             
            
             
         }
         
        if(ws[formString(i,j+1)]!=null){ // making new 3rd column equal to the 2nd column of the input file
            temp2Type = ws[formString(i,j+1)].t
            temp2 = ws[formString(i,j+1)].v; 
            ws[formString(i,j+1)].t = temp1Type;
            ws[formString(i,j+1)].v = temp1;
            temp1 = temp2;
            temp1Type = temp2Type;
             }
        else
             {   
                add_cell_to_sheet(ws,formString(i,j+1),temp1,temp1Type);     
             }

         
       }
       else{ // for all rows after the first row
        
           if(j==1){
               splitData(ws[formString(i,j)]); // splitting data of first column
               ws[formString(i,j)].t = "d"; // specifies its a date
               ws[formString(i,j)].v = tempDate;
               temp1Type = ws[formString(i,j+1)].t;
               temp1 = ws[formString(i,j+1)].v;
               ws[formString(i,j+1)].t = "s"; // making the 2nd column as time
               ws[formString(i,j+1)].v = tempTime;
               
               
               
           }
           else{
               if(ws[formString(i,j+1)]!=null){
                temp2Type = ws[formString(i,j+1)].t;// making 3rd column equal to the 2nd column of the input file
               temp2 = ws[formString(i,j+1)].v;
               ws[formString(i,j+1)].t = temp1Type;
               ws[formString(i,j+1)].v = temp1;
               temp1=temp2;
               temp1Type = temp2Type;
               }
               else{
                add_cell_to_sheet(ws,formString(i,j+1),temp1,temp1Type);    
                   
                   
               }
           }
       }


 }




} 

// creating a new file with updated cell values

var newWb = xlsx.utils.book_new();

xlsx.utils.book_append_sheet(newWb,ws,"New Data");

xlsx.writeFile(newWb,"new Data.xlsx");
var data = xlsx.utils.sheet_to_json(newWb.Sheets['New Data']);
console.log(data[2])



    res.end()
})



function formString( i, j){ // forming a string given the values of column and row
    var startAscii = 64;
    var str = String.fromCharCode(startAscii+j) + i.toString();
    return str;

}

function splitData(data){ // splits data
    
    let cell = data.v.split(','); 
    var cellParts = cell[0].split('-');
    var cellPartsTime = cell[1].split(':');
    var dateObj = new Date(cellParts[0],cellParts[1] -1,cellParts[2],cellPartsTime[0]-6 -5 + 12,cellPartsTime[1]-30 -30,cellPartsTime[2] -10);
    var timeObj = cellPartsTime[0] + ':' +  cellPartsTime[1] + ":" + cellPartsTime[2];

    tempDate = dateObj;
    tempTime = timeObj;
    

    
}

function add_cell_to_sheet(worksheet, address, value,type) {
	/* cell object */
	var cell = {t:'?', v:value};

	/* assign type */
	cell.t = type
	/* add to worksheet, overwriting a cell if it exists */
	worksheet[address] = cell;

	/* find the cell range */
	var range = xlsx.utils.decode_range(worksheet['!ref']);
	var addr = xlsx.utils.decode_cell(address);

	/* extend the range to include the new cell */
	if(range.s.c > addr.c) range.s.c = addr.c;
	if(range.s.r > addr.r) range.s.r = addr.r;
	if(range.e.c < addr.c) range.e.c = addr.c;
	if(range.e.r < addr.r) range.e.r = addr.r;

	/* update range */
	worksheet['!ref'] = xlsx.utils.encode_range(range);
}