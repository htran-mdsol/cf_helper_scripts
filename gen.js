const xlsx = require("xlsx");

// read xls file
var wb = xlsx.readFile('./4858-202_Patient_Cloud_Draft_2.0.xlsx');

// select sheet by name
var sh = wb.Sheets["Fields"];

// convert A1 to range format
var range = xlsx.utils.decode_range(sh['!ref']);

// r is row, c is column, s is first cell, e is last cell
for(var R = range.s.r; R <= range.e.r; ++R) {
    for(var C = range.s.c; C <= range.e.c; ++C) {

      var cellref = xlsx.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
      if(!sh[cellref]) continue; // if cell doesn't exist, move on
      var cell = sh[cellref];

      // if form oid != ECOA_SF36, move on
      if (C === 0 && cell.v !== "ECOA_SF36")
        break;
        
      // print
      if (C === 1) {
        //console.log(`const string FIELD_OID_${cell.v} = "${cell.v}"`);
	console.log(`||DataPoint|${cell.v}||ECOA_SF36|${cell.v}|0||||||\nIsPresent|||||||||||||\nOr|||||||||||||`);
        break;
      }
    }
};

