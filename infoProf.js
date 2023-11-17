////////////////////////////////////// Functions for retrieving the info /////////////////////////////////////////


const firstSetofYears = SpreadsheetApp.openById("1HbOLvAWargyzhcjK7FVd49ZBUoQj0VUfgVi__HNS2GI");
const secondSetofYears = SpreadsheetApp.openById("1ngWYZ-mrTysjpSeaTYEFR5z1AaS33cXhhHI-k2YQvSk");



const arrayofYears = []
const arrayNamesOfYears = []
const arrayInfoOfYears = []

for(var i = 2008; i < 2023; i += 1) {
  arrayofYears.push(i);
}

for(var i = 0; i < arrayofYears.length; i += 1) {
  arrayNamesOfYears.push(arrayofYears[i].toString());
}

for(var j = 0; j < 7; j += 1) {
  arrayInfoOfYears.push(firstSetofYears.getSheetByName(arrayNamesOfYears[j]).getRange("A:Z").getValues());
}

for(var j = 7; j < arrayofYears.length; j += 1) {
  arrayInfoOfYears.push(secondSetofYears.getSheetByName(arrayNamesOfYears[j]).getRange("A:Z").getValues());
}



/////////////////////////////////////////// Functions for extracting the info //////////////////////////////


function recoverInfo() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchFormat = ss.getSheetByName("Consulta");

  const profID = searchFormat.getRange("D5").getValue();

  const activityType = ss.getSheetByName("HojaAuxiliar").getRange("C2:J2").getValues()[0];
  const activityTypeSelection = searchFormat.getRange("D8:K8").getValues()[0];
  const activities = neededActivities(activityTypeSelection,activityType);


  const columnNames = ss.getSheetByName("HojaAuxiliar2").getRange("A1:Z1").getValues()[0];
  const columnsWanted = searchFormat.getRange("D10:R10").getValues()[0].filter((el) => el != '');
  const columnsWantedIndex = colIndex(columnsWanted, columnNames);

  //console.log(columnsWanted)

  const indexID = 3;


  //console.log(activities)
  //console.log(columnsWanted)
  //console.log(columnNames)
  //console.log(columnsWantedIndex)

  const yearsToSearch = searchFormat.getRange("D3:Q3").getValues()[0].filter((el) => el != '');

  var arrayOfIndecesForYears = [];
  var differentYearsInfo = [];
  var filteredDifferentYearsInfo = [];

    //differentYearsInfo is already filtered

  for(var i = 0; i < yearsToSearch.length; i +=1) {
    arrayOfIndecesForYears.push(arrayofYears.indexOf(yearsToSearch[i]));
  }

  const indecesLen = arrayOfIndecesForYears.length;

  for(var i = 0; i < indecesLen; i+=1) {
    differentYearsInfo.push(arrayInfoOfYears[arrayOfIndecesForYears[i]].filter(r => r[3] == profID & activities.includes(r[8])));
  }


  var colsWantedinfoProf = [];
  var littleInfo = [];
  var numColsWanted = columnsWanted.length;


    //// algo de pseudocodigo

  for(var i = 0; i < differentYearsInfo.length; i += 1){
    for(var j = 0; j < differentYearsInfo[i].length; j += 1){
      var temporary = [];
      for(var k = 0; k < numColsWanted; k += 1) {
        littleInfo.push(differentYearsInfo[i][j][columnsWantedIndex[k]]);
      }

      colsWantedinfoProf.push(littleInfo);
      littleInfo = [];
    }
  }

   // Now we write to the search format

  searchFormat.getRange(13,3, 1, columnsWanted.length).setFontFamily("Times New Roman");
  searchFormat.getRange(13,3, 1, columnsWanted.length).setFontWeight("bold");
  searchFormat.getRange(13,3, 1, columnsWanted.length).setValues([columnsWanted]);

    // Now we write the information


    //console.log(colsWantedinfoProf.length)
   // console.log(colsWantedinfoProf)

  for(var i = 0; i < colsWantedinfoProf.length; i += 1) {
     searchFormat.getRange(14+i,3, 1, columnsWanted.length).setValues([colsWantedinfoProf[i]]);
  }







    //Este si funciona
    //searchFormat.getRange(14,3, 1, columnsWanted.length).setValues([colsWantedinfoProf[0]])

}



function cleanSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchFormat = ss.getSheetByName("Consulta");

  searchFormat.getRange("D5").clearContent();
  searchFormat.getRange("C13:Z385").clearContent();
  searchFormat.getRange("D3:R3").clearContent();

}







////////////////////////////// Functions for doing other stuff ////////////////////////////////

function neededActivities(array1, array2) {

  var values = [];

  for(var i = 0; i < array1.length; i += 1) {
    if (array1[i] === true) {
      values.push(array2[i])
    }
  }

  return values;
}


function colIndex(colsWanted, colsNames) {

  var indeces = [];

  for(var i = 0; i < colsWanted.length; i += 1) {
    indeces.push(colsNames.indexOf(colsWanted[i]))
  }

  return indeces;
}
