


////// Function for obtaining the rest

function computeRest() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadSheetFormulas = ss.getSheetByName("HojaFormulas");
    const formatSheet = ss.getSheetByName("FormatoEntregaImplementos");
  
  
    const rangeFormulas = spreadSheetFormulas.getRange("A2:A70");
    const formatFormulas = formatSheet.getRange("F11:F79");
  
    const formulas = rangeFormulas.getFormulas();
  
    formatFormulas.setFormulas(formulas);
  
  }
  
  
  
  //// Function for identifying index of similar element and retun ir
  
  function identifyIndex(array1, array2) {
    var index_array = []
    for(var i = 0; i < array1.length; i += 1) {
      for(var j = 0; j < array2.length; j += 1 ) {
        if(array1[i] === array2[j]) {
          index_array.push(j);
        }
      }
    }
  
    return index_array;
  }
  
  
  
  /// Function for searching if element is or not is in array
  
  function isElementInArray(element, array) {
    for (var i = 0; i < array.length; i++) {
      if (array[i] === element) {
        return true; // Element found in the array
      }
    }
    return false; // Element not found in the array
  }
  
  
  
  /// Function for removing duplicates when merging
  
  function removeDuplicates(inputArray) {
    var uniqueArray = [];
    var seen = {};
  
    for (var i = 0; i < inputArray.length; i++) {
      var value = inputArray[i];
      if (!seen[value]) {
        uniqueArray.push(value);
        seen[value] = true;
      }
    }
    return uniqueArray;
  }
  
  
  
  /// Function for converting a column to an array
  
  function convertToArray(columnValues) {
  
    const array = []
  
    for (var i = 0; i < columnValues.length; i++) {
      var value = columnValues[i][0]; 
      array.push(value); 
    }
  
    return array;
  }
  
  
  
  // Function for converting to 2D array for writing to spreadsheet
  
  function prepareToWrite(originalArray) {
  
    var outerArray = [],
      tempArray = [],
      j=0;
  
    for (j=0; j < originalArray.length; j+=1) {
      tempArray = [];
      tempArray.push(originalArray[j]);
      outerArray.push(tempArray);
    };
  
    return outerArray;
  
  }
  
  
  
  
  // Function for unifying cols 
  
  
  function mergeCols() {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    const dataDOT = ss.getSheetByName("DOTACION");
    const dataEPP = ss.getSheetByName("EPP");
    const auxiliarySheet = ss.getSheetByName("HojaAuxiliar");
  
    const IDDOT = dataDOT.getRange("B2:B")
    const IDEPP = dataEPP.getRange("B3:B")
  
    const IDDOTValues = IDDOT.getValues();
    const IDEPPValues = IDEPP.getValues();
  
    const arrayIDDOT = convertToArray(IDDOTValues);
    const arrayIDEPP = convertToArray(IDEPPValues);
  
    const totalIDs = arrayIDDOT.concat(arrayIDEPP);
  
    const notrepeatedIDs = removeDuplicates(totalIDs);
  
    const joinedCOl = prepareToWrite(notrepeatedIDs);
  
    var rangeJoinedIDs = auxiliarySheet.getRange(2,2, joinedCOl.length, 1 );
    rangeJoinedIDs.setValues(joinedCOl)
  
  }
  
  
  /////////////////////////////////////////////////////////////////////////////////
  
  
  
  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
  
  
  function writeData() {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formatSheet = ss.getSheetByName("FormatoEntregaImplementos");
    const auxiliary = ss.getSheetByName("HojaAuxiliar");
    const auxiliaryWithDates = ss.getSheetByName("RegistroConFechas");
  
    const registerSheet = auxiliary.getRange("A2:EZ1402").getValues();
    const implementNames = auxiliary.getRange(1, 1, 1, auxiliary.getMaxColumns()).getValues()[0]
  
    //getSheetValues(startRow, startColumn, numRows, numColumns)
  
    const employeeID = formatSheet.getRange("D4").getValue();
    const dateOfRegister = formatSheet.getRange("G4").getValue();
    const employeeCargo = formatSheet.getRange("D8").getValue();
  
  
    const allIDs = convertToArray(auxiliary.getRange("B2:B").getValues())
  
    const implementsAsignedtoEmployee  = convertToArray(formatSheet.getRange(11,3,auxiliary.getMaxColumns(),1).getValues());  
    const implementsDeliveredQuantity = convertToArray(formatSheet.getRange(11,7,auxiliary.getMaxColumns(),1).getValues());  
  
  
  
    const noBlanksimplementsAsignedtoEmployee = implementsAsignedtoEmployee.filter((el) => el != '');
    const noBlanksimplementsDeliveredQuantity = implementsDeliveredQuantity.filter((el) => el != '');
  
    // We create an array to store the index of the col associated to the element obtain in format in the auxiliary sheet
    
    var index = identifyIndex(noBlanksimplementsAsignedtoEmployee, implementNames);
  
    var implementIndex = []
  
    for(var i = 0; i < index.length; i+= 1) {
      implementIndex.push(index[i] + 1)
    }
  
    // The row in HojaAuxiliar for the given employeeID
  
    const rowToFIll = allIDs.indexOf(employeeID) + 2 
  
    const registeredEntries = registerSheet.filter(r => r[1] == employeeID)[0]
  
    //console.log(implementIndex)
   // console.log(registeredEntries)
  
    var registeredQuantities = []
  
    for(var i = 0; i < index.length; i+=1) {
      registeredQuantities.push(registeredEntries[index[i]])
    }
  
    var updatedQuantities = []
  
    for(var i = 0; i < index.length; i+=1) {
      updatedQuantities.push(registeredQuantities[i] + noBlanksimplementsDeliveredQuantity[i])
    }
  
    console.log(registeredQuantities)
    console.log(noBlanksimplementsDeliveredQuantity)
  
    // Now we write the imeplents to the datasheet
  
   for(var i = 0; i < implementIndex.length; i += 1) {
      auxiliary.getRange(rowToFIll, implementIndex[i], 1,1).setValue(updatedQuantities[i])
    } 
  
  
  
  //// Now we want to obtain a row with the date 
  
  var size = auxiliaryWithDates.getMaxColumns();
  var infoWithDate = new Array(size);
  
  updatedQuantities.push(dateOfRegister);
  implementIndex.push(size-1);
  
  
  for (var i = 0; i < size; i++) {
    infoWithDate[i] = "";
  }
  
  
  for(var i = 0; i < implementIndex.length; i += 1) {
    infoWithDate[implementIndex[i]] = updatedQuantities[i]
  } 
  
  infoWithDate[1] = employeeID; 
  infoWithDate[2] = employeeCargo;
  
  console.log(infoWithDate);
  auxiliaryWithDates.appendRow(infoWithDate)
   
  }
  
  
  
  
  
  
  
  
  
  
    //console.log(implementIndex)
    // Now we write the imeplents to the datasheet
  
   
  
  
  
  
  
  
  
  ////////////////////////////////////////////// Retrieve en posesion //////////////////////////////////////
  
  function enPosesion() {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formatSheet = ss.getSheetByName("FormatoEntregaImplementos");
    const registerSheet = ss.getSheetByName("HojaAuxiliar");
    
    const employeeID = formatSheet.getRange("D4").getValue();
  
    const implementNames = registerSheet.getRange(1, 1, 1, registerSheet.getMaxColumns()).getValues()[0]
    const wholeRegisteredData = registerSheet.getRange("A2:EZ1402").getValues();
  
  
  
    //getSheetValues(startRow, startColumn, numRows, numColumns)
  
  
    const implementsAsignedtoEmployee  = convertToArray(formatSheet.getRange(11,3,registerSheet.getMaxColumns(),1).getValues());  
    const noBlanksimplementsAsignedtoEmployee = implementsAsignedtoEmployee.filter((el) => el != '');
  
    const employeeRegisteredInfo = wholeRegisteredData.filter(r => r[1] == employeeID)[0];  
  
    //console.log(employeeRegisteredInfo)
  
    var index = identifyIndex(noBlanksimplementsAsignedtoEmployee, implementNames);
    
  
    // Now we search for the quantity related to the element corresponding to the index
  
    var quantitiesInPosession = []
  
    for(var i = 0; i < index.length; i+=1) {
      quantitiesInPosession.push(employeeRegisteredInfo[index[i]])
    }
  
    //console.log(quantitiesInPosession)
  
    // We create an array to store the index of the col associated to the element obtain in format in the auxiliary sheet
  
    var quantitiesToWrite = prepareToWrite(quantitiesInPosession)
  
  
    var rangeEnPosession = formatSheet.getRange(11,5, index.length,1);
  
    rangeEnPosession.setValues(quantitiesToWrite);
  
  
  
  }
  
  
  
  ///////////////////////////////////////////////New Functions for the formats //////////////////////////////
  
  
  function evaluateRetrieveEmployeeInfoCase() {
   
   const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //  En premieur lieu on trouve l'information sur la dotacion
  
    const dotSheet = ss.getSheetByName("DOTACION");
    const idsDOT = convertToArray(dotSheet.getRange("B2:B").getValues());
  
  // Après on cherche l'information sur les EPP
  
    const eppSheet = ss.getSheetByName("EPP");  
    const idsEPP = convertToArray(eppSheet.getRange("B3:B").getValues());
  
    const formatDelivery = ss.getSheetByName("FormatoEntregaImplementos");
    const employeeID = formatDelivery.getRange("D4").getValue();    
  
    var ui = SpreadsheetApp.getUi();
    
    if(formatDelivery.getRange("D4").isBlank() == true){
      ui.alert('Por favor ingrese la identificación del empleado.');
      return false;
    }
  
    if(isElementInArray(employeeID, idsDOT) & isElementInArray(employeeID, idsEPP)){
      retrieveEmployeeInfoCase1();
    } else if (isElementInArray(employeeID, idsDOT) == true & isElementInArray(employeeID, idsEPP) == false) {
      retrieveEmployeeInfoCase2();
    } else if (isElementInArray(employeeID, idsDOT) == false & isElementInArray(employeeID, idsEPP) == true){
      retrieveEmployeeInfoCase3();
    }
  
    enPosesion();
  
    computeRest();
  }
  
  
  /////////////////////////////////////////
  /////// Functions for the formats ///////
  /////////////////////////////////////////
  
  
  function retrieveEmployeeInfoCase1() {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //  En premieur lieu on trouve l'information sur la dotacion
  
    const dotSheet = ss.getSheetByName("DOTACION");
    const dataDOT =  dotSheet.getRange("A1:AZ1396").getValues();
    const implementNamesDOT = dotSheet.getRange("A1:AZ1").getValues();
    
  
  // Après on cherche l'information sur les EPP
  
    const eppSheet = ss.getSheetByName("EPP");  
    const dataEPP =  eppSheet.getRange("A2:DG1199").getValues();
    const implementNamesEPP = eppSheet.getRange("A2:DG2").getValues();
  
  /////////
    const formatDelivery = ss.getSheetByName("FormatoEntregaImplementos");
    const employeeID = formatDelivery.getRange("D4").getValue();    
  
  
  /////////////// We do all the process for DOTACION ///////////////
  
    const employeeInfoDOT = dataDOT.filter(r => r[1] == employeeID);  
    const employeeInfoArray1DOT = employeeInfoDOT[0];
    const implementNamesArrayDOT = implementNamesDOT[0];
    const employeCharge = employeeInfoDOT[0][2]
  
    var notblankindexDOT = []
    var resultingImplementsDOT = []
  
     for(i = 0; i < employeeInfoArray1DOT.length; i+=1) {
      if (employeeInfoArray1DOT[i] != ''){
        notblankindexDOT.push(i);
      }
     }
  
     for (i = 0; i < notblankindexDOT.length; i+=1){
      resultingImplementsDOT.push(implementNamesArrayDOT[notblankindexDOT[i]]); 
     }
  
  
    employeeInfoDOT[0].splice(0,7);
    resultingImplementsDOT.splice(0,7);
  
    const eraseblanksDOT = employeeInfoArray1DOT.filter((el) => el != '');
  
    const employeeInfoArrayDOT = eraseblanksDOT;
  
    //Write the name of the elements from the DOT list associated to the employee
    var outerArrayDOT = [],
      tempArrayDOT = [],
      j=0;
  
    for (j=0; j < resultingImplementsDOT.length; j+=1) {
      tempArrayDOT = [];
      tempArrayDOT.push(resultingImplementsDOT[j]);
      outerArrayDOT.push(tempArrayDOT);
    };
  
  
    //////////////////////////////////////////////////////////////////
  
      //Write the quantity of elements to the forms list
    var outerArrayDOTQ = [],
      tempArrayDOTQ = [],
      i=0;
  
    for (i=0; i < employeeInfoArrayDOT.length; i+=1) {
      tempArrayDOTQ = [];
      tempArrayDOTQ.push(employeeInfoArrayDOT[i]);
      outerArrayDOTQ.push(tempArrayDOTQ);
    };
  
  
  
  
  
   ////////////////////////////////////////////////////////////////////////////////
  
   // Now we do all the process for EPP /////////////////////
  
  
    const employeeInfoEPP = dataEPP.filter(r => r[1] == employeeID);
  
      
    const employeeInfoArray1EPP = employeeInfoEPP[0];
    const implementNamesArrayEPP = implementNamesEPP[0];
  
    // This array will have the index of the elements different from blank
     
    var notblankindexEPP = []
    var resultingImplementsEPP = []
  
     for(i = 0; i < employeeInfoArray1EPP.length; i+=1) {
      if (employeeInfoArray1EPP[i] != ''){
        notblankindexEPP.push(i);
      }
     }
  
     for (i = 0; i < notblankindexEPP.length; i+=1){
      resultingImplementsEPP.push(implementNamesArrayEPP[notblankindexEPP[i]]); 
     }
  
    // That was just a test
  
    employeeInfoEPP[0].splice(0,7);
  
    resultingImplementsEPP.splice(0,7);
  
    const eraseblanksEPP = employeeInfoArray1EPP.filter((el) => el != '');
  
    const employeeInfoArrayEPP = eraseblanksEPP;
  
    //Write the name of the elements from the EPP list associated to the employee
    var outerArrayEPP = [],
      tempArrayEPP = [],
      j=0;
  
    for (j=0; j < resultingImplementsEPP.length; j+=1) {
      tempArrayEPP = [];
      tempArrayEPP.push(resultingImplementsEPP[j]);
      outerArrayEPP.push(tempArrayEPP);
    };
   
  
   //Write the quantity of elements to the forms list
    var outerArrayEPPQ = [],
      tempArrayEPPQ = [],
      i=0;
  
    for (i=0; i < employeeInfoArrayEPP.length; i+=1) {
      tempArrayEPPQ = [];
      tempArrayEPPQ.push(employeeInfoArrayEPP[i]);
      outerArrayEPPQ.push(tempArrayEPPQ);
    };
  
  
  
  
  
  
  /// The total names of implements 
  
    var totalArrayNames = outerArrayDOT.concat(outerArrayEPP);
  
    var outerArrayNames = [],
      tempArrayNames = [],
      i=0;
  
    for (i=0; i < totalArrayNames.length; i+=1) {
      tempArrayNames = [];
      tempArrayNames.push(totalArrayNames[i]);
      outerArrayNames.push(tempArrayNames);
    };
  
  
  // The total amount of each implement
  
  var totalArrayQuantity = outerArrayDOTQ.concat(outerArrayEPPQ);
  
   var outerArrayQ = [],
      tempArrayQ = [],
      i=0;
  
    for (i=0; i < totalArrayQuantity.length; i+=1) {
      tempArrayQ = [];
      tempArrayQ.push(totalArrayQuantity[i]);
      outerArrayQ.push(tempArrayQ);
    };
  
  
  
  
    /// For writing to the FormatoSheet
    
  
    var rangeNames = formatDelivery.getRange(11,3, totalArrayNames.length, 1 );
    rangeNames.setValues(totalArrayNames);  
  
    var rangeQuantities = formatDelivery.getRange(11,4, totalArrayQuantity.length, 1 );
    rangeQuantities.setValues(totalArrayQuantity);
  
  
    var typeDOTARRAY = [];
    var typeEPPARRAY = [];
  
    var stringDOT = "DOT";
    var stringEPP = "EPP";
  
    for(var i = 0; i < outerArrayDOTQ.length; i += 1) {
      typeDOTARRAY.push(stringDOT)
    } 
  
    for(var i = 0; i < outerArrayEPPQ.length; i += 1) {
      typeEPPARRAY.push(stringEPP);
    } 
  
    var etiquetes = typeDOTARRAY.concat(typeEPPARRAY);
  
    var columnEtiquetes = prepareToWrite(etiquetes);
  
    var rangeEtiquetes = formatDelivery.getRange(11,2, columnEtiquetes.length,1);
  
    rangeEtiquetes.setValues(columnEtiquetes);
  
    formatDelivery.getRange("D8").setValue(employeCharge);
  
  }
  
  
  //// Now we go for the second case (only DOT) //// 
  
  function retrieveEmployeeInfoCase2() {
  
   const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //  En premieur lieu on trouve l'information sur la dotacion
  
    const dotSheet = ss.getSheetByName("DOTACION");
    const dataDOT =  dotSheet.getRange("A1:AZ1396").getValues();
    const implementNamesDOT = dotSheet.getRange("A1:AZ1").getValues();
  
    const formatDelivery = ss.getSheetByName("FormatoEntregaImplementos");
    const employeeID = formatDelivery.getRange("D4").getValue();    
  
  
  /////////////// We do all the process for DOTACION ///////////////
  
    const employeeInfoDOT = dataDOT.filter(r => r[1] == employeeID);  
    const employeeInfoArray1DOT = employeeInfoDOT[0];
    const implementNamesArrayDOT = implementNamesDOT[0];
    const employeCharge = employeeInfoDOT[0][2]
  
    var notblankindexDOT = []
    var resultingImplementsDOT = []
  
     for(i = 0; i < employeeInfoArray1DOT.length; i+=1) {
      if (employeeInfoArray1DOT[i] != ''){
        notblankindexDOT.push(i);
      }
     }
  
     for (i = 0; i < notblankindexDOT.length; i+=1){
      resultingImplementsDOT.push(implementNamesArrayDOT[notblankindexDOT[i]]); 
     }
  
  
    employeeInfoDOT[0].splice(0,7);
    resultingImplementsDOT.splice(0,7);
  
    const eraseblanksDOT = employeeInfoArray1DOT.filter((el) => el != '');
  
    const employeeInfoArrayDOT = eraseblanksDOT;
  
    //Write the name of the elements from the DOT list associated to the employee
   
    var outerArrayDOT = prepareToWrite(resultingImplementsDOT);
  
   //Write the quantity of elements to the forms list
  
    var outerArrayDOTQ = prepareToWrite(employeeInfoArrayDOT);
  
  
    var rangeNames = formatDelivery.getRange(11,3, outerArrayDOT.length, 1 );
    rangeNames.setValues(outerArrayDOT);  
  
    var rangeQuantities = formatDelivery.getRange(11,4, outerArrayDOTQ.length, 1 );
    rangeQuantities.setValues(outerArrayDOTQ);
  
  
    var typeDOTARRAY = [];
    var stringDOT = "DOT";
  
    for(var i = 0; i < outerArrayDOTQ.length; i += 1) {
      typeDOTARRAY.push(stringDOT)
    } 
  
    var columnEtiquetes = prepareToWrite(typeDOTARRAY);
  
    var rangeEtiquetes = formatDelivery.getRange(11,2, columnEtiquetes.length,1);
  
    rangeEtiquetes.setValues(columnEtiquetes);
    formatDelivery.getRange("D8").setValue(employeCharge);
  }
  
  ///// Now we go for the third case (Only EPP) //// 
  
  function retrieveEmployeeInfoCase3() {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Après on cherche l'information sur les EPP
  
    const eppSheet = ss.getSheetByName("EPP");  
    const dataEPP =  eppSheet.getRange("A2:DG1199").getValues();
    const implementNamesEPP = eppSheet.getRange("A2:DG2").getValues();
  
    const formatDelivery = ss.getSheetByName("FormatoEntregaImplementos");
    const employeeID = formatDelivery.getRange("D4").getValue();    
  
  
  
    const employeeInfoEPP = dataEPP.filter(r => r[1] == employeeID);
  
      
    const employeeInfoArray1EPP = employeeInfoEPP[0];
    const implementNamesArrayEPP = implementNamesEPP[0];
    const employeCharge = employeeInfoEPP[0][2]
  
  
    // This array will have the index of the elements different from blank
     
    var notblankindexEPP = []
    var resultingImplementsEPP = []
  
     for(i = 0; i < employeeInfoArray1EPP.length; i+=1) {
      if (employeeInfoArray1EPP[i] != ''){
        notblankindexEPP.push(i);
      }
     }
  
     for (i = 0; i < notblankindexEPP.length; i+=1){
      resultingImplementsEPP.push(implementNamesArrayEPP[notblankindexEPP[i]]); 
     }
  
    // That was just a test
  
    employeeInfoEPP[0].splice(0,7);
  
    resultingImplementsEPP.splice(0,7);
  
    const eraseblanksEPP = employeeInfoArray1EPP.filter((el) => el != '');
  
    const employeeInfoArrayEPP = eraseblanksEPP;
  
    //Write the name of the elements from the EPP list associated to the employee
    var outerArrayEPP = prepareToWrite(resultingImplementsEPP);
   
  
   //Write the quantity of elements to the forms list
     var outerArrayEPPQ = prepareToWrite(employeeInfoArrayEPP);
  
  
  
    var rangeNames = formatDelivery.getRange(11,3, outerArrayEPP.length, 1 );
    rangeNames.setValues(outerArrayEPP);  
  
    var rangeQuantities = formatDelivery.getRange(11,4, outerArrayEPPQ.length, 1 );
    rangeQuantities.setValues(outerArrayEPPQ);
  
  
    var typeEPPARRAY = [];
    var stringEPP = "EPP";
  
    for(var i = 0; i < outerArrayEPPQ.length; i += 1) {
      typeEPPARRAY.push(stringEPP)
    } 
  
    var columnEtiquetes = prepareToWrite(typeEPPARRAY);
  
    var rangeEtiquetes = formatDelivery.getRange(11,2, columnEtiquetes.length,1);
  
    rangeEtiquetes.setValues(columnEtiquetes);
  
    formatDelivery.getRange("D8").setValue(employeCharge);
  }
  
  
  function cleanForm() {
  
    ss = SpreadsheetApp.getActiveSpreadsheet();
    const formatDelivery = ss.getSheetByName("FormatoEntregaImplementos");
    
    const idCell = formatDelivery.getRange("D4");
    const dateCell = formatDelivery.getRange("D6");
    const range = formatDelivery.getRange("B11:G77");
    const cargo = formatDelivery.getRange("D8");
  
    idCell.clearContent();
    dateCell.clearContent();
    range.clearContent();
    cargo.clearContent();
  
  }
  
  
  function searchSTOCK () {
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stockFormat = ss.getSheetByName("FormatoConsultaStock");
    const registerFormat = ss.getSheetByName("HojaAuxiliar")
    const stock = ss.getSheetByName("Stock");
    const formulasSheet = ss.getSheetByName("HojaFormulas");
  
    const implements = registerFormat.getRange("A1:EZ1").getValues()[0];
    const specificImplement = stockFormat.getRange("D4").getValue();
    
    const IDs = convertToArray(registerFormat.getRange("B2:B").getValues()); 
  
  
    const auxiliaryList = [specificImplement];
  
    const index = identifyIndex(auxiliaryList, implements);
  
    const realIndex = index[0] + 1;
  
    const columnOfSpecificElement = convertToArray(registerFormat.getRange(2, realIndex,IDs.length ,1).getValues());
  
    var sum = 0.0;
  
    for(var i = 0; i < columnOfSpecificElement.length; i += 1) {
      sum += columnOfSpecificElement[i]
    }
  
    stockFormat.getRange("D7").setValue(sum);
    
  
  // Now we search for the elements in stock
  
    const imeplentsNamesInStock = stock.getRange(1,1,1,stock.getMaxColumns()).getValues()[0];
    const implementsQuantityStock = stock.getRange(2,1,1, stock.getMaxColumns()).getValues()[0];
  
    const indexToSearchStock = imeplentsNamesInStock.indexOf(specificImplement);
    const specificImplementStock = implementsQuantityStock[indexToSearchStock];
  
    stockFormat.getRange("C7").setValue(specificImplementStock);
    
  
  
    const formulas = formulasSheet.getRange("C2").getFormula();
    
    stockFormat.getRange("E7").setFormula(formulas);
  
    //formatFormulas.setFormulas(formulas);
  
  }
  
  
  function cleanStock(){
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stockFormat = ss.getSheetByName("FormatoConsultaStock");
  
    stockFormat.getRange("D4").clearContent();
    stockFormat.getRange("C7").clearContent();
    stockFormat.getRange("D7").clearContent();
    stockFormat.getRange("E7").clearContent();
  
  
  }
  
  
  
  // Now we have the STOCK
  
  
  function refillStock()
  {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet = ss.getSheetByName("Stock");
    const formatoConsultaStuck = ss.getSheetByName("FormatoConsultaStock")
  
  
    var implementsStock = []
  
    for(var i = 0; i < stockSheet.getMaxColumns(); i += 1){
  
      const min = 10;
      const max = 2000;
      implementsStock[i] = Math.floor(Math.random() * (max - min) + min);
    }
  
    var rangeToFill = stockSheet.getRange(2,1, 1,stockSheet.getMaxColumns()).setValues([implementsStock])
  
  
  
  }
  
  
  
  
  
  
  
  
  
  
  