//global var
var app = SpreadsheetApp.openById('1Np3UhwCmN9hqWHGrhrFOIxteeFaDkU55q0lWurSE_EQ');
var reconSheet = app.getSheetByName('Recon');
var uploadSheet = app.getSheetByName('Upload');

function reconList() {
  
  reconSheet.clear();

  reconSheet.getRange(1,1).setFormula('=UNIQUE(Upload!B:B)'); //get unique company
  reconSheet.getRange(2,2).setFormula('=IFNA(VLOOKUP(A2,UNIQUE({Upload!B:B,Upload!A:A}),2,0))')
    .copyTo(reconSheet.getRange('Recon!B3:B'),SpreadsheetApp.CopyPasteType.PASTE_FORMULA);//get company PIC
  reconSheet.getRange(1,2).setValue('Name');
  reconSheet.getRange(2,3).setFormula('=IFNA(VLOOKUP(A2,UNIQUE(Upload!B:K),10,0))')
    .copyTo(reconSheet.getRange('Recon!C3:C'),SpreadsheetApp.CopyPasteType.PASTE_FORMULA);//get PIC email
  reconSheet.getRange(1,3).setValue('Email');

  // get end date, plan type, location, expiry status & remarks
  reconSheet.getRange(2,4).setFormula('=ArrayFormula(IFNA(VLOOKUP(A2,Upload!B:J,{3,4,8,9},0)))')
    .copyTo(reconSheet.getRange('Recon!D3:D'),SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  reconSheet.getRange(1,4).setValue('End Date');
  reconSheet.getRange(1,5).setValue('Plan Type');
  reconSheet.getRange(1,6).setValue('Location');
  reconSheet.getRange(1,7).setValue('Expiry Status');
  reconSheet.getRange(1,8).setValue('Remarks');

  reconSheet.autoResizeColumns(1,15);

  //filterTable(); 
}

function filterTable(){
  //create new array to remove blank end date and old date less than TODAY()
  let allArrDatas = reconSheet.getRange(1,1,reconSheet.getMaxRows(),reconSheet.getLastColumn()).getValues();
  var colNum = 3;// choose the column to check : 0=A, 1=B, 2=C etc... 
  var colNum2 = 6;
  var targetData = new Array();
  for(n=0;n<allArrDatas.length;++n){
    if(allArrDatas[n][colNum]!='' && allArrDatas[n][colNum2]!='Expired' && allArrDatas[n][colNum2]!='Not Expiring'){ targetData.push(allArrDatas[n])};
    // check the cell for "not empty" (does not detect formulas !!)
  }
  reconSheet.getDataRange().clear();
  reconSheet.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);

  reconSheet.autoResizeColumns(1,15);
  reconSheet.getRange('Recon!A2:G').sort({column:4,ascending:true});

  let lrow = reconSheet.getLastRow();

  for ( let i = 2 ; i <= lrow ; i++){

   var oldDate = reconSheet.getRange(i , 4).getValue().slice(0, 10).split('-');

   var _date =oldDate[2] +'-'+ oldDate[1] +'-'+ oldDate[0];

   reconSheet.getRange(i , 4).setValue(_date);

  }

}
