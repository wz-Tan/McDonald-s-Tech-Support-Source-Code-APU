let invoicesheetcolumnwidth=[70,200,500,180,200,750];
let receiptsheetcolumnwidth=[70,200,500,180,200,300,750];
let tablewidthlist=[50,200,75,75,75];

let headerlist=["Index","Product Name","Price(RM)","Quantity","Total(RM)"];
let questions=["product name","price","quantity"];

let headerStyle = {};  
headerStyle[DocumentApp.Attribute.FONT_SIZE]=14;
headerStyle[DocumentApp.Attribute.BOLD]=true;
headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
headerStyle[DocumentApp.Attribute.FONT_FAMILY]="Inconsolata";

let bodyStyle = {};  
bodyStyle[DocumentApp.Attribute.FONT_SIZE]=14;
bodyStyle[DocumentApp.Attribute.BOLD]=false;
bodyStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
bodyStyle[DocumentApp.Attribute.FONT_FAMILY]="Inconsolata";

let textStyle = {};  
textStyle[DocumentApp.Attribute.FONT_SIZE]=14;
textStyle[DocumentApp.Attribute.BOLD]=false;
textStyle[DocumentApp.Attribute.FONT_FAMILY]="Inconsolata";

let boldtextstyle={};
boldtextstyle[DocumentApp.Attribute.FONT_FAMILY]="Inconsolata";
boldtextstyle[DocumentApp.Attribute.FONT_SIZE]=14;
boldtextstyle[DocumentApp.Attribute.BOLD]=true;

function onOpen() {
  let ui = DocumentApp.getUi();
  ui.createMenu('AIRI').addItem('Invoice', 'CreateInvoice').addItem('Receipt','CreateReceipt').addToUi();
}

function CreateInvoice() {
  let ui = DocumentApp.getUi();
  let doc = DocumentApp.getActiveDocument();
  let body = doc.getBody();
  let documentid=ui.prompt("What is your Document ID?").getResponseText();
  let documentidlength=documentid.length;
  let clientname=ui.prompt("What is your Client's Name?").getResponseText().toUpperCase();
  let duedate=ui.prompt("What is your Due Date? (Answer in DD/MM/YY Format)").getResponseText();
  let date=Date().slice(4,21).toUpperCase();
  let copytext="COPY";
  let dummytext="";

  //Adding format at top
  let firsttext=[copytext.padEnd(53-"0"-documentidlength)+"INVOICE NO:"+documentid+"\n\n\n",
  "                     (Your Company Name Here)                     ",
  "                          (Address Here)                          ",
  "                             INVOICE                              ",
  "TO,",
  clientname,
  "\n                                          DATE:"+date,
  "                                          DUE DATE:"+duedate+"\n"
  ];
  let textlength=firsttext.length;
  let paragraph;
  for (let i=0;i<textlength;i++){
    body.appendParagraph(firsttext[i]).setAttributes(textStyle);
  }

  //Add table
  let productsum=MakeTable();

  //Bottom text
  let bottomtext="\n\n\n\n................			     	           .................\nCLIENT SIGNATURE					      COMPANY SIGNATURE";
  body.appendParagraph(bottomtext).setAttributes(textStyle);

  let datalist=[documentid,clientname,date,productsum];
  let transferrequest=ui.prompt("Would you like to transfer this data into your Invoice Spreadsheet? (Yes/No)").getResponseText().toLowerCase();
  if (transferrequest==="yes") TransferToInvoiceSheet(datalist);
}

function MakeTable(){
  //Create Table
  let doc=DocumentApp.getActiveDocument();
  let body=doc.getBody();
  let newtable=body.appendTable();

  //Create First Row
  let row=newtable.appendTableRow();
  for (let i=0;i<5;i++){
    let cell=row.appendTableCell(headerlist[i]);
    newtable.getRow(0).getCell(i).getChild(0).asParagraph().setAttributes(headerStyle);
    newtable.setColumnWidth(i,tablewidthlist[i]);
  }

  //Ask how many rows
  let ui=DocumentApp.getUi();
  let counter=QuestionAndCheck(ui,"How many rows would you like to have? ",false);

  //Create next lines(Adding data)
  let productsum=0;
  let price=0;
  let quantity=0;
  let quantitysum=0;
  let name;
  for (let i=1;i<counter+1;i++){
    row=newtable.appendTableRow();
    //Index column
    row.appendTableCell(i.toFixed(0));
    newtable.getRow(i).getCell(0).getChild(0).asParagraph().setAttributes(bodyStyle);

    //Product Name column(string-don't mix with numbers)
    prompt=ui.prompt("What is your "+questions[0]+" for product number "+i+" ?");
    let response=prompt.getResponseText();
    name=response;
    cell=row.appendTableCell(name);
    newtable.getRow(i).getCell(1).getChild(0).asParagraph().setAttributes(bodyStyle);

    //Rest of the columns(quantity,price)
    let floatcheck=true;
    for (let j=1;j<3;j++){
      //Check quantity is NOT float
      if (j==1) floatcheck=true;
      if (j==2) floatcheck=false;

      response=QuestionAndCheck(ui,"What is your "+questions[j]+" for product "+ name +"?",floatcheck); 
      cell=(j==1)? row.appendTableCell(response.toFixed(2)) : row.appendTableCell(response.toFixed(0));
      if (j==1) price=response.toFixed(2);
      else if (j==2) quantity=response;
      newtable.getRow(i).getCell(j+1).getChild(0).asParagraph().setAttributes(bodyStyle);
    }
    row.appendTableCell((quantity*price).toFixed(2));
    newtable.getRow(i).getCell(4).getChild(0).asParagraph().setAttributes(bodyStyle);
    productsum+=(quantity*price);   
    quantitysum+=(quantity);
  }

  //Adding discount and Total Sum
  row=newtable.appendTableRow();
  for (let i=0;i<2;i++){
    row.appendTableCell();
  }
  row.appendTableCell("GRAND TOTAL");
  row.appendTableCell(quantitysum.toFixed(0));
  row.appendTableCell(productsum.toFixed(2));
  for (let i=2;i<5;i++){
    newtable.getRow(counter+1).getCell(i).getChild(0).asParagraph().setAttributes(bodyStyle);
  }

  return productsum;
}

function CreateReceipt(){
  let ui=DocumentApp.getUi();
  let body=DocumentApp.getActiveDocument().getBody();
  let dummytext="";

  //Gather Info
  let date=Date().slice(4,21).toUpperCase();
  let documentid=ui.prompt("What is the Receipt ID?").getResponseText();
  let client=ui.prompt("What is the Client's Name?").getResponseText();
  let amount=QuestionAndCheck(ui,"What is the Amount(RM) Received?",true);
  let purpose=ui.prompt("What is the Purpose of this Transaction?").getResponseText();
  let paymentmethod=ui.prompt("What is the Method of Payment?").getResponseText();

  let idlength=documentid.length;

  //Writing Body
  body.appendParagraph("                           PAYMENT RECEIPT                           ").setAttributes(boldtextstyle);
  body.appendParagraph("................................................................\n");
  body.appendParagraph(dummytext.padEnd(53-'0'-idlength)+"RECEIPT NO:"+documentid).setAttributes(textStyle);
  body.appendParagraph("DATE: "+date).setAttributes(textStyle);
  body.appendParagraph("\nFROM: "+client).setAttributes(textStyle);
  body.appendParagraph("\nAMOUNT (RM): "+amount.toFixed(2)).setAttributes(textStyle);
  body.appendParagraph("\nFOR: "+purpose).setAttributes(textStyle);
  body.appendParagraph("\nPAYMENT METHOD: "+paymentmethod).setAttributes(textStyle);
  body.appendParagraph("\n\n\n................			     	           .................\nCLIENT SIGNATURE					      COMPANY SIGNATURE").setAttributes(textStyle);

  let datalist=[documentid,client,date,amount,purpose];
  let transferrequest=ui.prompt("Would you like to transfer this data into your Receipt Spreadsheet? (Yes/No)").getResponseText().toLowerCase();
  if (transferrequest==="yes") TransferToReceiptSheet(datalist);
}

function TransferToInvoiceSheet(datalist){
  let ui=DocumentApp.getUi();
  let url=DocumentApp.getActiveDocument().getUrl();
  let documentid=datalist[0];
  let client=datalist[1];
  let date=datalist[2];
  let amount=datalist[3];

  //Getting Spreadsheet and Last Row
  let sheeturl=ui.prompt("What is the url of your Invoice Spreadsheet? (Create New File If None)").getResponseText();
  let sheet=SpreadsheetApp.openByUrl(sheeturl);
  let lastrow=sheet.getLastRow();
  let displayindex=(lastrow==0|| lastrow==1)? 0-'0':sheet.getRange("A"+lastrow).getDisplayValue();

  //Add new row if empty file
  if (lastrow==0){
    sheet.appendRow(["ID","INVOICE ID","CLIENT","DATE","AMOUNT(RM)","REFERENCE URL"]);
    for (let i=0;i<6;i++){
      sheet.setColumnWidth(i+1,invoicesheetcolumnwidth[i]);
    }

    //Customise header
    let header=sheet.getRange("A1:F1");
    header.setFontFamily("Times New Roman");
    header.setFontSize(14);
    lastrow++;
  }

  displayindex++;
  //Add row
  sheet.appendRow([displayindex,documentid,client,date,amount,url]);
  
  //Customise Body(Not Complete)
  let customiserow=("A"+(lastrow+1))+":"+("F"+(lastrow+1));
  let row=sheet.getRange(customiserow);
  row.setFontFamily("Times New Roman");
  row.setFontSize(13);
  row.setHorizontalAlignment("Left");

  //Style Table Borders
  let allcell=sheet.getDataRange();
  allcell.setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID);
}

function TransferToReceiptSheet(datalist){
  let ui=DocumentApp.getUi();
  let url=DocumentApp.getActiveDocument().getUrl();
  let documentid=datalist[0];
  let client=datalist[1];
  let date=datalist[2];
  let amount=datalist[3];
  let purpose=datalist[4];

  //Getting Spreadsheet and Last Row
  let sheeturl=ui.prompt("What is the url of your Receipt Spreadsheet? (Create New File If None)").getResponseText();
  let sheet=SpreadsheetApp.openByUrl(sheeturl);
  let lastrow=sheet.getLastRow();
  let displayindex=(lastrow==0|| lastrow==1)? 0-'0':sheet.getRange("A"+lastrow).getDisplayValue();

  //Add new row if empty file
  if (lastrow==0){
    sheet.appendRow(["ID","RECEIPT ID","CLIENT","DATE","AMOUNT(RM)","PURPOSE","REFERENCE URL"]);
    for (let i=0;i<7;i++){
      sheet.setColumnWidth(i+1,receiptsheetcolumnwidth[i]);
    }

    //Customise header
    let header=sheet.getRange("A1:G1");
    header.setFontFamily("Times New Roman");
    header.setFontSize(14);
    lastrow++;
  }

  displayindex++;
  //Add row
  sheet.appendRow([displayindex,documentid,client,date,amount,purpose,url]);
  
  //Customise Body(Not Complete)
  let customiserow=("A"+(lastrow+1))+":"+("G"+(lastrow+1));
  let row=sheet.getRange(customiserow);
  row.setFontFamily("Times New Roman");
  row.setFontSize(13);
  row.setHorizontalAlignment("Left");

  //Style Table Borders
  let allcell=sheet.getDataRange();
  allcell.setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID);
}

function QuestionAndCheck(ui,quesText,isFloat){
  let prompt=ui.prompt(quesText);
  let response=prompt.getResponseText();
  if (isNaN(response) || (response-'0')<=0){
    ui.alert("Please enter a valid number.");
    return QuestionAndCheck(ui,quesText,isFloat);
  }
  else{
    if (isFloat==false && response%1!=0){
      ui.alert("Please enter a valid whole number.");
      return QuestionAndCheck(ui,quesText,isFloat);
    }
    else{
      return response-'0';
    }  
  }
}