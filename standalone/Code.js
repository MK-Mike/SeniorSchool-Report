//let templateURL = "https://docs.google.com/document/d/1NKzo5CNhfWyeTXe-Fn0jXrK3czr0Nmxv1z_IL-F0wrw/edit"
let templateURL = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1bNV2m6IwblOJ7gyxfOSWykXU8AaRmCGn-uSD3XJaFdM/edit#gid=0').getSheetByName('Sheet1').getRange('H1').getValue()
let template = DocumentApp.openByUrl(templateURL)
let sheetURL = `https://docs.google.com/spreadsheets/d/1bNV2m6IwblOJ7gyxfOSWykXU8AaRmCGn-uSD3XJaFdM/edit#gid=0`;
let infoTabName = 'Sheet1';
let ss = SpreadsheetApp.openByUrl(sheetURL);
let infoTab = ss.getSheetByName(infoTabName);
let entryRange = infoTab.getRange(`A1`).getValue();
let entries = ss.getRange(entryRange).getValues();
let yearLoc = `D1`
let termLoc = `F1`
let reportsFolder = DriveApp.getFoldersByName('Mikes Testing Area').next()
let templateFileName = `Davids Report Card Template Dreams`;


/** Search keys to replace in the template */
let name = '{Name}'
attendedSG = '{X}'
totalSG = '{Y}'
attendedEK = '{Q}'
totalEK = '{Z}'
assessPS = '{A}'
assessPSTot = '{B}'
markPS = '{C}'
gradeAvPS = '{D}'
assessMath = '{L}'
assessMathTot = '{M}'
markMath = '{N}'
gradeAvMath = '{O}'
assessEng = '{P}'
assessEngTot = '{Q}'
markEng = '{R}'
gradeAvEng = '{S}';
pmComments = '{comments}'



function enterDetails(id,learnerEntry) {
  console.log(learnerEntry.eName)
  let doc = DocumentApp.openById(id)
  //doc.setName(learnerEntry.eName)
  let body = doc.getBody()

  body.replaceText(name, learnerEntry.eName)
  body.replaceText(attendedSG, learnerEntry.eattendedSG)
  body.replaceText(totalSG, learnerEntry.etotalSG)
  body.replaceText(attendedEK, learnerEntry.eattendedEK)
  body.replaceText(totalEK, learnerEntry.etotalEK)
  body.replaceText(assessPS, learnerEntry.eassessPS)
  body.replaceText(assessPSTot, learnerEntry.eassessPSTot)
  body.replaceText(markPS, learnerEntry.emarkPS)
  body.replaceText(gradeAvPS, learnerEntry.egradeAvPS)
  body.replaceText(assessMath, learnerEntry.eassessMath)
  body.replaceText(assessMathTot, learnerEntry.eassessMathTot)
  body.replaceText(markMath, learnerEntry.emarkMath)
  body.replaceText(gradeAvMath, learnerEntry.egradeAvMath)
  body.replaceText(assessEng, learnerEntry.eassessEng)
  body.replaceText(assessEngTot, learnerEntry.eassessEngTot)
  body.replaceText(markEng, learnerEntry.emarkEng)
  body.replaceText(gradeAvEng, learnerEntry.egradeAvEng)
  body.replaceText(pmComments, learnerEntry.ecomments)

}

function Entry(
  eName,
  eGrade,
  eSchool,
  eattendedSG,
  etotalSG,
  eattendedEK,
  etotalEK,
  eassessPS,
  eassessPSTot,
  emarkPS,
  egradeAvPS,
  eassessMath,
  eassessMathTot,
  emarkMath,
  egradeAvMath,
  eassessEng,
  eassessEngTot,
  emarkEng,
  egradeAvEng,
  ecomments) {
  this.eName = eName,
    this.eattendedSG = eattendedSG,
    this.eGrade = eGrade,
    this.eSchool = eSchool,
    this.etotalSG = etotalSG,
    this.eattendedEK = eattendedEK,
    this.etotalEK = etotalEK,
    this.eassessPS = eassessPS,
    this.eassessPSTot = eassessPSTot,
    this.emarkPS = emarkPS,
    this.egradeAvPS = egradeAvPS,
    this.eassessMath = eassessMath,
    this.eassessMathTot = eassessMathTot,
    this.emarkMath = emarkMath,
    this.egradeAvMath = egradeAvMath,
    this.eassessEng = eassessEng,
    this.eassessEngTot = eassessEngTot,
    this.emarkEng = emarkEng,
    this.egradeAvEng = egradeAvEng,
    this.ecomments = ecomments
}


function processSpreadSheet() {
  
  let year = infoTab.getRange(yearLoc).getValue();
  let term = infoTab.getRange(termLoc).getValue()

  // check if folder for year, if not create folder for year
  if (!reportsFolder.getFoldersByName(year).hasNext()) {
    reportsFolder.createFolder(year)   
  }

  // get list of grades 

  let getGrades = (arr)=>{
    let newArr = []
    for (let r of arr) newArr.push(r[1])
    return [...new Set(newArr)]
  }
  let grades = getGrades(entries)
  
  for (let g of grades){

    //create a new file for each grade 
    // name string YYYY-Term_X-Grade_Y-REPORTS
    const getLast = (str) => {return str.slice(str.length -1)}
    let nameString = `${year}-Term_${getLast(term)}-Grade_${g}-REPORTS`
    console.log(nameString)

    let style = template.getBody().getAttributes()
    
    if (!reportsFolder.getFoldersByName(year).next().getFilesByName(nameString).hasNext()){
      var gradeDoc = DocumentApp.create(nameString)
      gradeDoc.getBody().setAttributes(style)
      var gradeFileID = DriveApp.getFileById(gradeDoc.getId())
      let yearfolder = DriveApp.getFoldersByName(year).next()
      yearfolder.addFile(gradeFileID)

    }
    // filter data to match grade 
    let filterdEntries = entries.filter((array)=> {if(array[1]===g) return array })
    var firstEntry = true;
    //for each row 
    for (let ent of filterdEntries){
      let docid = DriveApp.getFilesByName(templateFileName).next().makeCopy().setName('temp').getId()
      let learnerObject = new Entry(...ent)
      let targetFileID = DriveApp.getFilesByName(nameString).next().getId()
      
      enterDetails(docid,learnerObject)
      mergeGoogleDocs(docid,targetFileID,firstEntry)
      DriveApp.getFileById(docid).setTrashed(true)
      firstEntry = false
    }
  }
}

// adapted from https://www.labnol.org/code/19892-merge-multiple-google-documents
function mergeGoogleDocs(docID,baseDocID,firstEntry) {

  var baseDoc = DocumentApp.openById(baseDocID)
  var body = baseDoc.getActiveSection();
  var otherBody = DocumentApp.openById(docID).getActiveSection();  
  var totalElements = otherBody.getNumChildren();
  

  for( var j = 0; j < totalElements; ++j ) {
    // var ogElement = otherBody.getChild(j)
    var element = otherBody.getChild(j).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH ){
      body.appendParagraph(element);
      if(firstEntry) body.getChild(0).removeFromParent();
      firstEntry = false;
    }
    else if( type == DocumentApp.ElementType.TABLE ){
      //f(body,ogElement,element)
      body.appendTable(element);
    }
    else if( type == DocumentApp.ElementType.LIST_ITEM )
      body.appendListItem(element);
    else
      throw new Error("Unknown element type: "+type);
  } 
  //body.appendParagraph('').appendPageBreak()
}

let g = ()=>{console.log(template.getBody().getAttributes())}


 
//https://gist.github.com/mhawksey/1170597/





