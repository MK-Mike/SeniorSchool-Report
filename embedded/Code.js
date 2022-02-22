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

/** Funtion to process the SpreadSheet data*/
function processSpreadSheet() {
  
  let year = infoTab.getRange(yearLoc).getValue();
  let term = infoTab.getRange(termLoc).getValue()

  // check if folder inside "reports folder" for the year, if not create folder for year
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
  
  //iterate over the grades present 
  for (let g of grades){

    // get the atributes of the template so they can be applied to the new docs
    let style = template.getBody().getAttributes()

    //create a new file for the grade with name string "YYYY-Term_X-Grade_Y-REPORTS"
    const getLast = (str) => {return str.slice(str.length -1)}
    let nameString = `${year}-Term_${getLast(term)}-Grade_${g}-REPORTS`
    console.log(nameString)
  
    // check if the file already exists, if not create the files 
    if (!reportsFolder.getFoldersByName(year).next().getFilesByName(nameString).hasNext()){
      
      //create the file and store it to a variable
      var gradeDoc = DocumentApp.create(nameString)
      // apply the styling 
      gradeDoc.getBody().setAttributes(style)

      //move the file into the correct folder and let the user know
      var gradeFileID = DriveApp.getFileById(gradeDoc.getId())
      let yearfolder = DriveApp.getFoldersByName(year).next()

      yearfolder.addFile(gradeFileID)
      t(`File created for Grade ${g} reports: ${nameString}`) //t for toast, see function at end
    }

    // filter data to match grade 
    let filterdEntries = entries.filter((array)=> {if(array[1]===g) return array })

    // set a state variable to enable the deletion of the empty paragraph created
    var firstEntry = true;
    
    //iterate over each row, create the report and then append it to the grade doc
    for (let ent of filterdEntries){

      let docid = DriveApp.getFilesByName(templateFileName).next().makeCopy().setName('temp').getId()
      let targetFileID = DriveApp.getFilesByName(nameString).next().getId()

      // create an object for each learner by destructuring the row 
      let learnerObject = new Entry(...ent)
      
      t(`Creating report for ${learnerObject.eName}`)

      enterDetails(docid,learnerObject)
      mergeGoogleDocs(docid,targetFileID,firstEntry)

      // delete the temp file and set state variable to false so you don't keep deleting lines
      DriveApp.getFileById(docid).setTrashed(true)
      firstEntry = false
    }

    // inform the user after finished with reports for the grade 
    t(`Finished creating Grade ${g} reports`)
  }

  //inform the user of success and give them a link to the folder
  let url = reportsFolder.getFoldersByName(year).next().getUrl()
  modalDialogue(url)
}

/**function to enter the details using info saved in the learner object  */
function enterDetails(id,learnerEntry) {

  //keep track of entries in case of error
  console.log(learnerEntry.eName)
  let doc = DocumentApp.openById(id)
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

/** t for toast */
let t = (msg) => { SpreadsheetApp.getActiveSpreadsheet().toast(msg)}

/** modalDialogue function */
let modalDialogue = (url) => {

  let msgTemplate = `Your reports can be found <a href=${url} target="_blank"> here </a> <br><br> You will need to set the orientation to landscape yourself, sorry :( but I hope this saved you some time <br><br><input type="button" value="Close" onclick="google.script.host.close()" />`

  //create HTML output from the template
  let html = HtmlService.createHtmlOutput(msgTemplate).setWidth(400).setHeight(200)

  //show the pop-up
  SpreadsheetApp.getUi().showModalDialog(html, 'Success!')
}

/** function to create a new object to hold the values for each learner. e for entry */
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
 
//more inspiration from https://gist.github.com/mhawksey/1170597/





