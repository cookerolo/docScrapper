function getJBContent(id) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var doc = DocumentApp.openById(id);
  var body = doc.getBody();
  var table= body.getTables()[0];
  try {
    var pipelineTitle = doc.getName().split(" ",-1);
    var position = pipelineTitle.length;
    var pipelineId = (pipelineTitle[position-1]*1).toFixed(0);    
    var jobTitles = table.getRow(0).getText().split("JobBoard:")[1];
    var jobTitle = jobTitles.substr(0,jobTitles.lastIndexOf("Linked")).trim();
    var keywords = table.getRow(1).getText().split("(Prioritized)")[1].trim();
    var jobDescription = table.getRow(3).getText().split("Job Description",-1)[1].trim();
  }
  catch(err) {
    var pipelineId = "";    
    var jobTitle = "";
    var keywords = "";
    var jobDescription = "";    
    }
  results = []
  results.push(pipelineId, jobTitle, keywords, jobDescription);
  Logger.log(results);
  return results;
}

function listFilesInFolder(id) {
  var folder = DriveApp.getFolderById(id);
  var contents = folder.getFiles();
  var file;
  var name;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var date;
  var size;
  sheet.clear();
  output = [];
  output.push(["PipelineId","Name", "Last Updated", "Id"]);
  while(contents.hasNext()) {
    file = contents.next();
    name = file.getName();
    var pipelineTitle = name.split(" ",-1);
    var position = pipelineTitle.length;
    var pipelineId = (pipelineTitle[position-1]*1).toFixed(0);    
    date = file.getLastUpdated()
    docId = file.getId()

    data = [pipelineId, name, date, docId]
    output.push(data);
  }
  sheet.getRange(1,1,output.length,4).setValues(output)
  sheet.sort(3, false);
  sheet.getRange(1,7).setValue("Last Updated (GMT-3):")
  sheet.getRange(1,8).setValue(new Date())
};

function scrapCps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var sh2 = ss.getSheets()[1];
  sh2.clear();
  var lastRow = sh.getRange("D:D").getValues().filter(String).length;
  var valuesOnSheet = sh.getRange(2,4,lastRow,1).getValues();
  var output =[['Pipeline Id', 'Job Title', 'Keywords', 'Job Description']];
  for(var i=0; i < valuesOnSheet.length-1; i++){
    var result = getJBContent(valuesOnSheet[i]);
    output.push(result);
  }
  sh2.getRange(1,1,output.length,4).setValues(output);
  sh2.getRange("A:D").setVerticalAlignment('top');
}
