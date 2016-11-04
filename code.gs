/***************************************************************************
                       Variables & Settings
*****************************************************************************/

//Date object marking the start of the execution
var executionStart;
//Marked true if the time limit is eceeded so the script knows to re-run itself
var exceededMax = false;
//Elapsed time updated each time a folder is opened
var elapsedTime = 0;
//The maximum allowed execution time in seconds
var maxExecutionTime = 250;

//Array of drive files used when generating a CSV/sheet
var filesArray = [];

//First row headers for the spreadsheet/CSV
var spreadsheetColumnNames = [['Name', 'Type', 'Size', 'Path', 'URL', 'Created At','Last Updated At', 'Owner', 'Permissions', 'Editors Can Share', 'Viewers', 'Editors']];

var settings = {
  emailJson: true,
  emailCsv: true,
  saveJson: false,
  saveSheet: false,
  mapEntireDrive: true,
  getSharedWithMe: true,
  getTrashed: true
}

/***************************************************************************
                      Data Getting/Saving/Continuation
*****************************************************************************/

function StartDriveMap(){
  if(settings.mapEntireDrive) {
    MapDrive();
  }
}

//entry point for searching
function SearchForFiles(){
  var filesList = DriveApp.searchFiles(params)
}

//Entry point for mapping
function MapDrive() {
  
  var structure = GetPreviousState();
  var path = {};

  executionStart = new Date();
  try {    
    
    var parentFolder = DriveApp.getRootFolder();
    //var parentFolder = DriveApp.getFolderById('0B3ii56aSLRlsNGE3a2I2emppbm8');
    path = {path: parentFolder.getName()}
    
    TraverseDriveTree(parentFolder, structure, path);
    
    if(exceededMax){
      SaveCurrentState(structure);
      ScheduleContinuation(structure);
      EmailDriveMap(structure, true);      
    } else {
      if(settings.getSharedWithMe){
        structure['Shared With Me'] = GetSharedWithMeFiles();
      }
      
      if(settings.getTrashed){
        structure['Trash'] = GetTrashedFiles();
      }
      
      SendCompletionEmail(structure);
      ClearTriggers();
    }
  } catch (e) {
    Logger.log(e.toString());   
  }
}

//Gets a list of files that are shared with you
function GetSharedWithMeFiles(){
  var output = {};
  output = IterateFileList(DriveApp.searchFiles('sharedWithMe'), 'Shared With Me');
  return output;
}

//Gets a list of files that are in the trash
function GetTrashedFiles(){
  var output = {};
  output = IterateFileList(DriveApp.searchFiles('trashed=true'), 'Trash');
  return output;
}

function SendCompletionEmail(driveMap){
  var attachments= [];
  var body = 'Drive Map Attached.'
  
  if(settings.emailCsv){    
    var csvData = GenerateCSV()
    attachments.push(csvData.attachment);
    body += ' ' + csvData.url;
  }
  
  if(settings.emailJson){
    var jsonBlob = CreateJSONBlob(driveMap);
    attachments.push(jsonBlob);
  }
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Mapped Drive Data', body ,{ attachments: attachments})
}

//Emails the completed map to the executing users email address
//UNUSED?
function EmailDriveMap(driveMap, partial){
  var blob = CreateJSONBlob(driveMap);
  var subject = 'Drive Map Attached';
  if(partial) {
    subject = 'Partial Drive Map Attached';
  }
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Folder Tree', 'Drive Map Attached' ,{ attachments: [blob]})
}

//Converts the drive map data to JSON, and then to a blob to be attached to an email
function CreateJSONBlob(data){
  var jsonString = JSON.stringify(data, null, '\t');
  return Utilities.newBlob(jsonString, ContentService.MimeType.JSON, 'DriveMap.json');
  //return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
}

//Flattens the files array
function GenerateCSV(){
  var width = 12;
  if(filesArray.length > 0){
    var sheet = CreateGoogleSheet(width);
    var csv = convertRangeToCsv(sheet.sheet.getRange(1,1, filesArray.length + 1, width));
    var csvBlob = Utilities.newBlob(csv, ContentService.MimeType.CSV, 'DriveMap.csv');
    return {attachment: csvBlob, url: sheet.url};
  }
}

function CreateGoogleSheet(width){
  var sheet = SpreadsheetApp.create('Drive Map Output').insertSheet('Drive Map', 0);
  sheet.getRange(1, 1, 1, width).setValues(spreadsheetColumnNames);
  sheet.getRange(2, 1, filesArray.length, width).setValues(filesArray);  
  return {url: sheet.getParent().getUrl(), sheet: sheet};
}

//Saves the current data
function SaveCurrentState(driveMap){
  var saveObject = {
    map: driveMap,
    files: filesArray
  }
  var json = JSON.stringify(saveObject);
  var blob = Utilities.newBlob(json, ContentService.MimeType.JSON, 'DriveMapContinuation.json');
  var file = DriveApp.createFile(blob);
  var id = file.getId();
  PropertiesService.getScriptProperties().setProperty('continuationJson', id);
  Logger.log(id)
}

//Necessary to remove unused triggers
function ClearTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

//Necessary to retrieve the drive maps previous map if it was stopped part way
function GetPreviousState(){
  var property = PropertiesService.getScriptProperties().getProperty('continuationJson');
  if(property !== null){
    var file = DriveApp.getFileById(property);
    var blob = file.getAs('application/json');
    var blobString = blob.getDataAsString();
    var saveObject = JSON.parse(blobString);
    
    filesArray = saveObject.files;
    //Clear file and property
    DriveApp.removeFile(file);
    PropertiesService.getScriptProperties().deleteProperty('continuationJson');
    
    return saveObject.map;
  }
  return {};
}

function ScheduleContinuation(){
  var triggerAfter = 30*1000;
  ScriptApp.newTrigger('MapDrive').timeBased().after(triggerAfter).create();
}



/***************************************************************************
                       File & Folder Traversal
*****************************************************************************/

//Recursive breadth-first folder and file mapping
function TraverseDriveTree(parent, structure, path){
  if(CheckExecutionTime()){
    var childFolders = parent.getFolders();
    var parentName = parent.getName(); 
    
    //If there is no parent path add it
    if(typeof path[parentName] === 'undefined'){
      path[parentName] = {path: path.path};
    }
    
    //If parent does not exist in structure, then add it and it's files
    if(typeof structure[parentName] === 'undefined'){
      structure[parentName] = {};
      structure[parentName]['files'] = GetFilesInfo(parent.getFiles(), path[parentName].path);
    }  
    
    while (childFolders.hasNext()) {    
      var childFolder = childFolders.next();
      var childfolderName = childFolder.getName();
      
      //Only add the path if it does not already exist
      if(typeof path[parentName][childfolderName] === 'undefined'){
        path[parentName][childfolderName] = {path: path[parentName].path + " > " +childfolderName};        
      }
    
      if(typeof structure[parentName][childfolderName] === 'undefined'){
        structure[parentName][childfolderName] = {};
      }
      
      //Only get files if they do not already exist
      if(typeof structure[parentName][childfolderName]['files'] === 'undefined') {
        structure[parentName][childfolderName]["files"] = GetFilesInfo(childFolder.getFiles(), path[parentName][childfolderName].path);
      }
         
      // Recursive call for any sub-folders
      TraverseDriveTree(childFolder, structure[parentName], path[parentName]);    
    }         
  }
}

/***************************************************************************
                       File & Folder Searching
*****************************************************************************/

function IterateFileList(files, baseFolder){
  
  var output = {};
  var filesArray = [];
  while(files.hasNext()){
    var file = files.next();
    var parents = file.getParents();
    var path = GeneratePath(parents, baseFolder);
    filesArray.push(GetFileInfo(file, path));
  }
  output.files = filesArray;
  output.filesCount = filesArray.length;
  Logger.log(output)
  return output;
}

function GeneratePath(parents, baseFolder){
  var output = '';
  var parentsArray = [];
  while(parents.hasNext()){
    var parent = parents.next();
    parentsArray.push(parent.getName());
  }
  
  if(parentsArray.length == 0){
    return baseFolder;
  }
  
  for(var i = parentsArray.length - 1; i >= 0; i--){
    if(i == 0) {
      output += parentsArray[i];
      continue;
    }
    output += parentsArray[i] + ' > ';
  }
  return output;
}


/***************************************************************************
                       File Info Getting/Formatting
*****************************************************************************/

//Returns a formatted path from an array of folders
function FormatFilePath(names, isHTML){
  var output = '';
  if(!isHTML){
    for(var i = 0; i < names.length; i++){
      if(i == 0){
        output += names[i];
        continue;
      }
      output += '>' + names[i];
    }
  }
  return output;
}

//Iterates through a list of fiels and gets their info
function GetFilesInfo(files, path){
  var output = [];
  while(files.hasNext()){
    var file = files.next();
    output.push(GetFileInfo(file, path));
  }
  return output;
}

//Gets the info for a file
function GetFileInfo(file, path){
  var output = {};
  var owner = file.getOwner();
  output = {
      name: file.getName(),
      type: GetMIMEType(file),
      size: GetFileSize(file),
      path: path,
      url:  file.getUrl(),
      created: file.getDateCreated(),
      lastUpdated: file.getLastUpdated(),
      owner: {
        name: owner.getName(),
        email: owner.getEmail()
      },
      permissions: {
        accessPermissions: GetSharedAccessAndPermissions(file),
        editorsCanShare: file.isShareableByEditors(),      
        viewers: GetUsersInfo(file.getViewers()),
        editors: GetUsersInfo(file.getEditors())            
      }
    }
  if(settings.emailCsv || settings.saveSheet){
    filesArray.push(FlattenFileInfo(output));
  }
  return output;
}

//Flattens a file info object
function FlattenFileInfo(originalFileInfo){
  var fileInfo = JSON.parse(JSON.stringify(originalFileInfo));
  fileInfo.owner = fileInfo.owner.email;
  fileInfo.accessPermissions = fileInfo.permissions.accessPermissions;
  fileInfo.editorsCanShare = fileInfo.permissions.editorsCanShare;
  fileInfo.created = FormatDate(fileInfo.created, 'MM/dd/YYYY HH:MM:SS');
  fileInfo.lastUpdated = FormatDate(fileInfo.lastUpdated, 'MM/dd/YYYY HH:MM:SS');
  
  var viewers = ''
  if(fileInfo.permissions.viewers.length == 0){
    viewers = 'None'
  }
  for(var i = 0; i < fileInfo.permissions.viewers.length; i++){
    if(i == fileInfo.permissions.viewers.length - 1){
      viewers += fileInfo.permissions.viewers[i].email;
    } else{
      viewers += fileInfo.permissions.viewers[i].email + ', '
    }
  }
  
  fileInfo.viewers = viewers;
  
  var editors = ''
  if(fileInfo.permissions.editors.length == 0){
    editors = 'None'
  }  
  for(var i = 0; i < fileInfo.permissions.editors.length; i++){
    if(i == fileInfo.permissions.editors.length - 1){
      viewers += fileInfo.permissions.editors[i].email;
    } else{
      viewers += fileInfo.permissions.editors[i].email + ', '
    }
  }  
  
  fileInfo.editors = editors;
  delete fileInfo.permissions;
  
  var infoArray = [];
  
  for(var property in fileInfo){
    infoArray.push(fileInfo[property]);
  }
  
  return infoArray
}

//Returns a formatted MIME type
function GetMIMEType(file){
  var type = file.getMimeType();
  
  switch(type){
    case 'application/vnd.google-apps.audio':
      return 'Audio';
    case 'application/vnd.google-apps.document':
      return 'Google Document';
    case 'application/vnd.google-apps.drawing':
      return 'Google Drawing';
    case 'application/vnd.google-apps.file':
      return 'Google Drive file';
    case 'application/vnd.google-apps.folder':
      return 'Google Drive Folder';
    case 'application/vnd.google-apps.form':
      return 'Google Form';
    case 'application/vnd.google-apps.fusiontable':
      return 'Fusion Table';
    case 'application/vnd.google-apps.map':
      return 'Google Map';
    case 'application/vnd.google-apps.photo':
      return 'Image';
    case 'application/vnd.google-apps.presentation':
      return 'Google Slide';
    case 'application/vnd.google-apps.script':
      return 'Apps Scrips';
    case 'application/vnd.google-apps.sites':
      return 'Google Site';
    case 'application/vnd.google-apps.spreadsheet':
      return 'Google Spreadsheet';
    case 'application/vnd.google-apps.unknown':
      return 'Unknown';
    case 'application/vnd.google-apps.video':
      return 'Video';
    default:
      return type;      
  }
  
  /*
  switch(type){
    case 'GOOGLE_APPS_SCRIPT':
      return 'Google Apps Script';
    case 'GOOGLE_DRAWINGS':
      return 'Google Drawing';
    case 'GOOGLE_DOCS':
      return 'Google Document';
    case 'GOOGLE_FORMS':
      return 'Google Form';
    case 'GOOGLE_SHEETS':
      return 'Google Sheet';
    case 'GOOGLE_SLIDES':
      return 'Google Slide';
    case 'FOLDER':
      return 'Drive Folder';
    case 'BMP':
      return 'BMP Image';
    case 'GIF':
      return 'GIF Image';
    case 'JPEG':
      return 'JPEG Image';
    case 'PNG':
      return 'PNG Image';
    case 'SVG':
      return 'SVG Image';
    case 'PDF':
      return 'PDF';
    case 'CSS':
      return 'CSS';
    case 'CSV':
      return 'CSV';
    case 'JAVASCRIPT':
      return 'JavaScript';
    case 'PLAIN_TEXT':
      return 'Plain Text';
    case 'RTF':
      return 'Rich Text';
    case 'OPENDOCUMENT_GRAPHICS':
      return 'OpenDocument Graphic';
    case 'OPENDOCUMENT_PRESENTATION':
      return 'OpenDocument Presentation';
    case 'OPENDOCUMENT_SPREADSHEET':
      return 'OpenDocument Spreadsheet';
    case 'OPENDOCUMENT_TEXT':
      return 'OpenDocument Text';
    case 'MICROSOFT_EXCEL':
      return 'Microsoft Excel';
    case 'MICROSOFT_EXCEL_LEGACY':
      return 'Legacy Microsoft Excel';
    case 'MICROSOFT_POWERPOINT':
      return 'Microsoft PowerPoint';  
    case 'MICROSOFT_POWERPOINT_LEGACY':
      return 'Legacy Microsoft PowerPoint';
    case 'MICROSOFT_WORD':
      return 'Microsoft Word';
    case 'MICROSOFT_WORD_LEGACY':
      return 'Legacy Microsoft Word';
    case 'ZIP':
      return 'ZIP';
    default:
      return type;
  } */
}

//Gets and formats the file size
function GetFileSize(file){
  var bytes = file.getSize();
  
  if(bytes <= 1023){
    return bytes + " Bytes"
  } else if(bytes <= 1048575) {
    return Math.floor((bytes/1024)*10)/10 + " KB";
  } else if(bytes <= 1073741823) {
    return Math.floor((bytes/1048576)*10)/10 + " MB";
  } else {
    return Math.floor((bytes/1073741824)*10)/10 + " GB";
  }
}

//Gets and formats the current sharing and access permissions for the file
function GetSharedAccessAndPermissions(file){
  var access;
  var permissions;
  
  //Necessary as some files can have conflicting sharing settings
  try {
    access = DriveApp.Access[file.getSharingAccess()].name();
    permissions = DriveApp.Permission[file.getSharingPermission()].name();
  } catch(e) {
    Logger.log("Invalid Sharing Settings: "+ file.getName() + ": " + file.getUrl());
    return 'Invalid Sharing Settings';
  }

  
  var output = '';
  switch(access){
    case 'ANYONE':
      output += 'Anyone can find and ';
      break;
    case 'ANYONE_WITH_LINK':
      output += 'Anyone with a link can ';
      break;
    case 'DOMAIN':
      output += 'Anyone within your domain can find and ';
      break;
    case 'DOMAIN_WITH_LINK':
      output += 'Anyone within your domain with a link can ';
      break;
    case 'PRIVATE':
      return 'Private';
      break;      
  }
  
  switch(permissions){
    case 'VIEW':
      output += 'view';
      break;
    case 'EDIT':
      output += 'edit';
      break;
    case 'COMMENT':
      output += 'comment';
      break;     
  }
  
  return output;
}

//Checks if the execution time has reached the limit
function CheckExecutionTime(){
  if(elapsedTime >= maxExecutionTime){
    return false;
  }
  
  elapsedTime = (new Date().getTime() - executionStart.getTime())/1000;
  if(elapsedTime >= maxExecutionTime){
    exceededMax = true;
    return false;
  }
  return true;
}

//Gets info for an aray of users
function GetUsersInfo(users){
  var output = [];
  for(var  i = 0; i < users.length; i++){
    output.push({
      name: users[i].getName(),
      email: users[i].getEmail()
    })
  }
  return output;
}

function convertRangeToCsv(range) {
  try {
    var data = range.getValues();
    var csvFile = undefined;

    // Loop through the data in the range and build a string with the CSV data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

//Reformts a date
function FormatDate(date, format)
{
  var temp = new Date(date);
  var output = Utilities.formatDate(temp, "PST", format);
  return output;
}
