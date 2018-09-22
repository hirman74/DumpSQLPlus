//include moment JS

//Base64.js
//var Base64 = ???


var zipperExe = "zip.exe"; //put to path environment

var shellapp = new ActiveXObject("Shell.Application");
var objShell = new ActiveXObject("WScript.shell");
var comspec = objShell.ExpandEnvironmentStrings("%comspec%"); //32bit or 64bit
var objFSO = new ActiveXObject("Scripting.FileSystemObject"),
    ForWriting = 2,
    ForReading = 1,
    ForAppending = 8,
    CreateIt = true,
    dontWantCreateIt = false,
    AsciiMode = 0,
    UnicodeMode = -1,
    systemDefaultMode = -2;

function js_yyyymmdd_hhmmss() {
    now = new Date();
    year = "" + now.getFullYear();
    month = "" + (now.getMonth() + 1);
    if (month.length == 1) {
        month = "0" + month;
    }
    day = "" + now.getDate();
    if (day.length == 1) {
        day = "0" + day;
    }
    hour = "" + now.getHours();
    if (hour.length == 1) {
        hour = "0" + hour;
    }
    minute = "" + now.getMinutes();
    if (minute.length == 1) {
        minute = "0" + minute;
    }
    second = "" + now.getSeconds();
    if (second.length == 1) {
        second = "0" + second;
    }
    return year + month + day + "_" + hour + minute + second;
}

var grabDateTime = js_yyyymmdd_hhmmss();

var shellRunOption = {
    hideWindow: function () {
        return 0
    },
    showWindow: function () {
        return 1
    },
    minimize: function () {
        return 2
    },
    maximize: function () {
        return 3
    },
    restore: function () {
        return 4
    },
    activeRestore: function () {
        return 5
    },
    minimizeZorder: function () {
        return 6
    }
} //different types of MS-DOS windows

var customFileFolder = {
    deleteFile: function (filePath) {
        if (objFSO.FileExists(filePath)) {
            var afile = objFSO.GetFile(filePath);
            afile.Attributes[0];
            afile.Delete();
        }
    },
    copyFile: function (FromFile, ToFile, overwrite) {
        // Set overwrite to true or false; FromFile, etc = full paths
        var f = objFSO.GetFile(FromFile);
        f.Attributes[0];
        f.Copy(ToFile, overwrite);
    },
    makeFolder: function (DesiredPath) {
        var f = objFSO.CreateFolder(DesiredPath); // 'DesiredFolderPath' : e.g., "C:\\MainFolder\\NewFolderName".
    },
    deleteFolder: function (DesiredPath) { // where gpath = full folder path
        if (objFSO.FolderExists(DesiredPath)) {
            var afolder = objFSO.GetFolder(DesiredPath);
			afolder.Attributes[0];
            afolder.Delete();
        }
    },
    copyFolder: function (FromFolder, ToFolder, overwrite) {
        // where FromFolder, etc = full folder paths
        var f = objFSO.GetFolder(FromFolder);
        f.Copy(ToFolder, overwrite);
    },
    checkMakeFolder: function (NewFolderName) {
        if (objFSO.FolderExists(NewFolderName)) {} else {
            var afolder = objFSO.CreateFolder(NewFolderName);
        }
    }
} //different types of file folder methods


var toWriteLog = ''; 

if (WScript.Arguments.length == 13) {
    var userName = WScript.Arguments.Item(0);
    var passWord = WScript.Arguments.Item(1);
    var dataBase = WScript.Arguments.Item(2);
    var splitFile = WScript.Arguments.Item(3); //duration of each split file in minutes
    var endDated = WScript.Arguments.Item(4); //"DD/MM/YYYY HH:mm:ss" // 1 = now
    var pastDays = WScript.Arguments.Item(5); //get start time via script using DAY count
    var missingFiles = WScript.Arguments.Item(6); //1 = find only missing file // 0 = redownload all files
    var resourceUpdates = WScript.Arguments.Item(7); //seconds intervals to update into log file or WScript.echo
    var showHideLauncher = WScript.Arguments.Item(8); //2 = taskbar 1 = show echo // 0 = hide fully
    var objStartFolder = WScript.Arguments.Item(9); //since run by Task Scheduler, need to write full directory name
    var scriptFolder = WScript.Arguments.Item(10); //since run by Task Scheduler, need to write full directory name
    var showEcho = WScript.Arguments.Item(11); //showEcho // noEcho
    var deleteDays = WScript.Arguments.Item(12); //delete how many days before

    customFileFolder.checkMakeFolder(objStartFolder);
    var stringElements = [];
    stringElements.push(userName);
    stringElements.push(passWord);
    stringElements.push(dataBase);
    stringElements.push(splitFile);
    stringElements.push(endDated); //"DD/MM/YYYY HH:mm:ss" // 1 = now
    stringElements.push(pastDays);
    stringElements.push(missingFiles);
    stringElements.push(resourceUpdates);
    stringElements.push(showHideLauncher);
    stringElements.push(objStartFolder);
    stringElements.push(scriptFolder);
    stringElements.push(showEcho);
    stringElements.push(deleteDays);

    toWriteLog = stringElements.join("\r\n");
    createLogFile(objStartFolder, toWriteLog);
    splitTimeDuration(stringElements);

} else {
    var toWriteLog = '===================================\r\n';
    toWriteLog += js_yyyymmdd_hhmmss() + 'Input_data_incomplete' + '\r\n';
    toWriteLog += '===================================';
    createLogFile(objStartFolder, toWriteLog);
    if (showEcho == 'showEcho') {
        WScript.echo(toWriteLog);
    }
} //checking if arguments collected is enough


function splitTimeDuration(stringElements) { //collect al the arguments and checking the start and end time
    //topResultDIV (stringElements);
    var userName = stringElements[0];
    var passWord = stringElements[1];
    var dataBase = stringElements[2];
    var splitSession = stringElements[3]; //duration of each split file in MINUTES
    var endDated = stringElements[4]; //"DD/MM/YYYY HH:mm:ss" // 1 = now
    var pastDays = stringElements[5]; //get start time via script using DAY count
    var missingFiles = stringElements[6]; //1 = find only missing file // 0 = redownload all files
    var resourceUpdates = stringElements[7]; //seconds intervals to update into log file or WScript.echo
    var showHideLauncher = stringElements[8]; //2 = taskbar 1 = show echo // 0 = hide fully
    var objStartFolder = objFSO.GetFolder(stringElements[9]); //since run by Task Scheduler, need to write full directory name
    var deleteDays = stringElements[12]; //delete how many days before

    customFileFolder.checkMakeFolder(objStartFolder);
    var folderStore = objStartFolder + '\\folderStore';
    customFileFolder.checkMakeFolder(folderStore);
    deletePrematureFiles(objStartFolder);


    var tempFolder = objStartFolder + '\\' + grabDateTime + '_tempFolder';
    //var tempFolder = objStartFolder + '\\tempFolder';
    customFileFolder.checkMakeFolder(tempFolder);

    deleteOlderZip(folderStore, deleteDays);
    deleteRogueFiles(objStartFolder);

    var endDateTimeFormat, endDateTimeActual, startDateTimeFormat, startDateTimeActual, countDays;
    //checking end date time value
    if (endDated == 'now') {
        //(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss'))
        //moment(startDateTime).format('DD/MM/YYYY HH:mm:ss'));
        var start = new Date();
        var dateNow = '';
        dateNow = start.getDate();
        dateNow = dateNow + '';
        if (dateNow.length == 1) {
            dateNow = '0' + dateNow
        }
        var monthNow = 0;
        monthNow = start.getMonth() + 1;
        monthNow = monthNow + '';
        if (monthNow.length == 1) {
            monthNow = '0' + monthNow
        }
        var fullYearNow = start.getFullYear();
        var dateString = dateNow + '/' + monthNow + '/' + fullYearNow + ' 00:00:00';
        var endDateTimeActual = moment(dateString, "DD/MM/YYYY HH:mm:ss"); //object
        var endDateTimeFormat = moment(endDateTimeActual).format('DD/MM/YYYY HH:mm:ss'); //string
        /*
        endDated = moment(); //selecting current date time as moment object
        endDateTimeFormat = moment().format('DD/MM/YYYY');
        endDateTimeActual = moment(endDated, "DD/MM/YYYY");
        */
        //} else if (moment(endDated, "DD/MM/YYYY HH:mm:ss").isValid()) { //converting collected string into moment object and check if its valid
    } else if (moment(endDated, "DD/MM/YYYY HH:mm:ss").isValid()) { //converting collected string into moment object and check if its valid
        endDateTimeActual = moment(endDated, "DD/MM/YYYY HH:mm:ss");
        endDateTimeFormat = moment(endDateTimeActual).format('DD/MM/YYYY HH:mm:ss');
    } else {
        var start = new Date();
        var dateNow = '';
        dateNow = start.getDate();
        dateNow = dateNow + '';
        if (dateNow.length == 1) {
            dateNow = '0' + dateNow
        }
        var monthNow = 0;
        monthNow = start.getMonth() + 1;
        monthNow = monthNow + '';
        if (monthNow.length == 1) {
            monthNow = '0' + monthNow
        }
        var fullYearNow = start.getFullYear();
        var dateString = dateNow + '/' + monthNow + '/' + fullYearNow + ' 00:00:00';
        var endDateTimeActual = moment(dateString, "DD/MM/YYYY HH:mm:ss");
        var endDateTimeFormat = moment(endDateTimeActual).format('DD/MM/YYYY HH:mm:ss');
    }

    //checking few days before end date as start date
    if (parseInt(pastDays) > 0) { //find out the start date time based on end time and length of days to collect
        countDays = parseInt(pastDays);
        if (moment(endDateTimeActual).isValid()) {
            var startTemp = moment(endDateTimeActual).subtract(countDays, 'days'); //object
            startTemp = moment(startTemp).format('DD/MM/YYYY HH:mm:ss').split(' ')[0]; //string
            startDateTimeActual = moment(startTemp + ' 00:00:00', "DD/MM/YYYY HH:mm:ss"); //object
            startDateTimeFormat = moment(startDateTimeActual).format('DD/MM/YYYY HH:mm:ss'); //string
        }
    }

    if (parseInt(pastDays) == 0) {
        if (moment(endDateTimeActual).isValid()) {
            var datePart = endDateTimeFormat.split(' ');
            startDateTimeActual = moment(datePart[0] + ' 00:00:00', "DD/MM/YYYY HH:mm:ss"); //object //need to choose yesterday 00:00:00
            startDateTimeFormat = moment(startDateTimeActual).format('DD/MM/YYYY HH:mm:ss'); //string
        }
    }

    //Wed Feb 14 12:57:20 UTC +0800 2018
    //listArrayModify.push (moment(colFiles.item().DateLastModified).isValid());
    //listArrayModify.push (moment(colFiles.item().DateLastModified, "DD/MM/YYYY HH:mm:ss"));
    //listArrayModify.push (colFiles.item().DateLastModified);
    //  var missingFiles = WScript.Arguments.Item(6); //1 = find only missing file // 0 = redownload all files

    switch (missingFiles) {
        case '1':
            missingFiles = true; //find only missing file 
            break;
        case '0':
            missingFiles = false; //redownload all files
            break;
        default:
            missingFiles = true; //find only missing file 
    }

    var timeSlots = zipFileExist(folderStore, startDateTimeActual, endDateTimeActual); //collecting all the time gaps which need to download file and then zip up 
    var connectTest = sqlTestConnect(stringElements, grabDateTime);
    if ((timeSlots.length == 0) && missingFiles && connectTest) {
        var toWriteLog = js_yyyymmdd_hhmmss() + ' = Already found zip file within time range of ' + startDateTimeFormat + ' to ' + endDateTimeFormat + '.\r\n';
        toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = Please move or delete those zip files containing the datetime range. Or choose a different datetime range to download.' + '\r\n';
        toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = This script will exit.'
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }
    } else if ((timeSlots.length > 0) && missingFiles && connectTest) {
        for (var sT = 0; sT < timeSlots.length; sT++) {
            var newStartDateTime = timeSlots[sT][0]; //moment object
            var newEndDateTime = timeSlots[sT][1]; //moment object
            var minuteDiff = newEndDateTime.diff(newStartDateTime, 'minutes');
            if (minuteDiff > 0) {
                //arrayShowFileState (minuteDiff + " " + parseInt(minuteDiff) + " " + splitSession);
                if ((parseInt(minuteDiff) / parseInt(splitSession)) > 1) {
                    //WScript.echo ('spoolTimeCalculatePositive' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + splitSession + ' = ' + tempFolder);
                    spoolTimeCalculatePositive(stringElements, newStartDateTime, newEndDateTime, splitSession, tempFolder, folderStore);
                } else if ((parseInt(minuteDiff) / parseInt(splitSession)) == 1) {
                    //WScript.echo ('spoolFileMaker0' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                    spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
                } else if ((parseInt(minuteDiff) / parseInt(splitSession)) < 1) {
                    //WScript.echo ('spoolFileMaker1' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                    spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
                }
            } else if (minuteDiff < 0) {
                if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) > 1) {
                    //WScript.echo ('spoolTimeCalculateNegative' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + splitSession + ' = ' + tempFolder);
                    spoolTimeCalculateNegative(stringElements, newStartDateTime, newEndDateTime, splitSession, tempFolder, folderStore);
                } else if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) == 1) {
                    //WScript.echo ('spoolFileMaker2' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                    spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
                } else if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) < 1) {
                    //WScript.echo ('spoolFileMaker3' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                    spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
                }
            } else if (minuteDiff == 0) {
                //WScript.echo ('spoolFileMaker4' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            }

        }
    } else if (!missingFiles && connectTest) {
        var newStartDateTime = startDateTimeActual; //moment object
        var newEndDateTime = endDateTimeActual; //moment object
        var minuteDiff = newEndDateTime.diff(newStartDateTime, 'minutes');
        if (minuteDiff > 0) {
            //arrayShowFileState (minuteDiff + " " + parseInt(minuteDiff) + " " + splitSession);
            if ((parseInt(minuteDiff) / parseInt(splitSession)) > 1) {
                //WScript.echo ('spoolTimeCalculatePositive' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + splitSession + ' = ' + tempFolder);
                spoolTimeCalculatePositive(stringElements, newStartDateTime, newEndDateTime, splitSession, tempFolder, folderStore);
            } else if ((parseInt(minuteDiff) / parseInt(splitSession)) == 1) {
                //WScript.echo ('spoolFileMaker0' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            } else if ((parseInt(minuteDiff) / parseInt(splitSession)) < 1) {
                //WScript.echo ('spoolFileMaker1' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            }
        } else if (minuteDiff < 0) {
            if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) > 1) {
                //WScript.echo ('spoolTimeCalculateNegative' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + splitSession + ' = ' + tempFolder);
                spoolTimeCalculateNegative(stringElements, newStartDateTime, newEndDateTime, splitSession, tempFolder, folderStore);
            } else if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) == 1) {
                //WScript.echo ('spoolFileMaker2' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            } else if ((parseInt(Math.abs(minuteDiff)) / parseInt(splitSession)) < 1) {
                //WScript.echo ('spoolFileMaker3' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
                spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            }
        } else if (minuteDiff == 0) {
            //WScript.echo ('spoolFileMaker4' + ' = ' + startDateTime + ' = ' + endDateTime + ' = ' + tempFolder);
            spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
        }
    } else if (!connectTest) {
        var toWriteLog = '===================================\r\n';
        toWriteLog += js_yyyymmdd_hhmmss() + '\r\n';
        toWriteLog += 'SQL connect test FAILED\r\n';
        toWriteLog += 'Exit Code\r\n';
        toWriteLog += '===================================';
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }
    }

    if (objFSO.FolderExists(tempFolder)) {
        var availFiles = [];
        var objFolder = objFSO.GetFolder(tempFolder);
        var colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            availFiles.push(colFiles.item().name);
        }
        if (availFiles.length == 0) {
			WScript.echo("Line : 2021");
            customFileFolder.deleteFolder(tempFolder);
        }
    }
}

function spoolTimeCalculatePositive(stringElements, startDateTime, endDateTime, splitSession, tempFolder, folderStore) { //for end time in future
    var newEndDateTime = startDateTime;
    var newStartDateTime = startDateTime;
    do {
        newEndDateTime = moment(newStartDateTime).add(splitSession, 'minute');
        if (moment(newEndDateTime) > moment(endDateTime)) {
            newEndDateTime = moment(endDateTime);
            spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            break;
        }
        spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
        newStartDateTime = newEndDateTime;
    } while (moment(newEndDateTime) < moment(endDateTime));
}

function spoolTimeCalculateNegative(stringElements, startDateTime, endDateTime, splitSession, tempFolder, folderStore) { //for end time value in reverse
    var newEndDateTime = startDateTime;
    var newStartDateTime = startDateTime;
    do {
        newEndDateTime = moment(newStartDateTime).subtract(splitSession, 'minute');
        if (moment(newEndDateTime) < moment(endDateTime)) {
            newEndDateTime = moment(endDateTime);
            spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
            break;
        }
        spoolFileMaker(stringElements, newStartDateTime, newEndDateTime, tempFolder, folderStore);
        newStartDateTime = newEndDateTime;
    } while (moment(newEndDateTime) > moment(endDateTime));
}

function spoolFileMaker(stringElements, startDateTime, endDateTime, tempFolder, folderStore) {
    var fileFolderName = tempFolder + '\\' + dateReformatted(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss')) + '_to_' + dateReformatted(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss'));
    var dateName = dateReformatted(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss')) + '_to_' + dateReformatted(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss')); //convert object to string

    var sqlFile = objFSO.OpentextFile(fileFolderName + '.sql', ForWriting, CreateIt, systemDefaultMode);
    sqlFile.writeline('set colsep ,' + '\r\n' + 'set headsep off' + '\r\n' + 'set pagesize 0' + '\r\n' + 'set trimspool off' + '\r\n' + 'set linesize 10000' + '\r\n' + 'set numw 50' + '\r\n' + 'set echo on');
    sqlFile.writeline("spool '" + fileFolderName + ".txt'");

    sqlFile.write("select count(*) FROM EV_COMBINED WHERE CREATETIME BETWEEN TO_DATE('");
    sqlFile.write(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss'));
    sqlFile.write("','DD/MM/YYYY HH24:MI:SS') AND TO_DATE('");
    sqlFile.write(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss'));
    sqlFile.writeline("','DD/MM/YYYY HH24:MI:SS');");

    sqlFile.write("SELECT SOURCE_TABLE," + "\r\n" + "PKEY," + "\r\n" + "SUBSYSTEM_KEY," + "\r\n" + "PHYSICAL_SUBSYSTEM_KEY," + "\r\n" + "LOCATION_KEY," + "\r\n" + "SEVERITY_KEY," + "\r\n" + "EVENT_TYPE_KEY," + "\r\n" + "ALARM_ID," + "\r\n" + "ALARM_TYPE_KEY," + "\r\n" + "MMS_STATE," + "\r\n" + "DSS_STATE," + "\r\n" + "AVL_STATE," + "\r\n" + "OPERATOR_KEY," + "\r\n" + "OPERATOR_NAME," + "\r\n" + "ALARM_COMMENT," + "\r\n" + "EVENT_LEVEL," + "\r\n" + "ALARM_ACK," + "\r\n" + "ALARM_STATUS," + "\r\n" + "SESSION_KEY," + "\r\n" + "SESSION_LOCATION," + "\r\n" + "PROFILE_ID," + "\r\n" + "ACTION_ID," + "\r\n" + "OPERATION_MODE," + "\r\n" + "ENTITY_KEY," + "\r\n" + "AVLALARMHEADID," + "\r\n" + "SYSTEM_KEY," + "\r\n" + "EVENT_ID," + "\r\n" + "ASSET_NAME," + "\r\n" + "SEVERITY_NAME," + "\r\n" + "EVENT_TYPE_NAME," + "\r\n" + "VALUE," + "\r\n" + "to_char(CREATEDATETIME,'DD/MM/YYYY HH12:MI:SS AM')," + "\r\n" + "CREATETIME," + "\r\n" + "DESCRIPTION FROM EV_COMBINED WHERE CREATETIME BETWEEN TO_DATE('");
    sqlFile.write(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss'));
    sqlFile.write("','DD/MM/YYYY HH24:MI:SS') AND TO_DATE('");
    sqlFile.write(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss'));
    sqlFile.writeline("','DD/MM/YYYY HH24:MI:SS') ORDER BY CREATETIME ASC ;");
    sqlFile.writeline("spool off");
    sqlFile.writeline("exit");
    sqlFile.close();

    var userName = stringElements[0],
        passWord = stringElements[1],
        dataBase = stringElements[2];

    var batFile = objFSO.OpentextFile(fileFolderName + '.bat', ForWriting, CreateIt, systemDefaultMode);
    batFile.writeline('title ' + dateName);
    batFile.writeline('cd "' + tempFolder + '\\"');
    batFile.writeline("sqlplus " + userName + "/" + passWord + "@" + dataBase + ' @"' + fileFolderName + '.sql' + '"');
    batFile.writeline("exit");
    batFile.close();


    //WScript.echo(dateReformatted(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss')) + '_to_'  + dateReformatted(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss')));

    if (missingFiles) { //read CSV file. If file firstline lastBuffer line exist = skip. else proceed to download SQL data file
        runSQLBatch(stringElements, tempFolder, fileFolderName, dateName);
        checkCountTallyWriteCSV(stringElements, startDateTime, endDateTime, dateName, fileFolderName, tempFolder, folderStore);
    } else if (!missingFiles) {
        var deleteOutputFile = folderStore + '\\' + dateName + '.CSV';
        var checkDelete = true; 
        while (!checkDelete) {
            WScript.sleep(10000)
            try {
                toWriteLog = '===================================\r\n';
                toWriteLog += js_yyyymmdd_hhmmss() + ' = Try to delete existing file = ' + deleteOutputFile + '\r\n';                 
                createLogFile(objStartFolder, toWriteLog);                
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);;
                }           
                customFileFolder.deleteFile(deleteOutputFile);
                checkDelete = false; 
            } catch (err) {
                checkDelete = true;
                toWriteLog = js_yyyymmdd_hhmmss() + ' = File still in use.' + '\r\n'; 
                createLogFile(objStartFolder, toWriteLog);                
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);;
                }
            }
        }        
        toWriteLog = js_yyyymmdd_hhmmss() + ' = Delete successful.' + '\r\n'; 
        toWriteLog += js_yyyymmdd_hhmmss() + ' = Doing a fresh download.' + '\r\n'; 
        toWriteLog += '===================================\r\n';                   
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {            
            WScript.echo(toWriteLog);;
        }
        runSQLBatch(stringElements, tempFolder, fileFolderName, dateName);
        checkCountTallyWriteCSV(stringElements, startDateTime, endDateTime, dateName, fileFolderName, tempFolder, folderStore);
        
    }

}

function runSQLBatch(stringElements, tempFolder, fileFolderName, dateName) {
    var toWriteLog = '';
    //var tempFolder = objStartFolder + '\\' + grabDateTime + '_tempFolder';
    var objStartFolder = stringElements[9]; //since run by Task Scheduler, need to write full directory name
    //var showHideLauncher = WScript.Arguments.Item(8); //2 = taskbar 1 = show echo // 0 = hide fully

    //var resourceUpdates = stringElements[7]; //seconds intervals to update into log file or WScript.echo

    var resourceUpdates;
    if (stringElements[7] == 'resourceUpdates') {
        resourceUpdates = 3000;
    } else {
        resourceUpdates = parseInt(stringElements[7] + '000');
    }

    //var myVar = setInterval(function(){ readFileCatcher (objStartFolder, fileFolderName + '.txt') }, parseInt(resourceUpdates)); //interval check SQL text file still downloading
    //var fileTxtSize = setInterval(function(){checkFileSize (tempFolder, objStartFolder, dateName, 'txt') }, parseInt(resourceUpdates)); //interval update

    //20180324 = need to add hide/show MS-DOS function

    collectResource(objStartFolder, false, scriptFolder);
    toWriteLog = '===================================\r\n';
    toWriteLog += js_yyyymmdd_hhmmss() + ' = Running SQL command.';
    createLogFile(objStartFolder, toWriteLog);
    if (showEcho == 'showEcho') {
        WScript.echo(toWriteLog);
    }


    var strCommand = '"' + fileFolderName + '.bat' + '"';
    switch (stringElements[8]) {
        case '2':
            objShell.Run(strCommand, 2, true);
            //WScript.echo (strCommand + ' = ' + 'minimize'); 
            break;
        case '1':
            objShell.Run(strCommand, 1, true);
            //WScript.echo (strCommand + ' = ' + 'showWindow'); 
            break;
        case '0':
            objShell.Run(strCommand, 0, true);
            //WScript.echo (strCommand + ' = ' + 'hideWindow '); 
            break;
        default:
            objShell.Run(strCommand, 0, true);
            //WScript.echo (strCommand + ' = ' + 'default = hideWindow '); 
    }

    var checkDelete = true; 
    while (!checkDelete) {
        WScript.sleep(10000)
        try {            
            toWriteLog = js_yyyymmdd_hhmmss() + ' = Checking if SQL file is still running = ' + fileFolderName + '.sql' ;
            createLogFile(objStartFolder, toWriteLog);
            if (showEcho == 'showEcho') {
                WScript.echo(toWriteLog);;
            }
			
			var readResult = objFSO.OpentextFile('"' + fileFolderName + '.txt' + '"', ForReading, dontWantCreateIt, systemDefaultMode);
			var lineCount = 0;
			while (!readResult.AtEndOfStream) {
				readResult.SkipLine();
				lineCount += 1;
				if ((lineCount % 9999) == 0) { WScript.echo("Line 2209: Still inside texStream inside " + '"' + colFiles.item() + '" Line: ' + lineCount ) }
			}
			readResult.close();			
			
			
            customFileFolder.deleteFile(fileFolderName + '.sql');
            checkDelete = false; 
        } catch (err) {
            checkDelete = true;
            toWriteLog = js_yyyymmdd_hhmmss() + ' = SQL file still running = ' + fileFolderName + '.sql' ;
            createLogFile(objStartFolder, toWriteLog);
            if (showEcho == 'showEcho') {
                WScript.echo(toWriteLog);;
            }            
        }
    }

    customFileFolder.deleteFile(fileFolderName + '.bat');
    toWriteLog = js_yyyymmdd_hhmmss() + ' = Completed writing SQL data to ' + fileFolderName + '.txt';
    collectResource(objStartFolder, false, scriptFolder);
    createLogFile(objStartFolder, toWriteLog);
    checkFileSize(tempFolder, objStartFolder, dateName, 'txt');
    if (showEcho == 'showEcho') {
        WScript.echo(toWriteLog);
    }
}

function checkCountTallyWriteCSV(stringElements, startDateTime, endDateTime, dateName, fileFolderName, tempFolder, folderStore) {
    var actualoutputFile = folderStore + '\\' + dateName + '.CSV';
    var actualInputFile = fileFolderName + ".txt";
    //var resourceUpdates = stringElements[7] + '000';
    var objStartFolder = stringElements[9];

    if (objFSO.FileExists(actualInputFile)) {
        var sqlCount = 0;
        var lineCount = 0;
        var copyCount = 0;
        var toWriteLog = '';

        //collect info from SQLPlus count(*)
        var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
        var fileStreamer = readSQLResult.AtEndOfStream;
        toWriteLog = js_yyyymmdd_hhmmss() + ' = Reading file = ' + actualInputFile + '\r\n';
        toWriteLog += js_yyyymmdd_hhmmss() + ' = Collecting SQL data count.';
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }

        while (!fileStreamer) {
            var readFileLine = readSQLResult.ReadLine(); //read the full SQLplus content into memory
            if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') == -1)) { //collect result of count(*) sql command
                sqlCount = trim(readFileLine);
                fileStreamer = true;
                break;
            }
        }
        readSQLResult.close();
        toWriteLog = js_yyyymmdd_hhmmss() + ' = inside checkCountTallyWriteCSV =  Count(*) result = ' + sqlCount;
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }	

			var firstLine, count = 0;
			var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
			fileStreamer = readSQLResult.AtEndOfStream;
			while (!fileStreamer) {
				var readFileLine = readSQLResult.ReadLine(); //read each line
				if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') != -1)) { //jump to first entry
                        toWriteLog = js_yyyymmdd_hhmmss() + ' = Found the relevant first data entry.';
                        createLogFile(objStartFolder, toWriteLog);
                        if (showEcho == 'showEcho') {
                            WScript.echo(toWriteLog);
                        }                    
					fileStreamer = true;
					count += 1;
					break;
				}
			}		
			
			fileStreamer = readSQLResult.AtEndOfStream;	
			count += 10;

			for (; count < parseInt(sqlCount); count++) {
				readSQLResult.SkipLine();
				if ((count % (sqlCount / 9)) == 0) {
					WScript.echo(js_yyyymmdd_hhmmss() + ' = Skipline at row number = ' + count);
				}				
			}			
		

			/*	
			var lastTenLength = 9999 * parseInt(sqlCount);     
			readSQLResult.Skip(fileSize - lastTenLength);
			WScript.echo ( js_yyyymmdd_hhmmss() + " = Skipped characters : " + (fileSize - lastTenLength) )
			*/

	
		//collect at end of SQL query which shows actual count	
        toWriteLog = js_yyyymmdd_hhmmss() + ' = Finding the SQL data rows collected.';
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }
        toWriteLog = '';

        var perLineRead;
        while (!readSQLResult.AtEndOfStream) {
            perLineRead = readSQLResult.ReadLine();
            if (perLineRead.indexOf('rows selected.') > 0) {
                //lineCount = perLineRead.replace(' rows selected.', ''); //collects the row count result of successfully spool result
                lineCount = parseInt(perLineRead.replace(' rows selected.', ''));
                toWriteLog = js_yyyymmdd_hhmmss() + ' = Found the SQL data rows collected.';
                createLogFile(objStartFolder, toWriteLog);
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);
                }
                collectResource(objStartFolder, false, scriptFolder);
                toWriteLog = '';
                break;
            }
        }
        readSQLResult.close();
		if (lineCount === 0)  {lineCount = "Not found.";}

        collectResource(objStartFolder, false, scriptFolder);

        toWriteLog = js_yyyymmdd_hhmmss() + ' = SQL result = ' + lineCount;
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);
        }

        if ((sqlCount == lineCount) & (sqlCount != 0) & (lineCount != 0)) {
            toWriteLog = js_yyyymmdd_hhmmss() + ' = Verified that collected data has rows matching EV_COMBINED.\r\n';
            toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = Proceed with processing/optimizing SQL data into ' + dateName + '.CSV';
            createLogFile(objStartFolder, toWriteLog);
            if (showEcho == 'showEcho') {
                WScript.echo(toWriteLog);
            }

            var writeHeader = 'SOURCE_TABLE, PKEY, SUBSYSTEM_KEY, PHYSICAL_SUBSYSTEM_KEY, LOCATION_KEY, SEVERITY_KEY, EVENT_TYPE_KEY, ALARM_ID, ALARM_TYPE_KEY, MMS_STATE, DSS_STATE, AVL_STATE, OPERATOR_KEY, OPERATOR_NAME, ALARM_COMMENT, EVENT_LEVEL, ALARM_ACK, ALARM_STATUS, SESSION_KEY, SESSION_LOCATION, PROFILE_ID, ACTION_ID, OPERATION_MODE, ENTITY_KEY, AVLALARMHEADID, SYSTEM_KEY, EVENT_ID, ASSET_NAME, SEVERITY_NAME, EVENT_TYPE_NAME, VALUE, CREATEDATETIME, CREATETIME, DESCRIPTION';
            var writeSQLResult = objFSO.OpentextFile(actualoutputFile, ForWriting, CreateIt, systemDefaultMode);
            writeSQLResult.write(writeHeader);
            var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
            while (!readSQLResult.AtEndOfStream) {
                var readFileLine = readSQLResult.ReadLine(); //read the full SQLplus content into memory' + '\r\n' +
                if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') > 0)) { //collect each row contents of the sql command
                    writeSQLResult.write('\r\n'); //to prevent last empty line		  
                    writeSQLResult.write(reformatCREATETIME(readFileLine.replace(/  +/g, ''))); //replace CREATETIME & remove spaces
                    copyCount += 1;
                    if ((copyCount % 10000) == 0) {
                        toWriteLog = js_yyyymmdd_hhmmss() + ' = Writing result to file ' + dateName + '.CSV\r\n';
                        toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = At row number ' + copyCount + '.';
                        createLogFile(objStartFolder, toWriteLog);
                        if (showEcho == 'showEcho') {
                            WScript.echo(toWriteLog);
                        }
                        collectResource(objStartFolder, false, scriptFolder);
                    }
                }
            }
            readSQLResult.close();
            writeSQLResult.close();

            if ((sqlCount == lineCount) && (copyCount != 0)) {
                //WScript.echo (sqlCount + ' --- line 499 ' + copyCount);                
                toWriteLog = js_yyyymmdd_hhmmss() + ' = Deleted raw data file from database ' + actualInputFile + '\r\n';
                toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = Processed output filename is ' + actualoutputFile;
                customFileFolder.deleteFile(actualInputFile);
                createLogFile(objStartFolder, toWriteLog);
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);
                }
                //checkFileSize(folderStore, objStartFolder, dateName, 'CSV');

				
                toWriteLog = '===================================\r\n';
                toWriteLog += js_yyyymmdd_hhmmss() + ' = Zipping up and to delete file ' + dateName + '.CSV' + '\r\n';
                toWriteLog += js_yyyymmdd_hhmmss() + ' = If works, file will zip up as ' + dateName + '.zip' + '\r\n';
                toWriteLog += '===================================\r\n';
                createLogFile(objStartFolder, toWriteLog);
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);
                }

                var zipSuccessful = false;
                //zipSuccessful = f_CreateZip(folderStore, dateName); //Removing zipping 2018_06_02
                zipSuccessful = oracleZipper (folderStore, dateName, zipperExe);
                //WScript.sleep(10000);

                toWriteLog = '===================================\r\n';
                if (zipSuccessful) {
                    toWriteLog += js_yyyymmdd_hhmmss() + ' = Zip successful. Filename is ' + dateName + '.zip' + '\r\n';
                } else if (!zipSuccessful) {
                    toWriteLog += js_yyyymmdd_hhmmss() + ' = Zip failed. Zip file deleted. ' + dateName + '.zip' + '\r\n';
                }
                toWriteLog += '===================================\r\n';
                createLogFile(objStartFolder, toWriteLog);
                checkFileSize(folderStore, objStartFolder, dateName, 'zip');
                if (showEcho == 'showEcho') {
                    WScript.echo(toWriteLog);
                }
				
            }
        } else {
            
            toWriteLog = '===================================\r\n';
            toWriteLog += toWriteLog + js_yyyymmdd_hhmmss() + ' = SQL downloaded file is incomplete. Affected file will be deleted. INSIDE checkCountTallyWriteCSV ' + '\r\n';
            toWriteLog += toWriteLog + js_yyyymmdd_hhmmss() + ' = Affected date not downloaded is between ' + dateReformatted(moment(startDateTime).format('DD/MM/YYYY HH:mm:ss')) + ' and ' + dateReformatted(moment(endDateTime).format('DD/MM/YYYY HH:mm:ss'));
            toWriteLog += toWriteLog + '===================================\r\n';
            customFileFolder.deleteFile(actualInputFile);
            createLogFile(objStartFolder, toWriteLog);
            if (showEcho == 'showEcho') {
                WScript.echo(toWriteLog);
            }
        }
        /*
        Microsoft (R) Windows Script Host Version 5.8
        Copyright (C) Microsoft Corporation. All rights reserved.

        rowCOUNT:,2146,2146

        D:\Hirman\sqlCollect>
        */
    }

}


function oracleZipper (folderStore, dateName, zipperExe) {
    var zipFilename = folderStore + '\\' + dateName + '.zip';
    var sourceFile = folderStore + '\\' + dateName + '.CSV';   
 
    if (objFSO.FileExists(zipperExe)) {
        //command = zip.exe -v -r 2018-09-03_060000_to_2018-09-03_120000.zip 2018-09-03_060000_to_2018-09-03_120000.CSV        
        
        var fileContent = js_yyyymmdd_hhmmss() + ' = oracleZipper Making zip file : ' + zipFilename + '\r\n';
        fileContent += js_yyyymmdd_hhmmss() + ' = oracleZipper zipping up : ' + sourceFile + '\r\n';
        createLogFile(objStartFolder, fileContent);

        var strCommand = comspec + ' /c "' + zipperExe + ' -r ' + zipFilename + ' ' + sourceFile + '"' ;
        try {
            objShell.Run(strCommand, 0, true);
        } catch (e) {
            createLogFile(objStartFolder, e.description);
            customFileFolder.deleteFile(zipFilename);
            zipSuccessful = false;
        }
        
        var getFileObject = objFSO.GetFile(zipFilename);
        var fileSize = getFileObject.size / 1024;
        if (parseInt(fileSize) < 260) {
            getFileObject.Attributes[0];
            getFileObject.Delete();
            var fileContent = js_yyyymmdd_hhmmss() + ' = oracleZipper : Deleting file because size is under 260KB : ' + zipFilename + '\r\n';
            fileContent += js_yyyymmdd_hhmmss() + ' = oracleZipper : Keeping CSV file for manual zip : ' + sourceFile + '\r\n';
            createLogFile(objStartFolder, fileContent);
            zipSuccessful = false;
        } else {
            customFileFolder.deleteFile(sourceFile);
            zipSuccessful = true;
        }
        return zipSuccessful;          
    }  
}

function sqlTestConnect(stringElements, grabDateTime) { //need to change to tnFsping + select * from dual
    //SELECT count(*) FROM DUAL ; tnsping

    var finalResult = false;

    var userName = stringElements[0],
        passWord = stringElements[1],
        dataBase = stringElements[2],
        splitFile = stringElements[3],
        endDated = stringElements[4],
        pastDays = stringElements[5],
        missingFiles = stringElements[6],
        resourceUpdates = stringElements[7],
        showHideLauncher = stringElements[8], //2 = taskbar 1 = show echo // 0 = hide fully
        objStartFolder = objFSO.GetFolder(stringElements[9]);

    if (showHideLauncher == 0) {
        showHideLauncher = shellRunOption.hideWindow();
    } else if (showHideLauncher == 1) {
        showHideLauncher = shellRunOption.activeRestore();
    } else if (showHideLauncher == 2) {
        showHideLauncher = shellRunOption.minimizeZorder();
    }

    var tempFolder = objStartFolder + '\\' + grabDateTime + '_tempFolder';
    customFileFolder.checkMakeFolder(objStartFolder);
    customFileFolder.checkMakeFolder(tempFolder);

    var sqlFile = objFSO.OpentextFile(tempFolder + '\\' + 'sqlCountOnly.sql', ForWriting, CreateIt, systemDefaultMode);
    sqlFile.writeline('set colsep ,' + '\r\n' + 'set headsep off' + '\r\n' + 'set pagesize 0' + '\r\n' + 'set trimspool off' + '\r\n' + 'set linesize 10' + '\r\n' + 'set numw 5' + '\r\n' + 'set echo on');
    sqlFile.writeline("spool '" + tempFolder + '\\' + "sqlCountOnly.txt'");
    sqlFile.writeline("");
    sqlFile.writeline("select count(*) from dual");
    sqlFile.writeline("");
    sqlFile.writeline("");
    sqlFile.writeline("spool off");
    sqlFile.writeline("");
    sqlFile.writeline("exit");
    sqlFile.writeline("exit");
    sqlFile.close();

    var batFile = objFSO.OpentextFile(tempFolder + '\\' + 'sqlCountOnly.bat', ForWriting, CreateIt, systemDefaultMode);
    batFile.writeline('cd "' + tempFolder + '"');
    batFile.writeline("sqlplus " + userName + "/" + passWord + "@" + dataBase + ' @"' + tempFolder + '\\' + 'sqlCountOnly.sql' + '"');
    batFile.writeline("");
    batFile.writeline('tnsping ' + dataBase + ' >> "' + tempFolder + '\\' + 'sqlCountOnly.txt"');
    batFile.close();

    var strCommand = comspec + ' /c "' + tempFolder + '\\' + 'sqlCountOnly.bat' + '"';
    objShell.Run(strCommand, 0, true);

    customFileFolder.deleteFile(tempFolder + '\\' + 'sqlCountOnly.sql');
    customFileFolder.deleteFile(tempFolder + '\\' + 'sqlCountOnly.bat');

    var dualResult = 'error';
    var tnsPingResult = 'Not OK'

    var readResult = objFSO.OpentextFile(tempFolder + '\\' + "sqlCountOnly.txt", ForReading, dontWantCreateIt, systemDefaultMode);
    while (!readResult.AtEndOfStream) {
        var perLine = readResult.ReadLine();
        if (perLine.indexOf('select count(*) from dual') > 0) {
            dualResult = readResult.ReadLine();
        } else if (perLine.indexOf('OK') > -1) {
            tnsPingResult = perLine;
            finalResult = true;
        }
    }
    readResult.close();
    customFileFolder.deleteFile(tempFolder + '\\' + "sqlCountOnly.txt");
    //customFileFolder.deleteFolder(tempFolder);

    var toWriteLog = '===================================\r\n';
    toWriteLog += js_yyyymmdd_hhmmss() + '\r\n';
    toWriteLog += 'SQL connect test \r\n';
    toWriteLog += 'Login: ' + userName + '\r\n';
    toWriteLog += 'Database: ' + dataBase + '\r\n';
    toWriteLog += 'Result: ' + tnsPingResult + ' | count(*) from dual = ' + dualResult + '\r\n';
    toWriteLog += '===================================';
    createLogFile(objStartFolder, toWriteLog);
    if (showEcho == 'showEcho') {
        WScript.echo(toWriteLog);
    }
    return finalResult;
}

function createLogFile(objStartFolder, toWriteLog) {
    if (objFSO.FolderExists(objStartFolder)) {
        var grabDateTime = js_yyyymmdd_hhmmss();
        var listArrayName = [],
            listArraySize = [];
        var objFolder = objFSO.GetFolder(objStartFolder);
        var colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            if (colFiles.item().name.split('.').pop().indexOf('log') == 0) {
                listArrayName.push(colFiles.item().name);
                listArraySize.push(parseInt(colFiles.item().size) / 1024 / 1024);
                //Wed Feb 14 12:57:20 UTC +0800 2018
                //listArrayModify.push (moment(colFiles.item().DateLastModified).isValid());
                //listArrayModify.push (moment(colFiles.item().DateLastModified, "DD/MM/YYYY HH:mm:ss"));
                //listArrayModify.push (colFiles.item().DateLastModified);
            }
        }

        if (listArraySize[listArraySize.length - 1] < 200) {
            try {
                var logFile = objFSO.OpentextFile(objFolder + '\\' + listArrayName[listArrayName.length - 1], ForAppending, dontWantCreateIt, systemDefaultMode);
                logFile.writeline(toWriteLog);
                logFile.close();
            } catch (err) {
                WScript.echo('Please close filename ' + listArrayName[listArrayName.length - 1] + '\r\nEV_COMBINED script trying to write logs to it.')
                var newLogFile = objFSO.OpentextFile(objFolder + '\\' + grabDateTime + '_LogFile.log', ForAppending, CreateIt, systemDefaultMode);
                newLogFile.writeline(toWriteLog);
                newLogFile.close();
            }
            /*
             finally {
              var newLogFile = objFSO.OpentextFile( objStartFolder + grabDateTime + '_LogFile.log'  , ForAppending, CreateIt, systemDefaultMode);
              newLogFile.writeline(toWriteLog);
              newLogFile.close();
            }
            */
        } else {
            var newLogFile = objFSO.OpentextFile(objFolder + '\\' + grabDateTime + '_LogFile.log', ForAppending, CreateIt, systemDefaultMode);
            newLogFile.writeline(toWriteLog);
            newLogFile.close();
        }
    }
}


function zipFileExist(folderStore, startDateTimeActual, endDateTimeActual) {
    //for each file found, if continuous need to form one array element. if gap, create another element.
    //if found gap element. will push to new array for return value to be use as collecting time duration
    var fileDateArr = [];
    var gapDate = [];
    var nameOnly = '';

    if (objFSO.FolderExists(folderStore)) {
        var foundFilesCSV = [];
        var foundFilesZIP = [];
        var objFolder = objFSO.GetFolder(folderStore);
        var colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            if (colFiles.item().name.split('.').pop().indexOf('zip') == 0) {
                var fileContent = {};
                fileContent.path = colFiles.item().name; //full folder name
                nameOnly = colFiles.item().name.split('\\');
                fileContent.name = nameOnly[nameOnly.length - 1]; //only the file name
                foundFilesZIP.push(fileContent);
            }
            if (colFiles.item().name.split('.').pop().indexOf('CSV') == 0) {
                var fileContent = {};
                fileContent.path = colFiles.item().name; //full folder name
                nameOnly = colFiles.item().name.split('\\');
                fileContent.name = nameOnly[nameOnly.length - 1]; //only the file name
                foundFilesCSV.push(fileContent);
            }            
        }

        var filesToDelete = [];
        for(var a=0; a<foundFilesZIP.length; a++) {
            for(var b=0; b<foundFilesCSV.length; b++){
                if(foundFilesCSV[b].name == foundFilesZIP[a].name) {
                    filesToDelete.push(foundFilesZIP[a]); //after successful zip, CSV files should be deleted, else means zip unsucessful, deletePrematureFiles incharge of checking if TXT already saved to CSV correctly
                }
            }
        }

        for (var c=0; c<filesToDelete.length; c++ ) {
            customFileFolder.deleteFile(filesToDelete[c]); //after successful zip, CSV files should be deleted, else means zip unsucessful,
        }
        filesToDelete = [];
        foundFilesCSV = [];
        foundFilesZIP = [];        

        filesToZip = [];
        objFolder = objFSO.GetFolder(folderStore);
        colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            if (colFiles.item().name.split('.').pop().indexOf('CSV') == 0) { 
                filesToZip.push(colFiles.item().name);
            }
        }

        for (var d=0; d<filesToZip.length; d++){
            nameOnly = filesToZip[d].split('\\');
            var dateName = nameOnly[nameOnly.length - 1].replace('CSV','');            
            oracleZipper (folderStore, dateName, zipperExe); //previous code dont have zip function. Added zip function + deletePrematureFiles incharge of checking if TXT already saved to CSV correctly
        }
        filesToZip = [];

        objFolder = objFSO.GetFolder(folderStore);
        colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            //if ((colFiles.item().name.split('.').pop().indexOf('zip') == 0) || (colFiles.item().name.split('.').pop().indexOf('CSV') == 0)) { //once all zip up, can comment out this line
            if (colFiles.item().name.split('.').pop().indexOf('zip') == 0) { //started using Zip. So dont check CSV as a relevant file entry                
                num = colFiles.item().size / 1024; //KB
                if (parseInt(num) < 3) {
                    colFiles.item().Delete(); //delete zip files which failed to zip up the csv files if size under 3KB
                    //customFileFolder.deleteFile(colFiles.item());
                } else {
                    var dateRetrieve = colFiles.item().name.split('.').shift().split("_to_");
                    //('DD/MM/YYYY HH:mm:ss')
                    //2017-12-27_120000_to_2017-12-27_115959
                    var fileStartDateTimeActual = moment(dateRetrieve[0], "YYYY-MM-DD_HHmmss"); //filename start
                    var fileEndDateTimeActual = moment(dateRetrieve[1], "YYYY-MM-DD_HHmmss"); //filename end
                    //fileDateStart.push(fileStartDateTimeActual);
                    //fileDateEnd.push(fileEndDateTimeActual);
                    fileDateArr.push([fileStartDateTimeActual, fileEndDateTimeActual]);
                }
            }
        }

        if (fileDateArr.length > 0) {

            fileDateArr.sort(function (a, b) {
                return a[1] - b[1]; //sort ending == early to later
            });

            fileDateArr.sort(function (a, b) {
                return a[0] - b[0]; //sort starting == early to later
            });

            var newStartFileName = fileDateArr[0][0];
            var newEndFileName = fileDateArr[0][1];

            //need to search if dates overlaps =========>>
            //WScript.echo (0 + " = " + moment(fileDateArr[0][0]).format('DD/MM/YYYY HH:mm:ss') + " to " + moment(fileDateArr[0][1]).format('DD/MM/YYYY HH:mm:ss'));
            var k = fileDateArr.length,
                w = 1;
            while (w < k) {
                if (fileDateArr[w - 1][0].diff(fileDateArr[w][0], 'minutes') == 0) { //starting same time for both files.
                    newStartFileName = fileDateArr[w][0];
                    if (fileDateArr[w][1].diff(fileDateArr[w - 1][1], 'minutes') < 0) { //fileDateArr[w-1][1] is later
                        newEndFileName = fileDateArr[w - 1][1];
                    }
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][1], 'minutes') < 0) { //fileDateArr[w][1] is later
                        newEndFileName = fileDateArr[w][1];
                    }
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][1], 'minutes') == 0) { //both same name = impossible
                        newEndFileName = fileDateArr[w][1];
                    }
                    fileDateArr.splice(w - 1, 2);
                    fileDateArr.push([newStartFileName, newEndFileName]);
                    k = fileDateArr.length;
                    w = 1;
                    fileDateArr.sort(function (a, b) {
                        return a[1] - b[1];
                    }); //sort ending == early to later
                    fileDateArr.sort(function (a, b) {
                        return a[0] - b[0];
                    }); //sort starting == early to later
                }
                w++;
            }

            k = fileDateArr.length, w = 1;
            while (w < k) {
                if ((fileDateArr[w - 1][0].diff(fileDateArr[w][0], 'minutes') < 0) && (fileDateArr[w][0].diff(fileDateArr[w - 1][1], 'minutes') < 0)) {
                    //fileDateArr[w][0] is later then fileDateArr[w-1][0] && fileDateArr[w-1][1] is later then fileDateArr[w][0]
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][1], 'minutes') < 0) { //fileDateArr[w][1] is later
                        newStartFileName = fileDateArr[w - 1][0];
                        newEndFileName = fileDateArr[w][1];
                    }
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][1], 'minutes') > 0) { //fileDateArr[w-1][1] is later	
                        newStartFileName = fileDateArr[w - 1][0];
                        newEndFileName = fileDateArr[w - 1][1];
                    }
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][1], 'minutes') == 0) { //both file end time same	
                        newStartFileName = fileDateArr[w - 1][0];
                        newEndFileName = fileDateArr[w - 1][1];
                    }
                    fileDateArr.splice(w - 1, 2);
                    fileDateArr.push([newStartFileName, newEndFileName]);
                    k = fileDateArr.length;
                    w = 1;
                    fileDateArr.sort(function (a, b) {
                        return a[1] - b[1];
                    }); //sort ending == early to later
                    fileDateArr.sort(function (a, b) {
                        return a[0] - b[0];
                    }); //sort starting == early to later
                }
                w++;
            }

            k = fileDateArr.length, w = 1;
            while (w < k) {
                if (fileDateArr[w - 1][1].diff(fileDateArr[w][0], 'minutes') == 0) { //firstEnd == nextStart
                    newStartFileName = fileDateArr[w - 1][0];
                    newEndFileName = fileDateArr[w][1];
                    fileDateArr.splice(w - 1, 2);
                    fileDateArr.push([newStartFileName, newEndFileName]);
                    k = fileDateArr.length;
                    w = 1;
                    fileDateArr.sort(function (a, b) {
                        return a[1] - b[1];
                    }); //sort ending == early to later
                    fileDateArr.sort(function (a, b) {
                        return a[0] - b[0];
                    }); //sort starting == early to later		
                }
                w++;
            }

            fileDateArr.sort(function (a, b) {
                return a[1] - b[1];
            }); //sort ending == early to later
            fileDateArr.sort(function (a, b) {
                return a[0] - b[0];
            }); //sort starting == early to later		



            if (fileDateArr.length > 0) {
                //need to find first date-time if match or is a gap
                if (startDateTimeActual.diff(fileDateArr[0][0], 'minutes') < 0) { //fileDateArr[0][0] is later than startDateTimeActual
                    gapDate.push([startDateTimeActual, fileDateArr[0][0]]);
                }
                //need to find last date-time if match or is a gap
                if (endDateTimeActual.diff(fileDateArr[fileDateArr.length - 1][1], 'minutes') > 0) { //fileDateArr[fileDateArr.length - 1][1] is earlier than endDateTimeActual
                    gapDate.push([fileDateArr[fileDateArr.length - 1][1], endDateTimeActual]);
                }
            }

            //go through all the array elements. since continuous element is merged into single element. any gaps is considered missing date time
            //create another array of those gaps //extreme edge begin and end already considered.
            // ======= need to consider what happens when date overlaps!!!!!!!
            if (fileDateArr.length > 1) {
                for (var w = 1; w < fileDateArr.length; w++) {
                    if (fileDateArr[w - 1][1].diff(fileDateArr[w][0], 'minutes') != 0) {
                        gapDate.push([fileDateArr[w - 1][1], fileDateArr[w][0]]);
                    }
                }
            }
            return gapDate;
        } else if (fileDateArr.length == 0) {
            gapDate.push([startDateTimeActual, endDateTimeActual]);
            return gapDate;
        }
    }

}

function checkFileSize(tempFolder, objStartFolder, dateName, extension) {
    if (objFSO.FolderExists(tempFolder)) {
        var listArrayName = [],
            listArraySize = [];
        var objFolder = objFSO.GetFolder(tempFolder);
        var colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            if ((colFiles.item().name.split('.').pop().indexOf(extension) == 0) && (colFiles.item().name.split('.')[0].indexOf(dateName) == 0)) {
                if (parseInt(colFiles.item().size / 1024) > 0) { //KB
                    num = colFiles.item().size / 1024;
                    toWriteLog = js_yyyymmdd_hhmmss() + ' = Filename: ' + tempFolder + '\\' + dateName + '.' + extension + ' Filesize: ' + parseInt(num) + 'KB';
                    createLogFile(objStartFolder, toWriteLog);
                    if (showEcho == 'showEcho') {
                        WScript.echo(toWriteLog);
                    }
                } else {
                    num = colFiles.item().size;
                    toWriteLog = js_yyyymmdd_hhmmss() + ' = Filename: ' + tempFolder + '\\' + dateName + '.' + extension + ' Filesize: ' + parseInt(num) + 'B';
                    createLogFile(objStartFolder, toWriteLog);
                    if (showEcho == 'showEcho') {
                        WScript.echo(toWriteLog);
                    }
                }
            }
        }
    }
}

//('DD/MM/YYYY HH:mm:ss')
//2017-12-27_120000_to_2017-12-27_115959
function dateReformatted(dateStringer) {
    var breakPieces = dateStringer.split(' ');
    var breakDate = breakPieces[0].split('/');
    var timing = breakPieces[1].replace(/:/g, "");;

    var chooseDay = breakDate[0];
    if (chooseDay.length == 1) {
        chooseDay = '0' + chooseDay;
    }

    var chooseMonth = breakDate[1];
    if (chooseMonth.length == 1) {
        chooseMonth = '0' + chooseMonth;
    }

    var chooseYear = breakDate[2];
    var newStringMerge = '';
    newStringMerge = chooseYear + '-' + chooseMonth + '-' + chooseDay + '_' + timing;
    return newStringMerge;
}

function reformatCREATETIME(rowEach) {
    var commaSplit = rowEach.split(',');
    var originalCreateTime = commaSplit[32];
    if ((originalCreateTime.indexOf(' AM') > 0) || (originalCreateTime.indexOf(' PM') > 0)) {
        var breakSpaces = originalCreateTime.split(' ');
        var datePart = breakSpaces[0];
        var timePart = breakSpaces[1];
        var amPM = breakSpaces[2];
        var dayPart = datePart.split('-')[0];
        var monPart = datePart.split('-')[1];
        var yerPart = datePart.split('-')[2];
        var hourPart = timePart.split('.')[0];
        var minsPart = timePart.split('.')[1];
        var secsPart = timePart.split('.')[2];
        var millPart = timePart.split('.')[3];
        switch (monPart) {
            case 'JAN':
                monPart = '01';
                break;
            case 'FEB':
                monPart = '02';
                break;
            case 'MAR':
                monPart = '03';
                break;
            case 'APR':
                monPart = '04';
                break;
            case 'MAY':
                monPart = '05';
                break;
            case 'JUN':
                monPart = '06';
                break;
            case 'JUL':
                monPart = '07';
                break;
            case 'AUG':
                monPart = '08';
                break;
            case 'SEP':
                monPart = '09';
                break;
            case 'OCT':
                monPart = '10';
                break;
            case 'NOV':
                monPart = '11';
                break;
            case 'DEC':
                monPart = '12';
                break;
        }
        var newCreateTime = dayPart + '/' + monPart + '/20' + yerPart + ' ' + hourPart + ':' + minsPart + ':' + secsPart + '.' + millPart + ' ' + amPM;
        var newRowEach = '';
        for (var a = 0; a < 32; a++) {
            newRowEach += commaSplit[a] + ',';
        }
        newRowEach += newCreateTime + ',' + commaSplit[33];
        return newRowEach;
    } else {
        var newRowEach = '';
        for (var a = 0; a < 33; a++) {
            newRowEach += commaSplit[a] + ',';
        }
        return newRowEach;
    }

}

function trim(stringToTrim) {
    return stringToTrim.replace(/^\s+|\s+$/g, "");
}

function ltrim(stringToTrim) {
    return stringToTrim.replace(/^\s+/, "");
}

function rtrim(stringToTrim) {
    return stringToTrim.replace(/\s+$/, "");
}

function collectResource(objStartFolder, collectInfo, scriptFolder) { //for checking PC resource
    var diskDrive = objStartFolder.split(':')[0];
    var nowDateTime = js_yyyymmdd_hhmmss();
    var nowDate = nowDateTime.split('_')[0];
    var nowTime = nowDateTime.split('_')[1];


    var vbsFile = objFSO.OpentextFile(scriptFolder + 'collectSQLData.vbs', ForWriting, CreateIt, systemDefaultMode);
    vbsFile.writeline('dim objArgs, showHideLauncher, logFileFolder, objFSO, nameSplitter, fileStream' + '\r\n' + 'set objArgs = WScript.Arguments' + '\r\n' + 'set objFSO = CreateObject("Scripting.FileSystemObject")' + '\r\n' + 'uptimeDays = 0' + '\r\n' + 'uptimeHrs = 0 ' + '\r\n' + 'uptimeMin = 0 ' + '\r\n' + 'thisComputer = "localhost"' + '\r\n' + 'if objArgs.Count = 1 then' + '\r\n' + '  logFileFolder = objArgs(0) ' + '\r\n' + '  nameSplitter = split(logFileFolder,":\\")' + '\r\n' + 'Set fileStream = objFSO.OpentextFile(logFileFolder, 2, True, -2)' + '\r\n' + 'fileStream.writeline "CPU Free: " & TotalCPU() & "%"' + '\r\n' + 'fileStream.writeline "Mem Available: " & round(availMemory()/1024/1024) & "MB"' + '\r\n' + 'fileStream.writeline "Disk Free: " & diskSpace(UCase(nameSplitter(0)))' + '\r\n' + 'fileStream.writeline "Disk Queue: " & diskSpeed(UCase(nameSplitter(0)))' + '\r\n' + 'fileStream.writeline "PC uptime: " & fnUptime(thisComputer)' + '\r\n' + 'fileStream.Close' + '\r\n' + 'end if' + '\r\n' + 'set objArgs = nothing' + '\r\n' + 'set showHideLauncher = nothing' + '\r\n' + 'set logFileFolder = nothing' + '\r\n' + 'set objFSO = nothing' + '\r\n' + 'set nameSplitter = nothing' + '\r\n' + 'Function fnUptime(strComputer) ' + '\r\n' + '    Set objWMIService = GetObject("winmgmts:\\\\" & strComputer & "\\root\\cimv2") ' + '\r\n' + '    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem") ' + '\r\n' + '    For Each objOS in colOperatingSystems ' + '\r\n' + '        dtmBootup = objOS.LastBootUpTime ' + '\r\n' + '        dtmLastBootupTime = WMIDateStringToDate(dtmBootup) ' + '\r\n' + '        dtmSystemUptime = DateDiff("n", dtmLastBootUpTime, Now)' + '\r\n' + '    Next      ' + '\r\n' + '    fnUptime = timeConversion(dtmSystemUptime)' + '\r\n' + 'End Function ' + '\r\n' + 'Function WMIDateStringToDate(dtmBootup) ' + '\r\n' + '    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & _ ' + '\r\n' + '        Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _ ' + '\r\n' + '            & " " & Mid (dtmBootup, 9, 2) & ":" & _ ' + '\r\n' + '                Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup,13, 2)) ' + '\r\n' + 'End Function   ' + '\r\n' + 'Function timeConversion(dtmSystemUptime) ' + '\r\n' + '    uptimeMin = dtmSystemUptime ' + '\r\n' + '    if uptimeMin >= 60 then ' + '\r\n' + '        uptimeHrs = Int(uptimeMin / 60)' + '\r\n' + '        uptimeMin = (uptimeMin mod 60)' + '\r\n' + '    end if ' + '\r\n' + '    if uptimeHrs >= 24 then ' + '\r\n' + '        uptimeDays = Int(uptimeHrs / 24)' + '\r\n' + '        uptimeHrs = (uptimeHrs mod 24)' + '\r\n' + '    end if ' + '\r\n' + '    timeConversion = uptimeDays & " Days " & uptimeHrs & " Hours " & uptimeMin & " Minutes"' + '\r\n' + 'End Function ' + '\r\n' + '' + '\r\n' + '' + '\r\n' + 'Sub sleepingVB()' + '\r\n' + '	Set oShell = CreateObject("WScript.Shell")' + '\r\n' + '	cmd = "%COMSPEC% /c ping -n 1 127.0.0.1>nul"' + '\r\n' + '	oShell.Run cmd,0,1' + '\r\n' + '	Set oShell = Nothing' + '\r\n' + 'End Sub' + '\r\n' + '' + '\r\n' + 'Function TotalCPU()' + '\r\n' + '    Dim objService, objInstance1, objInstance2, N1, N2, D1, D2' + '\r\n' + '    Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\root\\cimv2")' + '\r\n' + '    Set objInstance1 = objService.Get("Win32_PerfRawData_PerfOS_Processor.Name=' + "'_Total'" + '")' + '\r\n' + '    N1 = objInstance1.PercentProcessorTime' + '\r\n' + '    D1 = objInstance1.TimeStamp_Sys100NS' + '\r\n' + '    call sleepingVB()' + '\r\n' + '    Set objInstance2 = objService.Get("Win32_PerfRawData_PerfOS_Processor.Name=' + "'_Total'" + '")' + '\r\n' + '    N2 = objInstance2.PercentProcessorTime' + '\r\n' + '    D2 = objInstance2.TimeStamp_Sys100NS' + '\r\n' + '    Nd = (N2 - N1)' + '\r\n' + '    Dd = (D2-D1)' + '\r\n' + '    TotalCPU = Round(((Nd/Dd)  *100), 2)' + '\r\n' + 'End Function' + '\r\n' + 'Function diskSpace(driveLetter)' + '\r\n' + '    Dim objWMIService, objItem, colItems, strComputer, myArray()' + '\r\n' + '    Redim myArray(0)' + '\r\n' + '    On Error Resume Next' + '\r\n' + '    strComputer = "."' + '\r\n' + '    Set objWMIService = GetObject("winmgmts:\\\\" & strComputer & "\\root\\cimv2")' + '\r\n' + '    Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")' + '\r\n' + '    For Each objItem in colItems' + '\r\n' + '        If objItem.Name = driveLetter & ":" Then' + '\r\n' + '            myArray (Ubound(myArray)) = objItem.Name' + '\r\n' + '            Redim Preserve myArray (Ubound(myArray)+ 1)' + '\r\n' + '            myArray (Ubound(myArray)) = Int(objItem.FreeSpace /1073741824) & "GB"' + '\r\n' + '            Redim Preserve myArray (Ubound(myArray)+ 1)' + '\r\n' + '        End If' + '\r\n' + '    Next' + '\r\n' + '    diskSpace = join(myArray,"")' + '\r\n' + 'End Function' + '\r\n' + 'Function diskSpeed(driveLetter)' + '\r\n' + '    Dim objWMIService, objCimv2, strComputer, myArray(), objRefresher, colDiskDrives, objDiskDrive, objDiskQueue' + '\r\n' + '    Redim myArray(0)' + '\r\n' + '    strComputer = "."' + '\r\n' + '    set objRefresher = CreateObject("WbemScripting.SWbemRefresher")' + '\r\n' + '    Set objCimv2 = GetObject("winmgmts:root\\cimv2")' + '\r\n' + '    Set objDiskQueue = objRefresher.AddEnum (objCimv2,"Win32_PerfFormattedData_PerfDisk_LogicalDisk").ObjectSet' + '\r\n' + '    objRefresher.Refresh' + '\r\n' + '    For each intDiskQueue in objDiskQueue' + '\r\n' + '        If intDiskQueue.Name = driveLetter & ":" Then' + '\r\n' + '            If intDiskQueue.CurrentDiskQueueLength > 2 Then' + '\r\n' + '                myArray (Ubound(myArray)) = "BUSY"' + '\r\n' + '            Else' + '\r\n' + '                myArray (Ubound(myArray)) = "NORMAL"' + '\r\n' + '            End If' + '\r\n' + '        End If' + '\r\n' + '    Next' + '\r\n' + '    diskSpeed = join(myArray,"")' + '\r\n' + 'End Function' + '\r\n' + '' + '\r\n' + 'Function availMemory()' + '\r\n' + '	Dim oWMI, Instance' + '\r\n' + '	Set oWMI = GetObject("WINMGMTS:\\\\.\\ROOT\\cimv2")' + '\r\n' + '    Set Instance = oWMI.Get("Win32_PerfFormattedData_PerfOS_Memory=@")' + '\r\n' + '    availMemory = Instance.AvailableBytes' + '\r\n' + 'End Function');
    vbsFile.close();

    var strCommand = comspec + ' /c cscript //nologo "' + scriptFolder + 'collectSQLData.vbs" "' + objStartFolder + 'resource"';
    //cscript .\src\collectSQLData.vbs a "C:\ViewToggle\02-sqlPlusSelect\01 toCSCRIPT" > resource
    objShell.Run(strCommand, shellRunOption.hideWindow(), true);


    if (objFSO.FileExists(objStartFolder + 'resource')) {
        var fileContent = '===================================\r\n';
        fileContent = fileContent + js_yyyymmdd_hhmmss() + '\r\n';
        var fileResource = objFSO.OpentextFile(objStartFolder + 'resource', ForReading, dontWantCreateIt, systemDefaultMode);
        while (!fileResource.AtEndOfStream) {
            fileContent = fileContent + fileResource.ReadAll();
        }
        fileResource.close();
        fileContent += '===================================\r\n';
        createLogFile(objStartFolder, fileContent);
    }

    customFileFolder.deleteFile(objStartFolder + 'resource');
    customFileFolder.deleteFile(scriptFolder + 'collectSQLData.vbs');
    if (collectInfo) {
        return fileContent
    };
}

function deleteOlderZip(folderStore, deleteDays) {
    var daysBefore = "-" + deleteDays;
    if (objFSO.FolderExists(folderStore)) {
        /*
        var start = new Date();
        var dateNow = '';
        dateNow = start.getDate();
        dateNow = dateNow + '';
        if (dateNow.length == 1) {
            dateNow = '0' + dateNow
        }
        var monthNow = 0;
        monthNow = start.getMonth() + 1;
        monthNow = monthNow + '';
        if (monthNow.length == 1) {
            monthNow = '0' + monthNow
        }
        var fullYearNow = start.getFullYear();
        var dateString = dateNow + '/' + monthNow + '/' + fullYearNow + ' 00:00:00';
        var nowDateTimeActual = moment(dateString, "DD/MM/YYYY HH:mm:ss"); //object
        */

        var availFiles = [];
        var objFolder = objFSO.GetFolder(folderStore);
        var colFiles = new Enumerator(objFolder.files);
        for (; !colFiles.atEnd(); colFiles.moveNext()) {
            if ((colFiles.item().name.split('.').pop().indexOf('zip') == 0) || (colFiles.item().name.split('.').pop().indexOf('CSV') == 0)) {
                var dateRetrieve = colFiles.item().name.split('.').shift().split("_to_");
                //('DD/MM/YYYY HH:mm:ss')
                //2017-12-27_120000_to_2017-12-27_115959
                var fileStartDateTimeActual = moment(dateRetrieve[0], "YYYY-MM-DD_HHmmss"); //filename start
                var fileEndDateTimeActual = moment(dateRetrieve[1], "YYYY-MM-DD_HHmmss"); //filename end
                //fileDateStart.push(fileStartDateTimeActual);
                //fileDateEnd.push(fileEndDateTimeActual);

                

                if (fileStartDateTimeActual.diff(moment(), 'days') < daysBefore ) { //deleting zip & CSV files older than 18 days
                    colFiles.item().Delete();
                }
            }
        }

    }
}

function deleteRogueFiles(objStartFolder) {
    var arrLogFiles = [];
    var otherTextFiles = [];
    var objFolder = objFSO.GetFolder(objStartFolder);
    var colFiles = new Enumerator(objFolder.files);
    for (; !colFiles.atEnd(); colFiles.moveNext()) {
        if ((colFiles.item().name.split('.').pop().indexOf('log') == 0) && (colFiles.item().name.indexOf('_LogFile') > 0) && (colFiles.item().name.length == 27)) {
            arrLogFiles.push(colFiles.item());
        } else {
            colFiles.item().Delete();
            //WScript.echo (colFiles.item());
        }
    }

    for (var k = 0; k < arrLogFiles.length - 4; k++) {
        arrLogFiles[k].Delete();
        //WScript.echo (arrLogFiles[k]);
    }


    var subfolders = new Enumerator(objFolder.SubFolders);
    for (; !subfolders.atEnd(); subfolders.moveNext()) {
        if (!((subfolders.item().name.indexOf('folderStore') > -1) || (subfolders.item().name.indexOf('scriptFiles') > -1) || (subfolders.item().name.indexOf('_tempFolder') > -1))) {
            subfolders.item().Delete();
            //WScript.echo (subfolders.item().name);
        }
    }
}

function deletePrematureFiles(objStartFolder) {
    var csvFiles = [];
    var folderStore = objStartFolder + '\\folderStore';

    var storeFolder = objFSO.GetFolder(folderStore);
    var storeFiles = new Enumerator(storeFolder.files);
    for (; !storeFiles.atEnd(); storeFiles.moveNext()) {
        if (storeFiles.item().name.split('.').pop().indexOf('CSV') == 0) {
            csvFiles.push(storeFiles.item());
        }
    }
    
    var folderToDelete = [],
        fileToDelete = []
		txtFiles = [],
		matchPairs = [];
		
	var objFolder = objFSO.GetFolder(objStartFolder);
	var subfolders = new Enumerator(objFolder.SubFolders);	
	for (; !subfolders.atEnd(); subfolders.moveNext()) {
		if (subfolders.item().name.indexOf('_tempFolder') > -1) {
			var tempFolder = objFSO.GetFolder(subfolders.item());
			var colFiles = new Enumerator(tempFolder.files);	
				for (; !colFiles.atEnd(); colFiles.moveNext()) {
                    /*
					if (colFiles.item().name.split('.').pop().indexOf('txt') == 0) {
						txtFiles.push(colFiles.item());
						
						toWriteLog = js_yyyymmdd_hhmmss() + ' = Inside deletePrematureFiles = Checking if file still open = ' + colFiles.item() + '\r\n';
						createLogFile(objStartFolder, toWriteLog);
						if (showEcho == 'showEcho') {
							WScript.echo(toWriteLog);;
						}						
						
						//collect info from SQLPlus count(*)
						var sqlCount = 0;
						var readResult = objFSO.OpentextFile(colFiles.item(), ForReading, dontWantCreateIt, systemDefaultMode);
						var fileStreamer = readResult.AtEndOfStream;
						while (!fileStreamer) {
							var readFileLine = readResult.ReadLine(); //read each line SQLplus content into memory
							if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') == -1)) { //collect result of count(*) sql command
								sqlCount = trim(readFileLine);
								fileStreamer = true;
								break;
							}
						}
						readResult.close();
						
						var readResult = objFSO.OpentextFile(colFiles.item(), ForReading, dontWantCreateIt, systemDefaultMode);
                        var lineCount = 0;
                        
						//while (lineCount < parseInt(sqlCount)) {
						//	readResult.SkipLine();
						//	lineCount += 1;
                        //}                        

						while (!readResult.AtEndOfStream) {
							readResult.SkipLine();
							lineCount += 1;
							if ((lineCount % 99) == 0) { 
								toWriteLog = js_yyyymmdd_hhmmss() + " = Line 3152: Still inside texStream = " + colFiles.item() + ' Line: ' + lineCount + '\r\n';
								createLogFile(objStartFolder, toWriteLog);
								if (showEcho == 'showEcho') {
									WScript.echo(toWriteLog);;
								}								

							}
						}
						readResult.close();
                    }
                    */
					if (colFiles.item().name.split('.').pop().indexOf('bat') == 0) {
						customFileFolder.deleteFile(colFiles.item());
                    }				
					if (colFiles.item().name.split('.').pop().indexOf('sql') == 0) {
                        var checkDelete = true; 
                        while (checkDelete) {
                            WScript.sleep(10000)
                            try {	
                                toWriteLog = js_yyyymmdd_hhmmss() + ' = Inside deletePrematureFiles = Try to delete SQL run file ' + colFiles.item() + '\r\n';
                                createLogFile(objStartFolder, toWriteLog);
                                if (showEcho == 'showEcho') {
                                    WScript.echo(toWriteLog);;
                                }
                                customFileFolder.deleteFile(colFiles.item());  
                                checkDelete = false;                              
                            } catch (err) {
                                checkDelete = true;             
                                toWriteLog = js_yyyymmdd_hhmmss() + ' = SQL file still running' + fileFolderName + '.sql' ;
                                createLogFile(objStartFolder, toWriteLog);
                                if (showEcho == 'showEcho') {
                                    WScript.echo(toWriteLog);;
                                }                                
                            }
                        }						
					}		
				}				
		}
	}	
	
	
	if (csvFiles.length > 0) {
		for (var k = 0; k < csvFiles.length; k++) {
			if (txtFiles.length > 0) { //never run becos of line 3172 commented off
				for (var j = 0; j < txtFiles.length; j++) {
					if ((objFSO.FileExists(csvFiles[k]) && objFSO.FileExists(txtFiles[j])) && (objFSO.GetBaseName(csvFiles[k]) == objFSO.GetBaseName(txtFiles[j]))){
						var actualInputFile = txtFiles[j];
						var actualoutputFile = csvFiles[k];
						
						fileContent = '===================================\r\n';
						fileContent += js_yyyymmdd_hhmmss() + '\r\n';
						fileContent += 'Deleting CSV file due to incomplete TXT to CSV conversion.\r\n'; //CSV should appear as itself
						fileContent += actualoutputFile  + '\r\n' + '\r\n';
						fileContent += 'Text file should be auto delete, once TXT to CSV conversion is completed.\r\n';
						fileContent += 'If text file remains means CSV is incomplete.\r\n\r\n';
						fileContent += 'Will proceed to check if left over TXT file is relevant by reading the rows before processing.\r\n';
						fileContent += 'Stop the script and delete affected files, if you dont wish to continue with the TXT to CSV conversion.\r\n';
						fileContent += actualInputFile + '\r\n';
						fileContent += '===================================\r\n';
						createLogFile(objStartFolder, fileContent);                                    
						if (showEcho == 'showEcho') {
							WScript.echo(fileContent)
						}
						//Start Processing TXT SQL data file
						//customFileFolder.deleteFile(colFiles.item());
						var writeSQLResult = objFSO.OpentextFile(actualoutputFile, ForWriting, CreateIt, systemDefaultMode); //empty out the problematic CSV file
						writeSQLResult.write('');
						writeSQLResult.close();									
						readSQLRaw (folderStore, actualInputFile, actualoutputFile);
						txtFiles.shift();
					}
				}
			}		
		}	
	}
	
	if (txtFiles.length > 0) { //never run becos of line 3172 commented off
		for (var j = 0; j < txtFiles.length; j++) {
			if  ((typeof txtFiles[j] != 'undefined') && (typeof txtFiles[j] != null) && objFSO.FileExists(txtFiles[j])) {
				var actualInputFile = txtFiles[j];
				var newCSVFile = objFSO.GetBaseName(txtFiles[j]);
				var actualoutputFile = folderStore + '\\' + newCSVFile + '.CSV';
				//WScript.echo (actualoutputFile);
				fileContent = '===================================\r\n';
				fileContent += js_yyyymmdd_hhmmss() + '\r\n';
				fileContent += 'Checking if leftover TXT RAW file is still relevant before deleting\r\n'; //CSV should appear as itself
				fileContent += actualInputFile + '\r\n' + '\r\n';
				fileContent += 'Will proceed with TXT to CSV conversion if found SQL raw TXT file still relevant.\r\n';
				fileContent += '===================================\r\n';
				createLogFile(objStartFolder, fileContent);                                    
				if (showEcho == 'showEcho') {
					WScript.echo(fileContent)
				}							
				readSQLRaw (folderStore, actualInputFile, actualoutputFile);		
			}
		}	
	}
	

	objFolder = objFSO.GetFolder(objStartFolder);
	subfolders = new Enumerator(objFolder.SubFolders);	
	for (; !subfolders.atEnd(); subfolders.moveNext()) {
		if (subfolders.item().name.indexOf('_tempFolder') > -1) {
			WScript.echo(subfolders.item());
			folderToDelete.push(subfolders.item());
		}
	}	
	
	for (var k = 0; k < folderToDelete.length; k++) {
		WScript.echo(folderToDelete[k]);
		WScript.echo("Line 3216");
		customFileFolder.deleteFolder(folderToDelete[k]);
	}
	folderToDelete = [];	

} //delete incomplete CSV file and see if can re-process TXT file


function readSQLRaw (folderStore, actualInputFile, actualoutputFile) {
	if (objFSO.FileExists(actualInputFile)) {
		var sqlCount = 0;
		var lineCount = 0;
		var toWriteLog = '';

		//collect info from SQLPlus count(*)

		toWriteLog = js_yyyymmdd_hhmmss() + ' = Reading file = ' + actualInputFile + '\r\n';
		toWriteLog += js_yyyymmdd_hhmmss() + ' = Collecting SQL data count.';
		createLogFile(objStartFolder, toWriteLog);
		if (showEcho == 'showEcho') {
			WScript.echo(toWriteLog);;
		}
		var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
		var fileStreamer = readSQLResult.AtEndOfStream;
		while (!fileStreamer) {
			var readFileLine = readSQLResult.ReadLine(); //read each line into memory
			if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') == -1)) { //collect result of count(*) sql command
				sqlCount = trim(readFileLine);
				fileStreamer = true;
				break;
			}
		}
		readSQLResult.close();

		toWriteLog = js_yyyymmdd_hhmmss() + ' = inside readSQLRaw = Count(*) result = ' + sqlCount;
		createLogFile(objStartFolder, toWriteLog);
		if (showEcho == 'showEcho') {
			WScript.echo(toWriteLog);;
		}

		//collect at end of SQL query which shows actual count
		
        var getFileObject = objFSO.GetFile(actualInputFile);
        var fileSize = getFileObject.size;
		
        var firstLine, count = 0;
        var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
		fileStreamer = readSQLResult.AtEndOfStream;
        while (!fileStreamer) {
            var readFileLine = readSQLResult.ReadLine(); //read each line
            if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') != -1)) { //jump to first entry
                fileStreamer = true;
				WScript.echo(js_yyyymmdd_hhmmss() + ' = Found the relevant first data entry.');
				count += 1;
                break;
            }
        }		

		fileStreamer = readSQLResult.AtEndOfStream;	
		count += 10;

        while (!fileStreamer) {
            for (; count < parseInt(sqlCount); count++) {
                readSQLResult.SkipLine();
                if ((count % (sqlCount / 9)) == 0) {
                    WScript.echo(js_yyyymmdd_hhmmss() + ' = Skipline at row number = ' + count);
                }				
            }
        }		

        toWriteLog = js_yyyymmdd_hhmmss() + ' = Finding the SQL data rows collected.';
        createLogFile(objStartFolder, toWriteLog);
        if (showEcho == 'showEcho') {
            WScript.echo(toWriteLog);;
        }
        toWriteLog = '';
		
		
			var perLineRead;
			while (!readSQLResult.AtEndOfStream) {
				perLineRead = readSQLResult.ReadLine();
				count += 1;
				WScript.echo(js_yyyymmdd_hhmmss() + ' = Row read at line number = ' + count);
			
				
				if (perLineRead.indexOf('rows selected.') > 0) {					
                    //lineCount = perLineRead.replace(' rows selected.', ''); //collects the row count result of successfully spool result
                    lineCount = parseInt(perLineRead.replace(' rows selected.', ''));
					toWriteLog = js_yyyymmdd_hhmmss() + ' = Found the SQL data rows collected.';
					createLogFile(objStartFolder, toWriteLog);
					if (showEcho == 'showEcho') {
						WScript.echo(toWriteLog);;
					}
					toWriteLog = '';
					break;
				}
			} 
			readSQLResult.close();
			if (lineCount === 0){lineCount = "Not found.";}
			
			
		//}


		toWriteLog = js_yyyymmdd_hhmmss() + ' = SQL result = ' + lineCount;
		createLogFile(objStartFolder, toWriteLog);
		if (showEcho == 'showEcho') {
			WScript.echo(toWriteLog);
		}

		if ((sqlCount == lineCount) && (sqlCount != 0) && (lineCount != "Not found.") ) {
			toWriteLog = js_yyyymmdd_hhmmss() + ' = Verified that collected data has rows matching EV_COMBINED.\r\n';
			toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = Proceed with processing/optimizing SQL data into ' + actualoutputFile ;
			createLogFile(objStartFolder, toWriteLog);
			if (showEcho == 'showEcho') {
				WScript.echo(toWriteLog);
			}

			var writeHeader = 'SOURCE_TABLE, PKEY, SUBSYSTEM_KEY, PHYSICAL_SUBSYSTEM_KEY, LOCATION_KEY, SEVERITY_KEY, EVENT_TYPE_KEY, ALARM_ID, ALARM_TYPE_KEY, MMS_STATE, DSS_STATE, AVL_STATE, OPERATOR_KEY, OPERATOR_NAME, ALARM_COMMENT, EVENT_LEVEL, ALARM_ACK, ALARM_STATUS, SESSION_KEY, SESSION_LOCATION, PROFILE_ID, ACTION_ID, OPERATION_MODE, ENTITY_KEY, AVLALARMHEADID, SYSTEM_KEY, EVENT_ID, ASSET_NAME, SEVERITY_NAME, EVENT_TYPE_NAME, VALUE, CREATEDATETIME, CREATETIME, DESCRIPTION';
			var writeSQLResult = objFSO.OpentextFile(actualoutputFile, ForWriting, CreateIt, systemDefaultMode);
			writeSQLResult.write(writeHeader);
			var readSQLResult = objFSO.OpentextFile(actualInputFile, ForReading, dontWantCreateIt, systemDefaultMode);
			var copyCount = 0;                                            
			while (!readSQLResult.AtEndOfStream) {
				var readFileLine = readSQLResult.ReadLine(); //read the full SQLplus content into memory' + '\r\n' +
				if ((readFileLine.length > 9999) && (readFileLine.indexOf(',') > 0)) { //collect each row contents of the sql command                                                    
					writeSQLResult.write('\r\n'); //to prevent last empty line		  
					copyCount += 1;
					if ((copyCount % (sqlCount / 9)) == 0) {
						WScript.echo(js_yyyymmdd_hhmmss() + ' = Processing data at row number = ' + copyCount);
					}
					writeSQLResult.write(reformatCREATETIME(readFileLine.replace(/  +/g, ''))); //replace CREATETIME & remove spaces
				}
			}
			readSQLResult.close();
			writeSQLResult.close();

			if ((sqlCount == lineCount) && (copyCount != 0)) {
				//toWriteLog = js_yyyymmdd_hhmmss() + ' = Inside readSQLRaw = Deleted raw data file from database ' + actualInputFile + '\r\n';
				
				var checkDelete = true; 
				while (checkDelete) {
					WScript.sleep(1000)
					try {
						toWriteLog = js_yyyymmdd_hhmmss() + ' = Inside readSQLRaw = Try delete created file ' + actualInputFile + '\r\n';
                        createLogFile(objStartFolder, toWriteLog);
                        if (showEcho == 'showEcho') {
                            WScript.echo(toWriteLog);
                        }                        					   
                        customFileFolder.deleteFile(actualInputFile);
                        checkDelete = false; 
					} catch (err) {
						checkDelete = true;
						toWriteLog = js_yyyymmdd_hhmmss() + ' = Inside readSQLRaw = Fail to delete created file ' + actualInputFile + '\r\n';
                        createLogFile(objStartFolder, toWriteLog);
                        if (showEcho == 'showEcho') {
                            WScript.echo(toWriteLog);
                        } 

					}
				}				
				
				//customFileFolder.deleteFile(actualInputFile);
				toWriteLog = toWriteLog + js_yyyymmdd_hhmmss() + ' = Processed output filename is ' + actualoutputFile;
				createLogFile(objStartFolder, toWriteLog);
				if (showEcho == 'showEcho') {
					WScript.echo(toWriteLog);
				}
				checkFileSize(folderStore, objStartFolder, objFSO.GetBaseName(actualoutputFile), 'CSV');

				
				toWriteLog = '===================================\r\n';
				toWriteLog += js_yyyymmdd_hhmmss() + ' = At readSQLRaw : Zipping up and to delete file ' + objFSO.GetBaseName(actualoutputFile) + '.CSV' + '\r\n';
				toWriteLog += js_yyyymmdd_hhmmss() + ' = At readSQLRaw : If works, file will zip up as ' + objFSO.GetBaseName(actualoutputFile) + '.zip' + '\r\n';
				toWriteLog += '===================================\r\n';
				createLogFile(objStartFolder, toWriteLog);
				if (showEcho == 'showEcho') {
					WScript.echo(toWriteLog);
				}
				var zipSuccessful = false
				//zipSuccessful = f_CreateZip(folderStore, objFSO.GetBaseName(actualoutputFile));
				zipSuccessful = oracleZipper (folderStore, objFSO.GetBaseName(actualoutputFile), zipperExe);

				toWriteLog = '===================================\r\n';
				if (zipSuccessful) {
					toWriteLog += js_yyyymmdd_hhmmss() + ' = At readSQLRaw : Zip successful.' + '\r\n';
				} else if (!zipSuccessful) {
					toWriteLog += js_yyyymmdd_hhmmss() + ' = At readSQLRaw : Zip failed. Zip file deleted.' + '\r\n';
				}
				toWriteLog += '===================================\r\n';
				createLogFile(objStartFolder, toWriteLog);
				if (showEcho == 'showEcho') {
					WScript.echo(toWriteLog);
				}
				
				
			}
		} else {
			toWriteLog = '===================================\r\n';
			toWriteLog += toWriteLog + js_yyyymmdd_hhmmss() + ' = SQL downloaded file is incomplete. Affected file will be deleted.' + '\r\n';
			toWriteLog += toWriteLog + actualInputFile + '\r\n';
			toWriteLog += toWriteLog + '===================================\r\n';
			createLogFile(objStartFolder, toWriteLog);
			if (showEcho == 'showEcho') {
				WScript.echo(toWriteLog);
			}
			customFileFolder.deleteFile(actualInputFile);
		}
	}

}