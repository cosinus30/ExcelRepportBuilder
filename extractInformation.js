const Excel = require("exceljs")

var files = [];
var data = "";
var dropZone = document.getElementById('dropZone');
var dropZoneText = document.getElementById('dropZoneText');
var dropZoneFile = document.getElementById('dropZoneFile');
var fileName = document.getElementById('fileName');
var errorText = document.getElementById('errorText');
var fileInfo = document.getElementById('fileInfo');
var selectFile = document.getElementById('selectFile');
var dynamic = document.getElementById('dynamic');
var donwloadRow = document.getElementById('downloadRow');
var unassigned = document.getElementById('unassigned');
var listElement = document.getElementById('list');
var continueButton = document.getElementById('continue-button');
var fileNumber = 0;
var selectedMonthDays = 31;
var items = [];
var personArr = [];


selectFile.addEventListener('change', handleFileSelect, false);

listElement.addEventListener("change",function(event){
    //case for 29 dayed FEB.
    if(parseInt(listElement.options[listElement.selectedIndex].value) === 28)
    {
        var date = parseInt(new Date().getFullYear());
        
        if(date % 4 == 0)
        {
            selectedMonthDays = 29;
        }
    }
    else{
        selectedMonthDays =  parseInt(listElement.options[listElement.selectedIndex].value);
    }
    
    if(items.length > 0){
        dynamic.innerHTML = "";
        findNumberOfPeople(items);
    }
})

dropZone.addEventListener('dragover', function (e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
});

dropZone.addEventListener('drop', function (e) {
    e.stopPropagation();
    e.preventDefault();
    errorText.innerText = "";
    files = e.dataTransfer.files;
    extractInformationFromFile();
});

continueButton.addEventListener("click", function(e){
    generateOutput(items);
})



function handleFileSelect(e) {
    files = e.target.files;
    extractInformationFromFile();
}

function extractInformationFromFile() {
    if (files.length > 0) {
        for (var i = 0; i < files.length; i++) {
            if (getFileExtension(files[i].name) != "csv") {
                files = [];
                errorText.innerHTML = "You can only upload .csv files";
                return;
            }
        }
    }

    errorText.style.display = "none";
    var output = [];
    dropZoneText.style.display = 'none';
    dropZoneFile.style.display = 'inline-block';

    for (var i = 0, file; file = files[i]; i++) {

        var reader = new FileReader();
        reader.onload = (function (theFile) {
            return function (e) {
                var base64Str = e.target.result.substring(e.target.result.indexOf(",") + 1);
                var resultString = b64DecodeUnicode(base64Str);
                prepareData(resultString, theFile.name);
            };
        })(file);

        reader.readAsDataURL(file);
    }
}

function prepareData(resultString, fileName) {

    var rows = resultString.split('\n');

    for (i = 1; i < rows.length; i++) {
        var startIndex = rows[i].indexOf(",\"");
        var endIndex = rows[i].indexOf("\",");

        if (startIndex > 0 && endIndex > 0) {
            for (j = startIndex + 1; j < endIndex - 1; j++) {
                if (rows[i].charAt(j) == ",") {
                    rows[i] = rows[i].replaceAt(j, " ");
                }
            }
        }

        var rowData = rows[i].split(',');

        if (rowData.length == 14) {
            item = {
                sprintName: fileName,
                issueKey: rowData[0],
                issueId: rowData[1],
                summary: rowData[2],
                assignee: isEmpty(rowData[3]) ? "Unassigned" : toTitleCase(rowData[3].split('.').join(' ')),
                status: rowData[4],
                projectKey: rowData[5],
                projectName: rowData[6],
                projectType: rowData[7],
                projectLead: rowData[8],
                projectDescription: rowData[9],
                projectUrl: rowData[10],
                storyPoints: rowData[11],
                wp: rowData[12],
                timeSpent: rowData[13]
            };
            items.push(item);
        }
    }
    fileNumber = fileNumber + 1;
    if (fileNumber == files.length) {
        findNumberOfPeople(items);
    }
}

function findNumberOfPeople(items){
    personArr = [];
    var nameArr = [];
    for(var i = 0; i < items.length; i++){
        if(personArr != undefined && nameArr.indexOf(items[i].assignee) == -1)
        {
            personArr.push(
                {
                    name: items[i].assignee,
                    days: new Array(selectedMonthDays).fill(true)
                }
                )
            nameArr.push(items[i].assignee)
        }
    }
    renderDaysAndButton()
}


function renderDaysAndButton()
{
    var numberOfPeople = personArr.length;
    for(var i = 0; i < numberOfPeople; i++)
    {
        var divPerson = document.createElement("div");
        divPerson.className = "col-sm-12";
        divPerson.style = "padding: 12px;"

        var title = document.createElement("h3");
        title.innerText = personArr[i].name;
        divPerson.appendChild(title);

        var divCheckBox = document.createElement("div");
        divCheckBox.className = "checkbox";
        divCheckBox.style = "margin-top: 20px; border: 1px solid lightblue; padding: 2px;"
        divPerson.appendChild(divCheckBox);

        dynamic.appendChild(divPerson);
        var date = new Date();
        for(var j = 1; j <= selectedMonthDays; j++)
        {
            var checkbox = document.createElement('input');
            checkbox.style = 'margin-left: 2px;'
            checkbox.type = "checkbox";
            checkbox.value = j;
            var curDate = new Date(date.getFullYear(),listElement.selectedIndex, j);
            //Make saturdays and sundays unchecked.
            if(curDate.getDay() == 0 || curDate.getDay() == 6)
            {
                checkbox.checked = false;
            }
            else
            {
                checkbox.checked = true;
            }
            checkbox.textContent = i;
            var id = j + "_" + i;

            var newlabel = document.createElement("label");
            newlabel.setAttribute("for",id);
            newlabel.innerHTML = j;

            divCheckBox.appendChild(newlabel);
            divCheckBox.appendChild(checkbox);

            
        }
    }
    addListeners();
    continueButton.style = "align-items: flex-end; justify-content: center; margin-top: 12px; display:inline-block;"
}

function addListeners()
{
    var checkboxes = document.querySelectorAll("input[type=checkbox]");
    for(var i = 0; i < checkboxes.length; i++)
    {
        if(checkboxes[i].checked === false)
        {
            personArr[Number(checkboxes[i].textContent)].days[Number(checkboxes[i].value) - 1] = false;
        }

        checkboxes[i].addEventListener("change", function(){
            if(!this.checked)   //if unchecked, person was not present that day.
            {
                personArr[Number(this.textContent)].days[Number(this.value) - 1] = false;
            }
        })
    }
}


async function generateOutput(items) 
{
    var groupByProject = await items.reduce((r,a) => {
        r[a.projectName] = [...r[a.projectName] || [], a]
        return r
    },[])

    var defaultWorkbook = new Excel.Workbook();
    var defaultWorksheet;
    defaultWorkbook.xlsx.readFile("excelFiles/rawFiles/H2020_timesheet_template.xlsx").then(async function(){
        defaultWorksheet = defaultWorkbook.getWorksheet(1);

        for(var project in groupByProject)
        {
            if(groupByProject.hasOwnProperty(project))
            {
                var groupByAssigneeForProject = await groupByProject[project].reduce((r,a) => {
                    r[a.assignee] = [...r[a.assignee] || [], a]
                    return r 
                }, [])
                for(var assignee in groupByAssigneeForProject)
                {
                    var workbook = new Excel.Workbook;
                    var sheet = workbook.addWorksheet();
                    defaultWorksheet.eachRow((row, rowNumber) => {
                        var newRow = sheet.getRow(rowNumber);
                        row.eachCell((cell, colNumber) => {
                            var newCell = newRow.getCell(colNumber)
                            for(var prop in cell){
                                newCell[prop] = cell[prop];
                            }
                        })
                    })            
                    if(groupByAssigneeForProject.hasOwnProperty(assignee))
                    {
                        var groupByWorkPackageForAssignee = await groupByAssigneeForProject[assignee].reduce((r,a) => {
                            r[a.wp] = [...r[a.wp] || [], a];
                            return r;
                        }, []);
                        var rowNumberForWorkPackage = 6;
                        
                        for(var workpackage in groupByWorkPackageForAssignee)
                        {
                            if(groupByWorkPackageForAssignee.hasOwnProperty(workpackage))
                            {
                                var totalTimeSpentOnWP = 0.0;
                                for (var i = 0; i < groupByWorkPackageForAssignee[workpackage].length; i++)
                                {
                                    if(groupByWorkPackageForAssignee[workpackage][i].status === "Done")
                                    {
                                        if(groupByWorkPackageForAssignee[workpackage][i].timeSpent != "" && groupByWorkPackageForAssignee[workpackage][i].timeSpent != null && groupByWorkPackageForAssignee[workpackage][i].timeSpent != undefined)
                                        {
                                            if (parseInt(groupByWorkPackageForAssignee[workpackage][i].timeSpent) > 0)   {
                                                totalTimeSpentOnWP += parseInt(groupByWorkPackageForAssignee[workpackage][i].timeSpent);
                                            }
                                        }
                                    }
                                }
                                totalTimeSpentOnWP = totalTimeSpentOnWP / 3600;
                                var totalTimeSpentPerDay = [];

                                if(totalTimeSpentOnWP > 0)
                                {
                                    var sheetRow = sheet.getRow(rowNumberForWorkPackage);
                                    sheet.getCell(rowNumberForWorkPackage,1).text = workpackage;
                                    
                                    for(var t = 0; t < personArr.length; t++){
                                        if(personArr[t].name === assignee){
                                            for(var n = 0; n < personArr[t].days.length; n++){
                                                //last element has been reached.
                                                if(n == personArr[t].days.length - 1){
                                                    while (totalTimeSpentOnWP > 0) {
                                                        var hours = Math.round(Math.random()*n);

                                                        if(hours < 5)
                                                        {
                                                            totalTimeSpentPerDay.push(hours);
                                                            totalTimeSpentOnWP -= hours;
                                                        }
                                                        
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    for(var t = 0; t < personArr.length; t++){
                                        if(personArr[t].name === assignee){
                                            for(var n = 0; n < personArr[t].days.length; n++){
                                                //burda da arrayden okuyup islem yap.
                                                if(personArr[t].days[n] != false){
                                                    
                                                    var columnSum = 0;
                                                    sheet.getColumn(n+2).eachCell((cell,rowNumber) => {
                                                        if(rowNumber >= 6 && rowNumber < rowNumberForWorkPackage){
                                                            columnSum += cell.value;
                                                        }
                                                    });
                                                    if(columnSum + totalTimeSpentPerDay[totalTimeSpentPerDay.length - 1] > 6)
                                                    {
                                                        sheetRow.getCell(n+2).value = 0;
                                                    }
                                                    else
                                                    {
                                                        sheetRow.getCell(n+2).value = totalTimeSpentPerDay.pop();
                                                    }
                                                    
                                                }
                                            }
                                        }
                                    }
                                    console.log(rowNumberForWorkPackage + " " + assignee + " " + project + " " + workpackage);
                                    rowNumberForWorkPackage++;
                                    totalTimeSpentOnWP = 0;
                                }
                            }
                        }
                    }
                    workbook.xlsx.writeFile("excelFiles/processedFiles/" + project.toString().substring(0,project.toString().indexOf("/")) + project.toString().substring(project.toString().indexOf("/") + 1) + "_" + assignee + ".xlsx").then(function(){});
                }
            }
        }
    });
    alert("Documents have been created.")
}

function clearFile() {
    dropZoneText.style.display = 'inline-block';
    dropZoneFile.style.display = 'none';
    files = [];
    data = "";
    fileName.innerText = "";
    errorText.innerText = "";
    fileInfo.innerHTML = "";
    dynamic.innerHTML = "";
    donwloadRow.innerHTML = "";
    unassigned.innerHTML = "";
    continueButton.style = "display:none" ;
}

// helper functions
String.prototype.replaceAt = function (index, replacement) {
    return this.substr(0, index) + replacement + this.substr(index + replacement.length);
}

function getFileExtension(fileName) {
    return fileName.slice((Math.max(0, fileName.lastIndexOf(".")) || Infinity) + 1);
}

function toTitleCase(str) {
    return str.replace(
        /\w\S*/g,
        function (txt) {
            return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        }
    );
}

function b64DecodeUnicode(str) {
    return decodeURIComponent(atob(str).split('').map(function (c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));
}

function isEmpty(value) {
    var answer = value === undefined || value === null || value === "";
    return answer;
}

