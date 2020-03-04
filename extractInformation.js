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
    if(listElement.options[listElement.selectedIndex].value === 28)
    {
        var date = new Date().getFullYear;
        if(date % 4 === 0)
        {
            selectedMonthDays = 29;
        }
    }
    selectedMonthDays =  listElement.options[listElement.selectedIndex].value
    
    if(items != undefined){
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
                wp: rowData[12]
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
                    days: new Array(selectedMonthDays)
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

        for(var j = 1; j <= selectedMonthDays; j++)
        {
            var checkbox = document.createElement('input');
            checkbox.style = 'margin-left: 2px;'
            checkbox.type = "checkbox";
            checkbox.value = j;
            checkbox.checked = true;
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
    console.log(checkboxes);
    for(var i = 0; i < checkboxes.length; i++)
    {
        checkboxes[i].addEventListener("change", function(){
            if(!this.checked)   //if unchecked, person was not present that day.
            {
            personArr[Number(this.textContent)].days[Number(this.value) - 1] = false;
            }
        })
    }
}


function generateOutput(items) 
{
    var groupByProject = items.reduce((r,a) => {
        r[a.projectName] = [...r[a.projectName] || [], a]
        return r
    },[])
    
    for(var project in groupByProject)
    {
        if(groupByProject.hasOwnProperty(project))
        {
            var groupByAssigneeForProject = groupByProject[project].reduce((r,a) => {
                r[a.assignee] = [...r[a.assignee] || [], a]
                return r 
            }, [])
            for(var assignee in groupByAssigneeForProject)
            {
                if(groupByAssigneeForProject.hasOwnProperty(assignee))
                {
                    var groupByWorkPackageForAssignee = groupByAssigneeForProject[assignee].reduce((r,a) => {
                        r[a.wp] = [...r[a.wp] || [], a];
                        return r;
                    })

                    for(var workpackage in groupByWorkPackageForAssignee)
                    {
                        if(groupByWorkPackageForAssignee.hasOwnProperty(workpackage))
                        {
                            //write to excel now!
                        }
                    }
                }
            }
        }
    }


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

