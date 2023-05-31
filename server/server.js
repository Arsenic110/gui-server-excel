const http = require("http");
const fs = require("fs").promises;
let config;
config = reloadJSON("server/config.json");

const Excel = require("exceljs");
const { Console } = require("console");
const { env } = require("process");
const { Condition } = require("selenium-webdriver");


const server = http.createServer(requestListener);
const __rootname = __dirname.split("\\").slice(0, -1).join("\\"); //lol

var _socket;

const io = require("socket.io")(server);

init();

function init()
{
    //server.listen(config.port, config.hostname, () => 
    {
        console.log(`Server is listening for http traffic on http://${config.hostname}:${config.port}/`);
    }//);

    //create mappings for socket.io

    io.on('connection', (socket) => {
        console.log("Connection established with " + socket.id);
        _socket = socket;

        socket.on("client-handshake", () => 
        {
            console.log(`Completed handshake with ${socket.id}`);
            socket.emit("server-response", "placeholder report");
        });

        socket.on("SEMS-export", () =>
        {
            SEMSExport();
        });

        socket.on("Claims-data", () =>
        {
            ClaimsData();
        });

    });

    //fs.mkdir("OPERATIONS/").then(initFilestructure).catch(() => console.log("OPERATIONS/ already exists, skipping initFilestructure()"));
    initFilestructure();

    //Put a method here to test it on boot - 
    ClaimsData();
}



function requestListener(req, res)
{
    var uri = translateURI(req.url);

    
    console.log(`req: ${req.url}`);

    fs.readFile(uri).then((contents) =>
    {
        res.writeHead(200);

        res.end(contents);
    }).catch((e) => {console.warn(e); res.end("500")});
}

function reloadJSON(path)
{
    path = path == undefined ? "server/config.json" : path;

    let temp = JSON.parse(require("fs").readFileSync(path, 'utf8'));
    config = temp;
    return temp;
}

function translateURI(url)
{
    var parsedUrl = __rootname + "\\public" + url;

    if(url == "/")
        return __rootname + "\\public\\index.html";
    
    return parsedUrl;
}

function initFilestructure()
{
    var folderList = ["SEMS Export", "Half 8", "Return Tracker"];

    folderList.forEach((child) => { fs.mkdir(`OPERATIONS/${child}/`).catch(()=>{}) });
}



















async function SEMSExport()
{
    //here, we will be combining multiple workbooks together into one new one. 
    const outputWB = new Excel.Workbook(); //create a blank WB we will be using as output.
    outputWB.addWorksheet("DATA");
    outputWB.addWorksheet("SEMS");

    reloadJSON();

    //collect the names of all workbooks that do not contain "SEMS" in the filename.

    var fileListRaw = await fs.readdir("OPERATIONS/SEMS Export/"), fileList = [];
    
    for(let i = 0; i < fileListRaw.length; i++)
    {
        if(!fileListRaw[i].toLowerCase().includes("sems"))
            fileList.push("OPERATIONS/SEMS Export/" + fileListRaw[i]);
    }

    if(fileList.length == 0)
    {
        try
        {
            _socket.emit("broadcast", `There are no valid files to operate on.`);
        }
        catch (err)
        {
            console.log("Failed Broadcast");
        }
        return;
    }

    try
    {
        _socket.emit("broadcast", "Starting export...");
    }
    catch (err)
    {
        console.log("Failed Broadcast");
    }


    var wbList = [], userCol, ageCol;

    for(let i = 0; i < fileList.length; i++)
    {
        wbList.push(new Excel.Workbook());
        await wbList[wbList.length -1].xlsx.readFile(fileList[i]);
    }

    //at this point in the program, all our sheets are loaded in. We can start the actual counting and import/export.

    for(let i = 0; i < wbList.length; i++)
    {//loop thru all wbs
        for(let row = 1; row <= wbList[i].worksheets[0].lastRow._number ; row++)
        {
            //this is surprisingly hell
            if(i > 0 && row == 1)
                continue;

            var cells = [];
            for(let c = 0; c < wbList[i].worksheets[0].getRow(row)._cells.length; c++)
            {
                var v = wbList[i].worksheets[0].getRow(row)._cells[c].value;
                cells.push(v);
                if (v == "Assigned to User")
                    userCol = c + 1;
                if(v == "Action age")
                    ageCol = c + 1;
            }

            outputWB.getWorksheet("SEMS").addRow(cells);
        }
    }

    var totalCount = 0, unassignedCount = 0, oldestAge = 0;
    
    //1 for header, and 1 for zero based indexing.
    totalCount = outputWB.getWorksheet("SEMS").getColumn(userCol).values.length - 2;

    //for each unassigned cell
    outputWB.getWorksheet("SEMS").getColumn(userCol).eachCell((cell, no) => {if (cell.value == "")
    { 
        unassignedCount++; 

        var ageCell = outputWB.getWorksheet("SEMS").getColumn(ageCol).values[no];

        var cVal = Number(ageCell.replace(":", "").replace(":", ""));
        oldestAge = Math.max(cVal, oldestAge);
        
    }});

    //copy the header
    outputWB.getWorksheet("DATA").addRow(outputWB.getWorksheet("SEMS").getRow(1).values);

    //clone the unassigned records to the DATA sheet
    outputWB.getWorksheet("SEMS").getColumn(userCol).eachCell((cell, no) => 
    {
        if(cell.value == "")
        {
            outputWB.getWorksheet("DATA").addRow(outputWB.getWorksheet("SEMS").getRow(no).values);
        }
    });

    //we will need to assign names from the list
    let rr = [" "];
    for(let i = 2; i <= outputWB.getWorksheet("DATA").lastRow._number; i++)
    {//uhh this should iterate through the rows here??
        rr.push(config.SEMSExport.names[(i - 2) % config.SEMSExport.names.length]);
        outputWB.getWorksheet("DATA").getColumn(userCol).values = rr;
    }

    //copy the header again cuz why not?(kidding, this fixes the header.)
    outputWB.getWorksheet("DATA").getRow(1).values = outputWB.getWorksheet("SEMS").getRow(1).values;

    let oldestAgeString = "";


    //this will 'fix' date formatting such that the output is actually human readable.
    let minutes = oldestAge % 100;
    let hours = oldestAge % 10000 - minutes == 0 ? "00" : oldestAge % 10000 - minutes;
    let days = oldestAge - hours - minutes == 0 ? "00" : oldestAge - hours - minutes;

    hours = hours == "00" ? "00" : hours / 100; 
    days = days == "00" ? "00" : days / 10000;

    oldestAgeString = `${days}:${hours}:${minutes}`;

    //debug log
    console.log(`Total: ${totalCount}, Unassigned: ${unassignedCount}, Age: ${oldestAgeString}`);

    let today = new Date();

    //write the file!
    await outputWB.xlsx.writeFile(`OPERATIONS/SEMS Export/SEMS ${`${today.getDate()}`.padStart(2, "0")}.${`${today.getMonth() + 1}`.padStart(2, "0")}.xlsx`);

    try
    {
        _socket.emit("broadcast", `Good morning! Total SEMS: ${totalCount}, Unassigned: ${unassignedCount}, Oldest Unassigned Age: ${oldestAgeString}`);
    }
    catch (err)
    {
        console.log("Failed Broadcast");
    }
}


















async function ClaimsData()
{
    //create a new workbook
    const outputWB = new Excel.Workbook();
    outputWB.addWorksheet("Sheet1");

    reloadJSON();

    var fileListRaw = await fs.readdir("OPERATIONS/Claim Data/"), fileList = [];
    
    for(let i = 0; i < fileListRaw.length; i++)
    {
        if(fileListRaw[i].toLowerCase().includes("claims") && fileListRaw[i].toLowerCase().includes("internal"))
            fileList.push("OPERATIONS/Claim Data/" + fileListRaw[i]);
    }

    if(fileList.length == 0)
    {
        Broadcast("There are no valid files to operate on.");
        return;
    }
    else if(fileList.length != 1)
    {
        Broadcast("Too many files detected. Please clear input!");
        return;
    }

    Broadcast("Starting export...");

    var claimsInputFilePath = fileList[0];

    const claimsInputFile = new Excel.Workbook();
    await claimsInputFile.xlsx.readFile(claimsInputFilePath);

    //at this point, we can assume the entire claims file has been loaded into memory
    //Step 1.1: find the correct sheet

    let validClaimsSheet = null;

    claimsInputFile.worksheets.forEach((sheet) => 
    { 
        if(sheet.name.toLowerCase().includes("valid claims") && sheet.name.toLowerCase().includes(config.ClaimsData.currentQuarter.toLowerCase()))
        {
            validClaimsSheet = sheet;
        }
    });
    if(!validClaimsSheet)
    {
        Broadcast("The correct sheet could not be found.");
        return;
    }

    console.log("Found sheet: " + validClaimsSheet.name);

    //Step 1.2: Read all the green rows
    //so green in this case can also be any order where the column AU or AV has not been filled in.

    var claimLetterDateCol = validClaimsSheet.getColumn('AU');
    var rowsToClaim = [];

    claimLetterDateCol.eachCell({includeEmpty: true}, (cell, rowNumber) => 
    {
        if(!cell.value)
        {
            rowsToClaim.push(validClaimsSheet.getRow(rowNumber));
        }

    });

    //last part is that we need to actually copy the headings on the template sheet.
    let templateSheet = new Excel.Workbook();
    await templateSheet.xlsx.readFile("OPERATIONS/Claim Data/template/claim letter data template.xlsx");

    templateSheet = templateSheet.worksheets[0]; //what do you want from me?

    templateSheet.eachRow({includeEmpty: true}, (row, i) => 
    {
        if(i == 1)
            rowsToClaim.unshift(row);
    });


    //Step 1.3: Copy them to new Excel sheet
    const outputSheet = outputWB.getWorksheet("Sheet1");

    

    for(let i = 1; i <= rowsToClaim.length; i++)
    {
        let row = outputSheet.getRow(i);
        row.values = rowsToClaim[i-1].values;
        row.height = rowsToClaim[i-1].height + 5;

        row.eachCell((cel, col) => { cel.style = rowsToClaim[i-1].getCell(col).style; });

        row.commit();
    }

    outputSheet.eachRow({includeEmpty: true}, (row, i) =>
    {
        row.values = rowsToClaim[i-1].values;
        row.height = rowsToClaim[i-1].height + 5;

        row.eachCell({includeEmpty: true}, (cel, col) => { cel.style = rowsToClaim[i-1].getCell(col).style; });

        row.commit();
    });

    await outputWB.xlsx.writeFile(`OPERATIONS/Claim Data/output.xlsx`);

    //process.exit(1);
}

function Broadcast(text2br)
{
    try
    {
        _socket.emit("broadcast", text2br);
    }
    catch (err)
    {
        console.log(`Failed Broadcast: ${text2br}`);
    }
    return;
}