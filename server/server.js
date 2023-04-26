const http = require("http");
const fs = require("fs").promises;
const config = require("./config.json");

const Excel = require("exceljs");


const server = http.createServer(requestListener);
const __rootname = __dirname.split("\\").slice(0, -1).join("\\"); //lol

var _socket;

const io = require("socket.io")(server);

init();

function init()
{
    server.listen(config.port, config.hostname, () => 
    {
        console.log(`Server is listening for http traffic on http://${config.hostname}:${config.port}/`);
    });

    //create mappings for socket.io

    io.on('connection', (socket) => {
        console.log("Connection established with " + socket.id);
        _socket = socket;

        socket.on("client-handshake", () => 
        {
            console.log(`Completed handshake with ${socket.id}`);
            socket.emit("server-response", "placeholder report");
        });

        socket.on("test-op", () => 
        {
            socket.emit("job-done", "placeholder reply to [test-op]");
        });

        socket.on("test-op-2", () => 
        {
            socket.emit("job-done", "placeholder reply to [test-op-2]");
        });

        socket.on("SEMS-export", () =>
        {
            socket.emit("broadcast", "Starting export...");
            SEMSExport();
        });
    });

    //fs.mkdir("OPERATIONS/").then(initFilestructure).catch(() => console.log("OPERATIONS/ already exists, skipping initFilestructure()"));
    initFilestructure();
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
    outputWB.addWorksheet("SEMS");
    //collect the names of all workbooks that do not contain "SEMS" in the filename.

    var fileListRaw = await fs.readdir("OPERATIONS/SEMS Export/"), fileList = [];
    

    for(let i = 0; i < fileListRaw.length; i++)
    {
        if(!fileListRaw[i].toLowerCase().includes("sems"))
            fileList.push("OPERATIONS/SEMS Export/" + fileListRaw[i]);
    }

    //console.log(fileList.join(" "));

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
    
    totalCount = outputWB.getWorksheet("SEMS").getColumn(userCol).values.length - 2;
    outputWB.getWorksheet("SEMS").getColumn(userCol).eachCell((cell, no) => {if (cell.value == "")
    { 
        unassignedCount++; 

        var ageCell = outputWB.getWorksheet("SEMS").getColumn(ageCol).values[no];
        console.log(ageCell);

        var cVal = Number(ageCell.replace(":", "").replace(":", ""));
        oldestAge = Math.max(cVal, oldestAge);
        
    }});
    outputWB.getWorksheet("SEMS").getColumn(ageCol).eachCell((cell, no) => 
    {


    });

    console.log(`Total: ${totalCount}, Unassigned: ${unassignedCount}, Age: ${oldestAge}`);

    await outputWB.xlsx.writeFile("OPERATIONS/SEMS Export/SEMS.xlsx");

    try
    {
        _socket.emit("broadcast", `SEMS Export complete. Total: ${totalCount}, Unassigned: ${unassignedCount}, Oldest Unassigned Age: ${oldestAge}`);
    }
    catch (err)
    {
        console.log("Failed Broadcast");
    }
}