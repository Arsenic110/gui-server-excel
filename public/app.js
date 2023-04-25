var socket;

var outputP;
var opsDiv, opData = [];

window.onload = init;

function init()
{
    outputP = document.getElementById("console-p");
    opsDiv = document.getElementById("operations");

    socket = io.connect(window.location.href);

    writeOut("Requesting handshake...");
    socket.emit("client-handshake");

    socket.on("server-response", (res) => {writeOut(`Connection Established with ${writeColor(socket.id, "#00FF00")}`)});

    socket.on("job-done", (res) => {writeReply(`Server ${writeColor("completed", "#00FF00")} task ${res}`)});

    Array.from(opsDiv.children).forEach((child) => 
    {
        if(child.tagName.toLowerCase() != "button")
            return;

        opData.push({child : child, id : child.id});
    });

    opData.forEach((element) => 
    {
        element.child.addEventListener("click", function(){ socket.emit(element.id); writeOut(`Emitting ${element.id}`)});
    });

}

function writeOut(str)
{
    if(outputP.innerHTML != "")
        outputP.innerHTML += "<br>";
    outputP.innerHTML += ">" + str;
}

function writeReply(str)
{
    str = `<br><span style="float:right;">${str}</span>`;   
    outputP.innerHTML += str;
}

function writeColor(str, col)
{
    return `<span style="color:${col};">${str}</span>`;
}