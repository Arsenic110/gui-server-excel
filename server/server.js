const http = require("http");
const fs = require("fs").promises;
const config = require("./config.json");


const server = http.createServer(requestListener);
const __rootname = __dirname.split("\\").slice(0, -1).join("\\");

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

        socket.on("hello", () => console.log(`Completed handshake with ${socket.id}`));
    })
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