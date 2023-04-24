var socket;

window.onload = init;

function init()
{
    socket = io.connect(window.location.href);

    socket.emit("hello");
}