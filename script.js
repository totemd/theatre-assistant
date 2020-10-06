let XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest
let request = new XMLHttpRequest();
request.open("GET", "../SJDA.json");
request.responseType = "json";
request.send();
request.onload = function(){
    const playText = request.response;
    const play = JSON.parse(playText);
    console.log(play);
}
