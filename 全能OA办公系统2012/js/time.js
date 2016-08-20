<!--//
img = new Array()
for(var i=0; i <= 14; i++) {
img[i] = new Image()
}
img[1].src = "../images/time/dg0.gif"
img[2].src = "../images/time/dg1.gif"
img[3].src = "../images/time/dg2.gif"
img[4].src = "../images/time/dg3.gif"
img[5].src = "../images/time/dg4.gif"
img[6].src = "../images/time/dg5.gif"
img[7].src = "../images/time/dg6.gif"
img[8].src = "../images/time/dg7.gif"
img[9].src = "../images/time/dg8.gif"
img[10].src = "../images/time/dg9.gif"
img[11].src = "../images/time/dgon.gif"
img[12].src = "../images/time/dgoff.gif"
img[13].src = "../images/time/dgam.gif"
img[14].src = "../images/time/dgpm.gif"
var base = "../images/time/dg"
var space = "../images/time/space.gif" 
var per = false

function go() {
per = true
start()
}

function start() {
if(per == true) {
var now = new Date()
var hours = now.getHours();
var ampm = (hours < 12) ? "am" : "pm"
hours = (hours > 12) ? (hours - 12) + "" : hours + ""
hours = (hours == "0") ? "12" : hours
hours = (hours < 10) ? "0" + hours : hours + ""
var minutes = now.getMinutes();
minutes = (minutes < 10) ? "0" + minutes : minutes + ""
var seconds = now.getSeconds();
seconds = (seconds < 10) ? "0" + seconds : seconds + ""
document.one.src = (hours.charAt(0)=="0") ? space : add(hours.charAt(0))
document.two.src = add(hours.charAt(1))
document.three.src = (now.getSeconds() % 2) ? add("on") : add("off")
document.four.src = add(minutes.charAt(0))
document.five.src = add(minutes.charAt(1))
document.six.src = add(ampm)
setTimeout("start()",1000)
}
}

function add(it) {
return base + it + ".gif"
}

//-->
