var vn="Microsoft Internet Explorer";
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "<font color=ff6600>星期日</font>"
if(today.getDay()==1) day = "星期一"
if(today.getDay()==2) day = "星期二"
if(today.getDay()==3) day = "星期三"
if(today.getDay()==4) day = "星期四"
if(today.getDay()==5) day = "星期五"
if(today.getDay()==6) day = "<font color=ff6600>星期六</font>"
document.fgColor = "000000";
date =  (today.getYear()) + "年" + (today.getMonth() + 1 ) + "月" + today.getDate() + "日 " + day +"";
document.write(date);
