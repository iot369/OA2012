var vn="Microsoft Internet Explorer";
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "<font color=ff6600>������</font>"
if(today.getDay()==1) day = "����һ"
if(today.getDay()==2) day = "���ڶ�"
if(today.getDay()==3) day = "������"
if(today.getDay()==4) day = "������"
if(today.getDay()==5) day = "������"
if(today.getDay()==6) day = "<font color=ff6600>������</font>"
document.fgColor = "000000";
date =  (today.getYear()) + "��" + (today.getMonth() + 1 ) + "��" + today.getDate() + "�� " + day +"";
document.write(date);
