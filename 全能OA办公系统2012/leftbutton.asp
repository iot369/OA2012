<%@ LANGUAGE = VBScript %>
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!-- Ooulook ������ʼ����-->
            <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")

cook_allow_see_all_workrep=request.cookies("cook_allow_see_all_workrep")
cook_allow_see_dept_workrep=request.cookies("cook_allow_see_dept_workrep")
cook_allow_control_dept_user=request.cookies("cook_allow_control_dept_user")
cook_allow_control_all_user=request.cookies("cook_allow_control_all_user")
cook_allow_send_note=request.cookies("cook_allow_send_note")
cook_allow_control_note=request.cookies("cook_allow_control_note")
cook_allow_control_file=request.cookies("cook_allow_control_file")
cook_allow_control_level=request.cookies("cook_allow_control_level")
'�����ݿ⣬����Ȩ��
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
cook_allow_see_all_personinf=rs("allow_see_all_personinf")
cook_allow_see_dept_personinf=rs("allow_see_dept_personinf")
cook_allow_edit_all_jobchanginf=rs("allow_edit_all_jobchanginf")
cook_allow_edit_dept_jobchanginf=rs("allow_edit_dept_jobchanginf")

cook_allow_edit_all_rewpuninf=rs("allow_edit_all_rewpuninf")
cook_allow_edit_dept_rewpuninf=rs("allow_edit_dept_rewpuninf")

cook_allow_see_all_workrep=rs("allow_see_all_workrep")
cook_allow_see_dept_workrep=rs("allow_see_dept_workrep")
cook_allow_edit_all_checkinf=rs("allow_edit_all_checkinf")
cook_allow_edit_dept_checkinf=rs("allow_edit_dept_checkinf")
allow_edit_work_time=rs("allow_edit_work_time")

%>
<html>
<head>
<title>�ޱ����ĵ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css" type="text/css">
<style type="text/css">
  .titleStyle{
       background-image:url(images/button.gif); color:#ffffff;
      font-size:9pt;cursor:hand;
  }
  .contentStyle{
      background-color:#E4E8F3;color:blue;font-size:9pt;
  }
<!--
.style1 {color: #003366}
a:link {
	color: #003366;
	text-decoration: underline;
}
a:visited {
	text-decoration: underline;
	color: #003366;
}
a:hover {
	text-decoration: none;
	color: #FF0000;
}
a:active {
	text-decoration: underline;
	color: #FF0000;
}
body,td,th {
	color: #003366;
}
-->
</style>
</head>

<body style="BACKGROUND-COLOR: transparent" onmouseover="self.status='��ӭ����������������е��OA���ܰ칫�Զ���ϵͳ';return true">
<script language="JavaScript">
<!--
 var layerTop=0;       //�˵����߾�
 var layerLeft=0;      //�˵���߾�
 var layerWidth=153;    //�˵��ܿ�
 var titleHeight=25;    //�������߶�
 var contentHeight=290; //�������߶�
 var stepNo=10;         //�ƶ���������ֵԽ���ƶ�Խ��

 var itemNo=0;runtimes=0;
 document.write('<span id=itemsLayer style="position:absolute;overflow:hidden;left:'+layerLeft+';top:'+layerTop+';width:'+layerWidth+';">');

 function addItem(itemTitle,itemContent){
   itemHTML='<div id=item'+itemNo+' itemIndex='+itemNo+' style="position:relative;left:0;top:'+(-contentHeight*itemNo)+';width:'+layerWidth+';"><table width=100% cellspacing="0" cellpadding="0">'+
       '<tr><td height='+titleHeight+' onclick=changeItem('+itemNo+') class="titleStyle" align=center >'+itemTitle+'</td></tr>'+
       '<tr><td height='+contentHeight+' class="contentStyle" valign=top>'+itemContent+'</td></tr></table></div>';
   document.write(itemHTML);
   itemNo++;
 }
    //��Ӳ˵���������ݣ���������ע���ʽ��
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_gonggongxinxi.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="noticelook.asp" target="main_body">��˾ͨ��</a></div></td></tr><%   if cook_allow_see_dept_workrep="yes" or cook_allow_see_dept_workrep="yes" then   %><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="stafdayrep.asp" target="main_body">�����ƻ�</a></div></td></tr><%  end if  %><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="stafaddressinf.asp" target="main_body">ͨѶ����</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="clientinf.asp" target="main_body">�ͻ���Դ</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="booking.asp" target="main_body">������Դ</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="addbooking.asp" target="main_body">��ԴԤԼ</a></div></td></tr></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_gongwenchuanyue.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="senddate.asp" target="main_body">���͹���</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="havesenddate.asp" target="main_body">�ѷ�����</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="havercidate.asp" target="main_body">���չ���</a></div></td></tr><%     if cook_allow_control_file="yes" then     %><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="senddatecontrol.asp" target="main_body">���Ĺ���</a></div></td></tr><%end if%></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_gerenbangong.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="dayrep.asp" target="main_body">���˹����ƻ�</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="personinf.asp" target="main_body">���˻�������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="personlist.asp" target="main_body">����ͨѶ¼</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="userinf.asp" target="main_body">��������ά��</a></div></td></tr></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_kaoqin.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="kq/kqframe.asp" target="main_body">��ʼ����</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="kq/nowkqinfo.asp" target="main_body">���տ�������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="kq/daykqinfo.asp" target="main_body">�տ���ͳ��</a></div></td></tr><tr>  <td height="23" bgcolor="F8FCFF"><div align="center"><a href="kq/monthkqinfo.asp" target="main_body">�¿���ͳ��</a></div></td></tr>  <%	if allow_edit_work_time="yes" then%><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="kq/settime.asp" target="main_body">���ÿ���ʱ��</a></div></td></tr>    <%	end if%></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_xiaoshouxitong.gif" width="153" height="25"></td></tr></table>','<center><table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><%   if cook_allow_see_dept_workrep="yes" or cook_allow_see_dept_workrep="yes" then   %><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/system_level.asp" target="main_body">ϵͳ����</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/manageUser.asp" target="main_body">�û�����</a></div></td></tr><%  end if  %><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/addData.asp" target="main_body">��������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/advanceSearch.asp" target="main_body">��������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/dataForm.asp" target="main_body">���ݱ���</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/exportData.asp" target="main_body">���ݵ���</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/transData.asp" target="main_body">ת������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/listAll.asp" target="main_body">��������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center" class="style1"><a href="crm/logout.asp" target="main_body">ע��ϵͳ</a></div></td></tr></table></center>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_hetongguanli.gif" width="153" height="25"></td></tr></table>','<center><table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><%	if allow_edit_work_time="yes" then%><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="hetong/htlist.asp?cmd=resetall" target="main_body">��ͬ�б�</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="hetong/htadd.asp" target="main_body">��Ӻ�ͬ</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="hetong/htsrch.asp" target="main_body">�߼���ѯ</a></div></td></tr><%end if%></table></center>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_yuangongguanli.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><%	if allow_edit_work_time="yes" then%><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="stafpersoninf.asp" target="main_body">Ա����������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="jobchanginf.asp" target="main_body">Ա��ְλ�䶯</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="rewpuninf.asp" target="main_body">Ա���������</a></div></td></tr><tr>  <td height="23" bgcolor="F8FCFF"><div align="center"><a href="checkinf.asp" target="main_body">Ա���������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="wageinf.asp" target="main_body">Ա�����ʵ���</a></div></td></tr>    <%	end if%></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_gerenyouxiang.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="email/sendemail.asp" target="main_body">���ʼ�</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="email/getnewemail.asp" target="main_body">δ���ʼ�</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="email/getemailbox.asp" target="main_body">�ռ���</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="email/sendemailbox.asp" target="main_body">�ļ���</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="email/delemailbox.asp" target="main_body">����Ͱ</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="mail.asp" target="main_body">��������</a></div></td></tr></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_gonggongfuwu.gif" width="153" height="25"></td></tr></table>','<table width=96% border="0" align=center cellpadding=0 cellspacing=1 bgcolor="E4E8F3"><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="upfile/default.asp" target="main_body">����Ӳ��</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="chat/index.asp" target="main_body">�������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="vote/default.asp" target="main_body">�������</a></div></td></tr><tr><td height="23" bgcolor="F8FCFF"><div align="center"><a href="bbsnew/index.asp" target="main_body">������̳</a></div></td></tr></table>');
 addItem('<table width="153" height="25"  border="0" cellpadding="0" cellspacing="0"><tr><td height="2"><img src="images/button/bt_yonghushezhi.gif" width="153" height="25"></td></tr></table>','<table cellpadding=0 cellspacing=0 align=center width=100%><%if cook_allow_control_dept_user="yes" then%><tr><td height="18"><div align="center"><a href="addstaf.asp" target="main_body">���������û�</a></div></td></tr><tr><td height="18"><div align="center"><a href="stafcontrol.asp" target="main_body">���������û�</a></div></td></tr><%end if%><%if cook_allow_control_all_user="yes" then%><tr><td height="18"><div align="center"><a href="adduser.asp" target="main_body">�����û�</a></div></td> </tr><tr><td height="18"><div align="center"><a href="usercontrol.asp" target="main_body">�����û�</a></div></td></tr><tr><td height="18"><div align="center"><a href="companymanager.asp" target="main_body">��λ����ά��</a></div></td></tr><% end if %><%if cook_allow_control_level="yes" then %><tr><td height="18"><div align="center"><a href="usercontrolpopedom.asp" target="main_body">�û�����Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="workplanpopedom.asp" target="main_body">�����ƻ�Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="noticefilepopedom.asp" target="main_body">ͨ�湫��Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="jobchanginfpopedom.asp" target="main_body">ְ��䶯Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="rewpunpopedom.asp" target="main_body">���ͱ༭Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="checkinfpopedom.asp" target="main_body">���˱༭Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="wageinfpopedom.asp" target="main_body">���ʱ༭Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="addressinfpopedom.asp" target="main_body">ͨѶ����Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="resourcesetting.asp" target="main_body">��Դ����Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="userkqmanager.asp" target="main_body">���ڹ���Ȩ��</a></div></td></tr><tr><td height="18"><div align="center"><a href="personinfpopedom.asp" target="main_body">��������Ȩ��</a></div></td></tr><% end if%></table>');

 document.write('</span>')
 document.all.itemsLayer.style.height=itemNo*titleHeight+contentHeight;

 toItemIndex=itemNo-1;onItemIndex=itemNo-1;

 function changeItem(clickItemIndex){
   toItemIndex=clickItemIndex;
   if(toItemIndex-onItemIndex>0) moveUp(); else moveDown();
   runtimes++;
   if(runtimes>=stepNo){
     onItemIndex=toItemIndex;
     runtimes=0;}
   else
     setTimeout("changeItem(toItemIndex)",10);
 }

 function moveUp(){
   for(i=onItemIndex+1;i<=toItemIndex;i++)
     eval('document.all.item'+i+'.style.top=parseInt(document.all.item'+i+'.style.top)-contentHeight/stepNo;');
 }

 function moveDown(){
   for(i=onItemIndex;i>toItemIndex;i--)
     eval('document.all.item'+i+'.style.top=parseInt(document.all.item'+i+'.style.top)+contentHeight/stepNo;');
 }
 changeItem(0);
//-->
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Hide
function killErrors() {
return true;
}
window.onerror = killErrors;
// -->
</SCRIPT>
<SCRIPT language=javascript>
<!--
if (window.Event) 
��document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
��if (e.which == 2 || e.which == 3)
�� return false;
 } else if (event.button == 2 || event.button == 3) {
�� event.cancelBubble = true
�� event.returnvalue = false;
�� return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;��// for IE5+
document.onmousedown = norightclick;���� // for all others
//-->
</SCRIPT>

<script language="javascript">
//����"ע��"����ʱ�������Ի����Ƿ�Ҫ���˳�ϵͳ
function closesystem()
{
	window.open('logout.asp?closeflag=0','closesystem','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
	window.location.href="default.asp";
}
</script>
</body>
</html>
