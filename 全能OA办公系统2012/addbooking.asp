<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>�齭�а칫ϵͳ</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
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
<center>
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                <td class="style7">������Դ</td>
              </tr>
          </table></td>
          <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
        </tr>
      </table>
      <font color="0D79B3"></font></div></td>
  </tr>
</table>
<table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="1%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
      <div align="center">��ԴԤԼ
        </center>

      </div>
      <center>
<%
if request("submit")="ԤԼ" then
starttime=request("starttime")
endtime=request("endtime")
username=oabusyusername
equipment=request("equipment")
purpose=request("purpose")
if isdate(starttime) and isdate(endtime) then
'�����ݿ⣬�ж�ԤԼʱ���Ƿ��ͻ
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from booking where equipment=" & sqlstr(equipment)
rs.open sql,conn,1
bookingallow="yes"
'response.write "bookingallow=" & bookingallow & "<br>"
while not rs.eof and not rs.bof
if (cdate(starttime)>cdate(rs("starttime")) and cdate(starttime)<cdate(rs("endtime"))) or (cdate(endtime)>cdate(rs("starttime")) and cdate(endtime)<cdate(rs("endtime"))) or (cdate(starttime)<cdate(rs("starttime")) and cdate(endtime)>cdate(rs("endtime"))) or (cdate(starttime)>=cdate(rs("starttime")) and cdate(endtime)<=cdate(rs("endtime"))) then bookingallow="no"
rs.movenext
wend
if bookingallow="no" then
%>
<br><br>��ԤԼ��ʱ����Ѿ���ռ�ã�<br><br>
<input type="button" value="����" onclick="window.location.href='addbooking.asp';">
<%
else
set conn=opendb("oabusy","conn","accessdsn")
sql = "Insert Into booking (username,starttime,endtime,equipment,purpose) Values( "
sql = sql & SqlStr(username) & ", "
sql = sql & "#" & starttime & "#, "
sql = sql & "#" & endtime & "#, "
sql = sql & SqlStr(equipment) & ", "
sql = sql & SqlStr(purpose) & ")"
conn.Execute sql
%>
<br><br>ԤԼ�ɹ���<br><br>
<form action="booking.asp">
<input type="submit" value="����">
</form>
<%
end if
else
%>
<br><br>
��������ڲ���ȷ����ע���С�º����£�<br>
<br>
<input type="button" value="����" onclick="window.location.href='addbooking.asp';">
<%
end if
else
%>

<script Language="JavaScript">

 function checktime(){
   var sy=document.form1.startyear.value;
   var sm=document.form1.startmonth.value;
   var sd=document.form1.startday.value;
   var sh=document.form1.starthour.value;
   var smin=document.form1.startminute.value;
   var ey=document.form1.endyear.value;
   var em=document.form1.endmonth.value;
   var ed=document.form1.endday.value;
   var eh=document.form1.endhour.value;
   var emin=document.form1.endminute.value;
   var stime=sy+"-"+sm+"-"+sd+" "+sh+":"+smin+":00";
   var etime=ey+"-"+em+"-"+ed+" "+eh+":"+emin+":00";
   document.form1.starttime.value=stime;
   document.form1.endtime.value=etime;

   a1=0
   
   if((ey-sy)>0){
            a1=1;
            }
   else{
     if(ey==sy){
         if((em-sm)>0){
                  a1=1;
                  }
         else{
           if(em==sm){
              if((ed-sd)>0){
                      a1=1;
                       }
              else{
                  if(ed==sd){
                      if((eh-sh)>0){
                               a1=1;
                               }
                       else{
                            if(eh==sh){
                                if((emin-smin)>0){
                                             a1=1;
                                              };
                                       };
                           };
                            };
                 };
                    };
             };
               };
          };

   if(a1==0){window.alert("����ʱ��Ӧ���ڿ�ʼʱ��֮ǰ");document.form1.startyear.focus();return (false);}
                    }

</script>
<%
'ȡ�õ�ǰСʱ
myhour=hour(now())
'ȡ�õ�ǰ����
myday=day(now())
'ȡ�õ�ǰ��
mymonth=month(now())
'ȡ�õ�ǰ��
myyear=year(now())
%>
<br>
<form method=post name="form1" action="addbooking.asp">
  <table border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td>ԤԼ��Դ:</td><td>
<select size=1 name="equipment">
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<option value="<%=rs("equipment")%>"><%=rs("equipment")%></option>
<%
rs.movenext
wend
%>
</select></td>
    </tr>
    <tr>
      <td>��ʼʹ��ʱ��:</td><td>
         <select size=1 name="startyear">
<%
for i=2001 to 2009
%>
          <option value="<%=i%>"<%=selected(i,myyear)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="startmonth">
<%
for i=1 to 12
%>
          <option value="<%=i%>"<%=selected(i,mymonth)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="startday">
<%
for i=1 to 31
%>
          <option value="<%=i%>"<%=selected(i,myday)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="starthour">
<%
for i=0 to 23
%>
          <option value="<%=i%>"<%=selected(i,myhour)%>><%=i%></option>
<%
next
%>
        </select>ʱ
        <select size=1 name="startminute">
<%
for i=10 to 50 step 10
%>
          <option value="<%=i%>"><%=i%></option>
<%
next
%>
        </select>��</td>
    </tr>
    <tr>
      <td>����ʹ��ʱ��:</td><td>
       <select size=1 name="endyear">
<%
for i=2001 to 2009
%>
          <option value="<%=i%>"<%=selected(i,myyear)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="endmonth">
<%
for i=1 to 12
%>
          <option value="<%=i%>"<%=selected(i,mymonth)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="endday">
<%
for i=1 to 31
%>
          <option value="<%=i%>"<%=selected(i,myday)%>><%=i%></option>
<%
next
%>
        </select>��
        <select size=1 name="endhour">
<%
for i=0 to 23
%>
          <option value="<%=i%>"<%=selected(i,myhour)%>><%=i%></option>
<%
next
%>
        </select>ʱ
        <select size=1 name="endminute">
<%
for i=10 to 50 step 10
%>
          <option value="<%=i%>"><%=i%></option>
<%
next
%>
        </select>��</td>
    </tr>
    <tr>
      <td colspan="2" align=center>ʹ��˵��</td>
    </tr>
    <tr>
      <td colspan="2"><textarea rows="9" cols="50" name="purpose"></textarea></td>
    </tr>
  </table>
<input type="hidden" name="starttime">
<input type="hidden" name="endtime">
<input type="submit" name="submit" value="ԤԼ" onclick="return checktime();">&nbsp;&nbsp;&nbsp;<input type="button" value="����" onclick="window.location.href='booking.asp'">
</form>
<%
end if
%>
</center></td>
  </tr>
</table>


</body>
</html>












