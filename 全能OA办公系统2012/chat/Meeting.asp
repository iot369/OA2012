<HTML>
<HEAD>
<TITLE>电子会议!</TITLE>
<style>
<!--
body {font-family:verdana,arial; font-size:14px; }
td {font-family:verdana,arial; font-size:14px; }
input.text {border:1px solid black; font-size:14px; background:#fcf0fc}
td.little {font-size:12px}
-->
</style>
</HEAD>
<CENTER>
<BODY style='margin:0;background:#F9F9FF'>





<table width="702"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="21" background="../images/r_bg.gif"><div align="center"><font color="0D79B3"><b>— 公 共 服 务 —</b></font></div></td>
  </tr>
</table>
<br><br><br><br>


<%
const ItemNum=20

nShowIP=Request.Form("ShowIp")
strName=Request.Form("Name")
strComment=Request.Form("comment")
strMeeting=Request.Form("Meeting")
strPassword=Request.Form("Password")
strDate=MID(Now(),1,8)

strMeetingID = strMeeting & strPassword & strDate & "Meet"
strTotalCount = strMeeting & strPassword & strDate & "Count"

strIPID = strMeeting & strPassword & strDate & "ID"
strIP = "<OPTION>" & Request.ServerVariables("REMOTE_ADDR") & "-" & CStr(Now()& "<BR>")
Application.Lock()
Application(strIPID)= Application(strIPID)& strIP
Application.UnLock()


If Len(strComment) > 0 Then

if strName="" then strName="NONAME"

  Application.Lock()
  If Application(strTotalCount) = 0 Then 
     Application(strMeetingID) = "<BR><font color=red>" & strName & ":&nbsp;&nbsp;</font>" & strComment
	 Application(strTotalCount) = ItemNum
	 Application(strIPID) = strIP
  Else
	 Application(strTotalCount) = Application(strTotalCount) -1 
     Application(strMeetingID) = Application(strMeetingID) & "<BR><font color=red>" & strName & ":&nbsp;&nbsp;</font>" & strComment
  End if
  Application.UnLock()
End if
%>

<table width=650 align="center">
<tr valign=top>
<td width=350>
<FORM ACTION="" METHOD = POST>


<br>
<center>
<table width=320>
<tr height=35>
<td width=150>
<font color = "#800040"><Strong>会议名</Strong></font>
</td>
<td width=125>
<INPUT TYPE = TEXT class=text size=16 name = Meeting value = <%=strMeeting%>></td>
<td width=45><img src="help.gif" width=15 height=12 alt="会议的名称或主题"><br>
</td>
</tr>

<tr height=35>
<td>
<font color = "#800040"><Strong>密    码</Strong></font>
</td>
<td>
<INPUT TYPE = password  class=text size=16 name = Password value = <%=strPassword%>></td><td width=20><img src="help.gif" width=15 height=12 alt="设定会议的密码"><br>
</td>
</tr>

<tr height=35>
<td>
<font color="#800040"><Strong>姓   名</Strong>
</td>
<td>
<input size=16 name=name  class=text value="<% = strName %>"></td><td width=20><img src="help.gif" width=15 height=12 alt="你的性名"><br>
</td>
</tr>

<tr height=35>
<td class=little>
<font color="#800040">检查IP地址 &nbsp;&nbsp;</font>
</td>
<td>
<INPUT TYPE=RADIO NAME="SHOWIP" value="1" class=white>&nbsp;<img src="help.gif" width=15 height=12 alt="Show the ip address of the conferrees"></td><td width=20><br>
</td>
</tr>
</table>
</center>

<hr size=1px width=300 align=left>
<br>

<table width=300>
<tr>
<td width=10>
</td>
<td>
<textarea name="comment" value="" rows=4 cols=36 style="border:1px solid black ;"></textarea><br>
<input type="submit" value="送出信息" style="border:1px solid white; width:272">
</td>
</tr>
</table>

<br>

</FORM><br><br>
</td>

<TD width=14><IMG alt="" border=0 height=1 src="spacer.gif" width=14></TD><TD bgColor=#003399 vAlign=top width=1><IMG alt="" border=0 height=7 src="spacer.gif" width=1></TD><TD width=18><IMG alt="" border=0 height=1 src="spacer.gif" width=18></TD>

<td>

<%
If nShowIP <> 1 Then
     Response.Write("")
  Application.Lock()
   Response.Write("<Strong>会议内容区 &nbsp; ")
    Response.Write(Application(strTotalCount))
     Response.Write("</Strong><br><br>")
      Response.Write(Application(strMeetingID))
       Application.UnLock() 
Else
  If len(strMeeting)>0 Then
      Application.Lock()
	    Response.Write("<Strong>Comments Clears at &nbsp;")
		   Response.Write(Application(strTotalCount))
		      Response.Write("</Strong><br>")
			      Response.Write("<SELECT>")
			        Response.Write(Application(strIPID))
				     Response.Write("</SELECT><BR>")
					     Response.Write(Application(strMeetingID))
			                Application.UnLock() 
  Else
      Application.Lock()
       Response.Write("")
	    Response.Write("<Strong>Comments Clears at &nbsp;")
	     Response.Write(Application(strTotalCount))
          Response.Write("</Strong><br>")
	        Response.Write("<font style=font-size:12px color=red>Show IP addresses only")
		     Response.Write("works for private meetings that have")
			  Response.Write("been given a name.</font>")
	           Response.Write("<br>")
  		 	   Response.Write(Application(strMeetingID))
                 Application.UnLock() 
  End if
End if
%></td>

</tr>
</table>

</BODY></CENTER>

</HTML>
