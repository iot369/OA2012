
<!--#include file="config.asp"-->
<%
Call DelPoll()
%>
<!--#include file="top.asp"-->
<!--#include file="menu.asp"-->
<br><br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#7C96B8">
  <tr> 
    <td width="6%" align="center" >״̬</td>
    <td width="48%" >��������</td>
    <td width="8%" align="center" >ͶƱ��</td>
    <td width="14%" align="center" >��ʼʱ��</td>
    <td width="14%" align="center" >����ʱ��</td>
    <td width="10%" align="center" >����</td>
  </tr>
  <%
Dim PageNo	
Dim mpage
Dim startime
Dim endtime
PageNo = 10
Dim PageNum		
Dim RCount		
Dim i
			
Dim sql	
Dim rs	
call OpenDB()
sql = "Select * From Question Order By Q_StartDate DESC, Q_EndDate DESC;"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, DBConn, 1,1
If Not rs.EOF Then
	call CheckPage()
	rs.AbsolutePage = PageNum
	For RCount = 1 To rs.PageSize
	startime=rs("Q_StartDate")
	endtime=rs("Q_EndDate")
%>
  <tr bgcolor="#FFFFFF"> 
    <td width="6%" align="center">
	<img src="images/<%if rs("Q_Active") = True Then response.write "poll.gif" else response.write "poll1.gif"%>" width="15" height="13" vspace="0" hspace="0" align="absmiddle"> 
    </td>
    <td width="48%"> 
      <%response.write "<a href=""default.asp?Q_ID="&rs("Q_ID")&""" > "&rs("Q_Title")&"</a>"%>
    </td>
    <td width="8%" align="center"><%=rs("Q_Vote")%></td>
    <td width="14%" align="center"><%=startime%></td>
    <td width="14%" align="center"><%=endtime%></td>
    <td width="10%" align="center"><a href="admin_edit_poll.asp?id=<%=rs("Q_ID")%>"><img src="images/edit.gif" width="15" height="13" border="0" hspace="2" alt="�༭"></a><a href="admin_poll.asp?pid=<%=rs("Q_ID")%>&T=DELQ"><img src="images/delete.gif" width="15" height="13" border="0" hspace="2" alt="ɾ��"></a></td>
  </tr>
  <%
	rs.MoveNext
	If rs.EOF Then Exit For
	Next
	if mpage>1 then
%>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="6"> 
      <%call DisplayPage()%>
    </td>
  </tr>
  <%
	end if
Else
	response.write "<tr bgcolor=""#FFFFFF""> <td colspan=""6"">���޵���</td></tr>"	
End IF
rs.Close
Set rs = Nothing
call CloseDB()
%>
</table>
      
      <table width="98%" border="0" cellspacing="0" cellpadding="2" align="center" class="Info_Title">
  <tr> 
    <td align="center"><font color="#0000FF">ͼʾ˵����</font><img src="images/poll.gif" width="13" height="15" align="absmiddle"> 
      ���ڽ��еĵ�����Ŀ��<img src="images/poll1.gif" width="13" height="15" align="absmiddle"> 
      ����������Ŀ </td>
  </tr>
</table>
   <br><br><br>

<!--#include file="foot.asp"-->
<%
Sub DelPoll()
	T = RequestText(Request.Querystring("T"))
	pID = RequestText(Request.Querystring("pid"))
	if T="DELQ" and pID<>"" then
		call OpenDB()
		DbConn.Execute("Delete From Answer WHERE Q_ID = " & pID )
		DbConn.Execute("Delete From Question WHERE Q_ID = " & pID )
		call CloseDB()
		Response.Redirect "admin_poll.asp" 
	end if
End Sub
%>
