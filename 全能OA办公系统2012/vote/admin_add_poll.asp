
<!--#include file="config.asp"-->
<%
Call AddPoll()
%>
<!--#include file="top.asp"-->
<!--#include file="menu.asp"--><br><br>
<%
T=request.form("T")
N=request.form("N")
if T="" or N="" then
%>

   <form action="" method="post">  
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#7C96B8">
    <tr> 
      <td  colspan="2" ><font color="#FF0000"><img src="images/poll.gif" width="13" height="15" align="absmiddle" vspace="2" hspace="2">����µ�����Ŀ<font color="#000000">(��һ��)</font></font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="Info_Title" width="18%" align="center">��������: </td>
      <td class="Info_Title" width="82%"> 
        <input name="T" type="text" size="30" maxlength="100">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="18%" align="center">����ѡ����:</td>
      <td width="82%">
        <select name="N">
		<%for i=1 to 50%>
          <option value="<%=i%>"><%=i%></option>
		  <%next%>
        </select>
      </td>
    </tr>
    <tr  bgcolor="#FFFFFF"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit2" value="��һ��">
      </td>
    </tr>
  </table>
  </form>
 <%else%>
  <form action="" method="post" name="add" >
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#7C96B8">
    <tr> 
      <td  colspan="2" ><font color="#FF0000"><img src="images/poll.gif" width="13" height="15" align="absmiddle" vspace="2" hspace="2">����µ�����Ŀ<font color="#000000">(�ڶ���)</font></font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="Info_Title" width="18%" align="center">��������: </td>
      <td class="Info_Title" width="82%"> 
        <input name="title" type="text" id="pollquestion" size="30" maxlength="100" value="<%=T%>">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="18%" align="center">��ʼʱ��:</td>
      <td width="82%"> 
        <input name="startt" type="text" id="startdate" value="<%=Date()%>" size="10" maxlength="10">
        (�磺2003-1-1)</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="18%" align="center">����ʱ��: </td>
      <td width="82%"> 
        <input name="endt" type="text" id="enddate" value="<%=Date+7%>" size="10" maxlength="10">
        (�磺2003-1-1)</td>
    </tr>
	<%for i=1 to N%>
    <tr bgcolor="#FFFFFF"> 
      <td align="center">ѡ��<%=i%>: </td>
      <td> 
        <input name="o<%=i%>" type="text" size="20" maxlength="50">
      </td>
    </tr>
	<%next %>
	<tr bgcolor="#FFFFFF"> 
      <td align="right">��Ϊ��ǰ������Ŀ��</td>
            
      <td colspan="2"> 
        <input type="radio" name="active" value="1" checked>
              �� 
         <input name="active" type="radio" value="0" >
              �� 
            </td>
          </tr>
   
    <tr  bgcolor="#FFFFFF"> 
      <td colspan="2" align="center"> 
        <input type="hidden" name="A" value="AddPoll">
        <input type="hidden" name="N" value="<%=N%>">
        <input type="submit" name="Submit" value="�� ��">
      </td>
    </tr>
  </table>
</form>
<%end if%>

<!--#include file="foot.asp"-->
<%
Sub AddPoll()
A=RequestText(request.form("A"))
If Request.ServerVariables("REQUEST_METHOD")="POST" and A="AddPoll" Then
	title = RequestText(Request.Form("title"))
	if title="" then out "�������ⲻ��Ϊ��"
	if not ISDATE(RequestText(Request.Form("startt"))) or not isdate(RequestText(Request.Form("endt"))) then
		out "���ڸ�ʽ���ԡ�"
	end if 
	startt = CDATE(RequestText(Request.Form("startt")))
	endt = CDATE(RequestText(Request.Form("endt")))
	active = RequestText(Request.Form("active"))
	n=CINT(request.form("N"))
	for i=1 to n
		if RequestText(Request.Form("o"&i))="" then out "���е���ѡ���Ϊ��"
	next
	
	call OpenDB()
	if active=1 then
		DbConn.Execute("UPDATE Question SET Q_Active=0")
	end if
	DbConn.Execute("Insert into Question (Q_Title,Q_StartDate,Q_EndDate,Q_Active) values('"& title & "','" & startt & "','" & endt & "','" & active & "')")
	Set rs = DbConn.Execute( "SELECT @@IDENTITY" )
	id = rs(0)
	rs.Close
	Set rs = Nothing
	for i=1 to n
	DbConn.Execute("Insert into Answer (Q_ID,A_Answer) values('" & id & "','" & RequestText(Request.Form("o"&i)) & "')")
	next
	call CloseDB()
	Response.Redirect ("admin_poll.asp?id="&id)
end if
End Sub

%>
