<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
Response.Buffer = True
Response.Expires = 0
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache" 
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"
%>
<!--#include file="Connections/conn.asp" -->
<!--��¼Ȩ���жϣ�Session��MD5�����ж�-->'
<%
''���������б�
If request.cookies("oabusyname")=""  Then Response.Redirect("../default.asp")
If request.cookies("cook_allow_control_all_user")=""  Then Response.Redirect("../default.asp")
Function getList(i,sTable,iId,sValue,sName,selfValue)
    If i < 1 Or i > 2 Then
	    getList = ""
		Exit Function
	End If
	Dim strList
	Dim rs
	If i = 1 Then
	    strList = "<select name=""" & sName & """ selfValue=""" & selfValue & """>"
		strList = strList & "<option value="""">��ѡ��</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From " & sTable & "",conn,3,1
		Do While Not rs.BOF And Not rs.EOF
		    strList = strList & "<option value=""" & rs(sValue) & """>" & rs(sValue) & "</option>"
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		getList = strList
	Else
	    strList = "<select name=""" & sName & """ selfValue=""" & selfValue & """>"
		strList = strList & "<option value="""">��ѡ��</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From " & sTable & "",conn,3,1
		Do While Not rs.BOF And Not rs.EOF
		    strList = strList & "<option value=""" & rs(iId) & """>" & rs(sValue) & "</option>"
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		getList = strList
	End If
End Function

Rem Session("CRM_account") �û��ʺ�
Rem Session("CRM_name") �û���
Rem Session("CRM_level") �û��ȼ�


Function getGroupName(gId)
    If Not IsNumeric(gId) Or gId < 0 Then
	    getGroupName = ""
	Else
	    Dim rs,gName
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gId = " & gId,conn,3,1
		If rs.RecordCount <> 1 Then
		    gName = ""
		Else
		    gName = rs("gName")
		End If
		rs.Close
		Set rs = Nothing
		getGroupName = gName
	End If
End Function

Function getLevelName(lId)
    If Not IsNumeric(lId) Or lId <= 0 Then
	    getLevelName = ""
	Else
	    Dim rs,lName
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_level Where lId = " & lId,conn,3,1
		If rs.RecordCount <> 1 Then
		    lName = ""
		Else
		    lName = rs("lName")
		End If
		rs.Close
		Set rs = Nothing
		getLevelName = lName
	End If
End Function

Dim strCounter,strToPrint

Dim rs,intTotalRecords,intTotalPages,intCurrentPage,intPageSize
intCurrentPage = CInt(ABS(Request("pageNum")))
If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
intPageSize = 10

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_user Order By uId",conn,3,1
intTotalRecords = rs.RecordCount
rs.PageSize = intPageSize
intTotalPages = rs.PageCount
If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
strCounter = strCounter & "�� " & intTotalRecords & " ����¼ "
strCounter = strCounter & "�� " & intTotalPages & " ҳ "
strCounter = strCounter & "��ǰ�� " & intCurrentPage & " ҳ "
If intCurrentPage <> 1 And intTotalRecords <> 0 Then
    strCounter = strCounter & "<a href=""?pageNum=1""><<��ҳ</a> "
Else
    strCounter = strCounter & "<<��ҳ "
End If
If intCurrentPage > 1 Then
    strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
Else
    strCounter = strCounter & "<��һҳ "
End If
If intCurrentPage < intTotalPages Then
    strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
Else
    strCounter = strCounter & "��һҳ> "
End If
If intCurrentPage <> intTotalPages Then
    strCounter = strCounter & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
Else
    strCounter = strCounter & "βҳ>>"
End If

Dim i
i = 0
Do While Not rs.BOF And Not rs.EOF
    i = i + 1
	strToPrint = strToPrint & "        <tr>" & VBCrlf
	strToPrint = strToPrint & "          <td align=""center"">" & rs("uId") & "</td>" & VBCrlf
	If rs("uBlock") = False Then
	    strToPrint = strToPrint & "          <td>" & rs("uAccount") & "</td>" & VBCrlf
	Else
	    strToPrint = strToPrint & "          <td><font color=""#FF0000"">" & rs("uAccount") & "</font></td>" & VBCrlf
	End If
	strToPrint = strToPrint & "          <td>" & rs("uPassword") & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & rs("uName") & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & getGroupName(rs("uGroup")) & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & getLevelName(rs("uLevel")) & "</td>" & VBCrlf
	If rs("uBlock") = False Then
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>�޸�</a>] [<a href=""?action=block&uId=" & rs("uId") & """>����</a>] [<a href=""?action=delete&uId=" & rs("uId") & """ onClick=""return confirm('ȷ��ɾ�����û�����\r�ص��������ϣ�');"">ɾ��</a>]</td>" & VBCrlf
	Else
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>�޸�</a>] [<a href=""?action=block&uId=" & rs("uId") & """>�ⶳ</a>] [<a href=""?action=delete&uId=" & rs("uId") & """ onClick=""return confirm('ȷ��ɾ�����û�����\r�ص��������ϣ�');"">ɾ��</a>]</td>" & VBCrlf
	End If
	strToPrint = strToPrint & "        </tr>" & VBCrlf
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���۹���ϵͳ</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function showHideHead(strSrc)
{
	var strFile = strSrc.substring(strSrc.lastIndexOf("/"),strSrc.length);
    if (strFile == "/arrow_up.gif"){
	    oHead.style.display = "none";
		oHeadCtrl.src = "images/arrow_down.gif";
		oHeadCtrl.alt = "��ʾͷ��";
		oHeadBar.title = "��ʾͷ��";		
	}
	else {
	    oHead.style.display = "block";
		oHeadCtrl.src = "images/arrow_up.gif";
		oHeadCtrl.alt = "����ͷ��";
		oHeadBar.title = "����ͷ��";
	}
}

function checkInput(o)
{
    var oo = eval("document.all." + o);
    var num = oo.length;
    for(var i=0;i<num;i++){
	    if(oo[i].value == ""){
		    alert(oo[i].selfValue + "����Ϊ�ա�");
			oo[i].focus();
			return false
			break;
		}
	}
}

if (this.location.href == top.location.href){
    top.location.href = "";
}
-->
</script>
</head>

<body >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="778" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="menu">
        <tr> 
          <td align="left" background='images/tab_top_background_runner.gif'> <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
            <tr>
              <td><img src="images/null.gif" width="1" height="1"></td>
            </tr>
          </table>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td height="5"><img src="images/null.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td bgcolor="#999999">&nbsp;</td>
      </tr>
    </table>
<%
Dim action
action = Trim(Request("action"))
Select Case action
Case "save"
    Call saveData()
Case "edit"
    Call editForm()
Case "saveEdit"
    Call saveEditData()
Case "delete"
    Call deleteData()
Case "block"
    Call block()
Case Else
    Call addForm()
End Select

Dim errMsg
errMsg = CInt(ABS(Request("errMsg")))
Select Case errMsg
Case 1
    errMsg = "<font color=""#FF0000"">��������ݲ����ڡ�</font>"
Case 2
    errMsg = "<font color=""#FF0000"">���͵����ݴ���</font>"
Case Else
    errMsg = ""
End Select

Sub addForm()
%>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="newUser" action="?action=save" method="post" onSubmit="return checkInput('newUser');">
		<tr> 
          <td align="right">&nbsp;</td>
          <td>�½��˺� <% = errMsg %></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>�˺ţ�<input name="account" type="text" id="account" size="16" maxlength="16" selfValue="�˺�">
            ��ʵ������<input name="name" type="text" id="name" size="16" maxlength="16" selfValue="��ʵ����">
            �û��ȼ���<% = getList(2,"baidu_level","lId","lName","level","�û��ȼ�") %></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>���룺<input name="password" type="password" id="password" size="16" maxlength="16" selfValue="����">
            ȷ�����룺<input name="confirmPWS" type="password" id="confirmPWS" size="16" maxlength="16" selfValue="ȷ������">
            ����С�飺<% = getList(2,"baidu_group","gId","gName","group","����С��") %></td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="submit" name="Submit" value=" �� �� "> &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "></td>
        </tr>
		</form>
      </table>
<%
End Sub

Sub saveData()
    Dim uAccount,uPassword,uConfirmPWS,uName,uLevel,uGroup
	uAccount = Trim(Request("account"))
	uPassword = Trim(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	uLevel = CInt(Request("level"))
	uGroup = CInt(Request("group"))
	If uAccount = "" Or uPassword = "" Or uPassword <> uConfirmPWS Or uName = "" Or uGroup = "" Then
	    Response.Redirect("adminuser.asp?errMsg=2")
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From baidu_user",conn,3,2
	rs.AddNew
	rs("uAccount") = uAccount
	rs("uPassword") = uPassword
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub editForm()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	Dim uAccount,uPassword,uName,uLevel,uGroup
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,1
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	uAccount = rs("uAccount")
	uPassword = rs("uPassword")
	uName = rs("uName")
	uLevel = rs("uLevel")
	uGroup = rs("uGroup")
%>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="newUser" action="?action=saveEdit" method="post" onSubmit="return checkInput('newUser');">
		<tr> 
          <td align="right">&nbsp;</td>
          <td>�޸��˺�</td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>�˺ţ� 
              <input name="account" type="text" id="account" value="<% = uAccount %>" size="16" maxlength="16">
            ��ʵ������ 
              <input name="name" type="text" id="name" value="<% = uName %>" size="16" maxlength="16">
            �û��ȼ���<% = getList(2,"baidu_level","lId","lName","level","�û��ȼ�") %>
              <input name="id" type="hidden" id="id" value="<% = uId %>"></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>���룺 
              <input name="password" type="password" id="password" value="<% = uPassword %>" size="16" maxlength="16">
            ȷ�����룺 
              <input name="confirmPWS" type="password" id="confirmPWS" value="<% = uPassword %>" size="16" maxlength="16">
            ����С�飺<% = getList(2,"baidu_group","gId","gName","group","����С��") %></td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="submit" name="Submit" value=" �� �� "> &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "></td>
        </tr>
		</form>
      </table>
<script language="JavaScript">
<!--
    var level = "<% = uLevel %>";
	var group = "<% = uGroup %>";
	for (i=0;i<document.newUser.level.options.length;i++){
        if (document.newUser.level.options[i].value == level)
            document.newUser.level.options[i].selected = true;
    }
	for (i=0;i<document.newUser.group.options.length;i++){
        if (document.newUser.group.options[i].value == group)
            document.newUser.group.options[i].selected = true;
    }
-->
</script>
<%
End Sub

Sub saveEditData()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("id")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
    Dim uAccount,uPassword,uConfirmPWS,uName,uLevel,uGroup
	uAccount = Trim(Request("account"))
	uPassword = Trim(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	uLevel = CInt(Request("level"))
	uGroup = CInt(Request("group"))
	If uAccount = "" Or uPassword = "" Or uPassword <> uConfirmPWS Or uName = "" Or uGroup = "" Then
	    Response.Redirect("adminuser.asp?errMsg=2")
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	rs("uAccount") = uAccount
	rs("uPassword") = uPassword
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub deleteData()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	''Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub block()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	''Dim rs
	Dim uBlock
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	uBlock = rs("uBlock")
	If uBlock = True Then
	    rs("uBlock") = False
	Else
	    rs("uBlock") = True
	End If
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub
%>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#999999" id="oHeadBar" style="cursor: hand;" title="����ͷ��" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="����ͷ��" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> 
            <% = strCounter %>
          </td>
          <td align="right"> [<a href="adminuser.asp">�û��б�</a>] [<a href="adminuser.asp">�½��û�</a>]</td>
        </tr>
      </table> 
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr>
          <td width="40" align="center" bgcolor="menu">���</td>
          <td width="80" align="center" bgcolor="menu">�˺�</td>
          <td width="80" align="center" bgcolor="menu">����</td>
          <td width="80" align="center" bgcolor="menu">��ʵ����</td>
          <td align="center" bgcolor="menu">����С��</td>
          <td width="60" align="center" bgcolor="menu">�û��ȼ�</td>
          <td align="center" bgcolor="menu">����</td><% = strToPrint %>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#999999"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
