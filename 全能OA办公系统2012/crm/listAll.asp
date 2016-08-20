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

<!--��¼Ȩ���жϣ�Session��MD5�����ж�-->
<%
Rem Session("CRM_account") �û��ʺ�
Rem Session("CRM_name") �û���
Rem Session("CRM_level") �û��ȼ�

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Function getGroupName(gId)
    If gId = "" Then
	    getGroupName = ""
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gId = " & gId,conn,3,1
		If rs.RecordCount <> 1 Then
		    getGroupName = ""
		Else
		    getGroupName = rs("gName")
		End If
		rs.Close
		Set rs = Nothing
	End If
End Function

Function getUserList(intLevel,intGroup)
    Dim rs,strUserList
	strUserList = "'" & Session("CRM_name") & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uLevel < " & intLevel & " And uGroup = " & intGroup,conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If strUserList = "" Then
		    strUserList = "'" & rs("uName") & "'"
		Else
		    strUserList = strUserList & ",'" & rs("uName") & "'"
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	getUserList = strUserList
End Function

Dim strNormal,strAdmin,strCounter,strToPrint
'strNormal = strNormal & "        <tr>" & VBCrlf
'strNormal = strNormal & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
'strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
'strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
'strNormal = strNormal & "          <td width=""80"" align=""center"" bgcolor=""menu"">����</td>" & VBCrlf
'strNormal = strNormal & "        </tr>" & VBCrlf

'strAdmin = strAdmin & "        <tr>" & VBCrlf
'strAdmin = strAdmin & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
'strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
'strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
'strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">����</td>" & VBCrlf
'strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
'strAdmin = strAdmin & "        </tr>" & VBCrlf

strToPrint = strToPrint & "        <tr>" & VBCrlf
strToPrint = strToPrint & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">�ͻ��ȼ�</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
strToPrint = strToPrint & "        </tr>" & VBCrlf

''strCounter = ""

'If Session("CRM_level") = 9 Then
'    strToPrint = strAdmin
'Else
'    strToPrint = strNormal
'End If

Dim action
Dim arrList
action = Trim(Request("action"))
If action <> "" Then
    Session("CRM_action") = action
Else
    action = Session("CRM_action")
End If
Select Case action
Case "com"
    If Trim(Request("searchCom")) <> "" Then
        Session("CRM_keyWords") = Trim(Request("searchCom"))
	End If
	If Session("CRM_keyWords") = "" Then
	    arrList = listAll
	    strToPrint = strToPrint &  arrList(1)
	    strCounter = arrList(0)
	Else
	    arrList = searchCom
	    strToPrint = strToPrint &  arrList(1)
	    strCounter = arrList(0)
	End If
Case "url"
    If Trim(Request("searchUrl")) <> "" Then
        Session("CRM_keyWords") = Trim(Request("searchUrl"))
	End If
	If Session("CRM_keyWords") = "" Then
	    arrList = listAll
	    strToPrint = strToPrint &  arrList(1)
	    strCounter = arrList(0)
	Else
	    arrList = searchUrl
		strToPrint = strToPrint &  arrList(1)
	    strCounter = arrList(0)
	End If
Case "killSession"
    Session("CRM_keyWords") = ""
	Session("CRM_action") = ""
	Response.Redirect("listAll.asp")
Case Else
	arrList = listAll
	strToPrint = strToPrint & arrList(1)
	strCounter = arrList(0)
End Select

Function searchCom()
    Dim rs,strOut(2),strUserList
	Dim intTotalRecords,intTotalPages,intCurrentPage,intPageSize
	intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 10
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Session("CRM_level") = 9 Then
	    rs.Open "Select * From baidu_client Where cCompany Like '%" & Session("CRM_keyWords") & "%' Order By cId Desc",conn,3,1
	Else
	    strUserList = getUserList(Session("CRM_level"),Session("CRM_group"))
		rs.Open "Select * From baidu_client Where cCompany Like '%" & Session("CRM_keyWords") & "%' And cUser In (" & strUserList & ") Order By cId Desc",conn,3,1
	    'If Session("CRM_level") = 2 Then
		'	rs.Open "Select * From baidu_client Where cCompany Like '%" & Session("CRM_keyWords") & "%' And cLocal = '" & getGroupName(Session("CRM_group")) & "' Order By cId Desc",conn,3,1
		'Else
	    '    rs.Open "Select * From baidu_client Where cCompany Like '%" & Session("CRM_keyWords") & "%' And cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
		'End If
	End If
	intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
    strOut(0) = strOut(0) & "�� " & intTotalRecords & " ����¼ "
    strOut(0) = strOut(0) & "�� " & intTotalPages & " ҳ "
    strOut(0) = strOut(0) & "��ǰ�� " & intCurrentPage & " ҳ "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=1""><<��ҳ</a> "
    Else
        strOut(0) = strOut(0) & "<<��ҳ "
    End If
    If intCurrentPage > 1 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
    Else
        strOut(0) = strOut(0) & "<��һҳ "
    End If
    If intCurrentPage < intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
    Else
        strOut(0) = strOut(0) & "��һҳ> "
    End If
    If intCurrentPage <> intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
    Else
        strOut(0) = strOut(0) & "βҳ>>"
    End If
	
	Dim k
	k = 0
	Do While Not rs.BOF And Not rs.EOF
	    k = k + 1
	    strOut(1) = strOut(1) & "        <tr>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td align=""center"">" & rs("cId") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" &  VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & rs("cType") & "</td>" & VBCrlf
		'If Session("CRM_level") = 9 Then
	        strOut(1) = strOut(1) & "        <td>" & rs("cUser") & "</td>" & VBCrlf
		'End If
	    strOut(1) = strOut(1) & "        </tr>" & VBCrlf
		If k >= intPageSize Then Exit Do
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	searchCom = strOut
End Function

Function searchUrl()
    Dim rs,strOut(2),strUserList
	Dim intTotalRecords,intTotalPages,intCurrentPage,intPageSize
	intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 10
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Session("CRM_level") = 9 Then
	    rs.Open "Select * From baidu_client Where cHomepage Like '%" & Session("CRM_keyWords") & "%' Order By cId Desc",conn,3,1
	Else
	    strUserList = getUserList(Session("CRM_level"),Session("CRM_group"))
		rs.Open "Select * From baidu_client Where cHomepage Like '%" & Session("CRM_keyWords") & "%' And cUser In (" & strUserList & ") Order By cId Desc",conn,3,1
	    'If Session("CRM_level") = 2 Then
		'	rs.Open "Select * From baidu_client Where cHomepage Like '%" & Session("CRM_keyWords") & "%' And cLocal = '" & getGroupName(Session("CRM_group")) & "' Order By cId Desc",conn,3,1
		'Else
	    '    rs.Open "Select * From baidu_client Where cHomepage Like '%" & Session("CRM_keyWords") & "%' And cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
		'End If
	End If
	intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
    strOut(0) = strOut(0) & "�� " & intTotalRecords & " ����¼ "
    strOut(0) = strOut(0) & "�� " & intTotalPages & " ҳ "
    strOut(0) = strOut(0) & "��ǰ�� " & intCurrentPage & " ҳ "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=1""><<��ҳ</a> "
    Else
        strOut(0) = strOut(0) & "<<��ҳ "
    End If
    If intCurrentPage > 1 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
    Else
        strOut(0) = strOut(0) & "<��һҳ "
    End If
    If intCurrentPage < intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
    Else
        strOut(0) = strOut(0) & "��һҳ> "
    End If
    If intCurrentPage <> intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
    Else
        strOut(0) = strOut(0) & "βҳ>>"
    End If
	
	Dim k
	k = 0
	Do While Not rs.BOF And Not rs.EOF
	    k = k + 1
	    strOut(1) = strOut(1) & "        <tr>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td align=""center"">" & rs("cId") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" &  VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & rs("cType") & "</td>" & VBCrlf
		'If Session("CRM_level") = 9 Then
	        strOut(1) = strOut(1) & "        <td>" & rs("cUser") & "</td>" & VBCrlf
		'End If
	    strOut(1) = strOut(1) & "        </tr>" & VBCrlf
		If k >= intPageSize Then Exit Do
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	strOut(0) = "�� ����¼ �� ҳ ��ǰ�� ҳ &lt;&lt;��ҳ &lt;��һҳ ��һҳ&gt; βҳ&gt;&gt;"
	searchUrl = strOut
End Function

Function listAll()
    Dim rs,strOut(2),strUserList
	Dim intTotalRecords,intTotalPages,intCurrentPage,intPageSize
	intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 10

	Set rs = Server.CreateObject("ADODB.Recordset")
	If Session("CRM_level") = 9 Then
	    rs.Open "Select * From baidu_client Order By cId Desc",conn,3,1
	Else
	    strUserList = getUserList(Session("CRM_level"),Session("CRM_group"))
		'Response.Write(strUserList)
		'Response.End()
		rs.Open "Select * From baidu_client Where cUser In (" & strUserList & ") Order By cId Desc",conn,3,1
	    'If Session("CRM_level") = 2 Then
	    '    rs.Open "Select * From baidu_client Where cLocal = '" & getGroupName(Session("CRM_group")) & "' Order By cId Desc",conn,3,1
		'Else
	    '    rs.Open "Select * From baidu_client Where cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
		'End If
	End If
	intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
    strOut(0) = strOut(0) & "�� " & intTotalRecords & " ����¼ "
    strOut(0) = strOut(0) & "�� " & intTotalPages & " ҳ "
    strOut(0) = strOut(0) & "��ǰ�� " & intCurrentPage & " ҳ "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=1""><<��ҳ</a> "
    Else
        strOut(0) = strOut(0) & "<<��ҳ "
    End If
    If intCurrentPage > 1 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
    Else
        strOut(0) = strOut(0) & "<��һҳ "
    End If
    If intCurrentPage < intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
    Else
        strOut(0) = strOut(0) & "��һҳ> "
    End If
    If intCurrentPage <> intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
    Else
        strOut(0) = strOut(0) & "βҳ>>"
    End If
	
	Dim k
	k = 0
	Do While Not rs.BOF And Not rs.EOF
	    k = k + 1
	    strOut(1) = strOut(1) & "        <tr>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td align=""center"">" & rs("cId") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" &  VBCrlf
	    strOut(1) = strOut(1) & "        <td>&nbsp;" & rs("cType") & "</td>" & VBCrlf
		'If Session("CRM_level") = 9 Then
	        strOut(1) = strOut(1) & "        <td>" & rs("cUser") & "</td>" & VBCrlf
		'End If
	    strOut(1) = strOut(1) & "        </tr>" & VBCrlf
		If k >= intPageSize Then Exit Do
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	listAll = strOut
End Function
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

if (this.location.href == top.location.href){
    top.location.href = "";
}

function ftnCom()
{
    if (document.searchComForm.searchCom.value == ""){
	    alert("������Ҫ��ѯ�Ĺ�˾���ơ�");
		document.searchComForm.searchCom.focus();
		return false;
	}
}

function ftnUrl()
{
    if (document.searchUrlForm.searchUrl.value == ""){
	    alert("������Ҫ��ѯ�Ĺ�˾��ַ��");
		document.searchUrlForm.searchUrl.focus();
		return false;
	}
}

var IsOpen = "<% = Session("CRM_planWin") %>";
if(IsOpen != "True"){
    window.open("plan_records.asp","View","menubar=no,width=360,height=480,resizable=no");
}
-->
</script>
<style type="text/css">
.style7 {color: #2d4865}
.style8 {color: #0d79b3;
	font-weight: bold;
}
</style>
</head>

<body topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style8"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style8"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">����ϵͳ</td>
                </tr>
            </table></td>
            <td width="1"><span class="style8"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br>
<table width="550" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
     
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td height="5"><img src="images/null.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td bgcolor="#88ADDF">&nbsp;</td>
      </tr>
    </table>
    <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
      <tr> 
        <td width="40" align="right">&nbsp;</td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <form name="searchComForm" method="post" action="?action=com" onSubmit="return ftnCom();">
	  <tr>
        <td>����˾���������� 
              <input name="searchCom" type="text" id="searchCom" size="24" maxlength="36" <% If Session("CRM_keyWords") <> "" And Session("CRM_action") = "com" Then Response.Write("value=""" & Session("CRM_keyWords") & """") %> onFocus="this.value='';">
              <input name="Search" type="submit" id="Search" value="����"></td>
      </tr></form>
     </table></td>
    </tr>
    <tr> 
      <td width="40" align="right">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <form name="searchUrlForm" method="post" action="?action=url" onSubmit="return ftnUrl();">
		  <tr>
            <td>����˾��ַ������ 
              <input name="searchUrl" type="text" id="searchUrl" size="24" maxlength="36" <% If Session("CRM_keyWords") <> "" And Session("CRM_action") = "url" Then Response.Write("value=""" & Session("CRM_keyWords") & """") %> onFocus="this.value='';">
              <input name="Submit" type="submit" id="Submit" value="����"></td>
          </tr></form>
        </table></td>
    </tr>
  </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="����ͷ��" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="����ͷ��" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>
            <% = strCounter %>
          </td>
          <td align="right">[<a href="?action=killSession">����ȫ���б�</a>]</td>
        </tr>
      </table> 
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF"><% = strToPrint %>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
