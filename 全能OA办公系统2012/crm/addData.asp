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

Dim strNormal,strAdmin,strToPrint
strNormal = strNormal & "        <tr>" & VBCrlf
strNormal = strNormal & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
strNormal = strNormal & "          <td width=""80"" align=""center"" bgcolor=""menu"">����</td>" & VBCrlf
strNormal = strNormal & "        </tr>" & VBCrlf

strAdmin = strAdmin & "        <tr>" & VBCrlf
strAdmin = strAdmin & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">����</td>" & VBCrlf
strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
strAdmin = strAdmin & "        </tr>" & VBCrlf

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

Function getGroup(gName)
    If gName = "" Then
	    getGroup = 0
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gName = '" & gName & "'",conn,3,1
		If rs.RecordCount <> 1 Then
		    getGroup = 0
		Else
		    getGroup = rs("gId")
		End If
		rs.Close
		Set rs = Nothing
	End If
End Function

''���������б�
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

function checkInput()
{
    if(document.all.company.value == ""){
	    alert("�����빫˾���ơ�");
		document.all.company.focus();
		return false;
	}
	if(document.all.linkman.value == ""){
	    alert("��������ϵ�ˡ�");
		document.all.linkman.focus();
		return false;
	}
}
-->
</script>
</head>

<body  topmargin="0" leftmargin="0">
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
      </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="����ͷ��" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="����ͷ��" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
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
Case Else
    Call addForm()
End Select

Sub addForm()
%>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form name="addDataForm" action="?action=save" method="post" onSubmit="return checkInput();">
          <tr> 
            <td width="138" height="25" align="right">��˾���ƣ�</td>
            <td width="90"><input name="company" type="text" id="company" size="24" maxlength="36"></td>
            <td width="138" align="right">��ϵ�ˣ�</td>
            <td width="90"><input name="linkman" type="text" id="linkman" size="24" maxlength="48"></td>
          </tr>
          <tr> 
            <td width="138" height="25" align="right">��ַ��</td>
            <td><input name="homepage" type="text" id="homepage" size="24" maxlength="36"></td>
            <td width="138" align="right">��ϵ�绰��</td>
            <td><input name="tel" type="text" id="tel" size="24" maxlength="36"></td>
          </tr>
          <tr> 
            <td width="138" height="25" align="right">�������䣺</td>
            <td><input name="email" type="text" id="email" size="24" maxlength="36"></td>
            <td width="138" align="right">��˾��ַ��</td>
            <td><input name="address" type="text" id="address" size="36"></td>
          </tr>
          <tr> 
            <td width="138" height="25" align="right">������</td>
            <td> 
              <% = getList(1,"baidu_area","","areaName","area","ҵ������") %></td>
            <td width="138" align="right">����С����</td>
            <td><% = getList(1,"baidu_square","","squareName","square","ҵ��С��") %></td>
          </tr>
		  <tr> 
            <td width="138" height="25" align="right">�ͻ��ȼ���</td>
            <td><% = getList(1,"baidu_clientsType","","clientsType","type","�ͻ��ȼ�") %>
            </td>
            <td width="138" align="right">��ҵ���ͣ�</td>
            <td>
              <% = getList(1,"baidu_clientsTrade","","clientsTrade","trade","��ҵ����") %>
            </td>
          </tr>
          <tr> 
            <td width="138" height="25" align="right">ҵ��Ա��</td>
            <td> 
              <input name="user" type="text" id="user" value="<% = Session("CRM_name") %>" size="12" maxlength="16" readonly="true"></td>
            <td width="138" align="right">ҵ���飺</td>
            <td><input name="group" type="text" id="group" size="16" maxlength="24" value="<% = getGroupName(Session("CRM_group")) %>" readonly="true"></td>
          </tr>
          <tr> 
            <td width="138" align="right">�ͻ�������</td>
            <td colspan="3"><textarea name="info" rows="4" id="info" style="width: 80%;"></textarea> 
            </td>
          </tr>
          <!--<tr> 
            <td width="100" align="right">�ݷü�¼��</td>
            <td colspan="3"><textarea name="record" rows="8" id="record" style="width: 80%;"></textarea></td>
          </tr>-->
          <tr> 
            <td colspan="4" align="center"> <input type="submit" name="Submit" value=" �� �� "> 
              &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "></td>
          </tr>
          <tr> 
            <td colspan="4"><hr size="1" noshade> <font color="#FF0000"><span class="emRed">˵����</span></font><br>
���ͬһ���������ж�����ݣ����á�|�����ŷָ������磺�ͻ��ж����ϵ�绰���ֱ���80000001 80000002 80000003������ϵ�绰һ��Ӧ��������ַ�Ϊ <strong>80000001|80000002|80000003</strong></td>
          </tr>
        </form>
      </table>
      <%
End Sub

Sub saveData()
    Dim cCompany,cLinkman,cHomepage,cTel,cEmail,cUser
	Dim cArea,cLocal,cType,cStatus,cAddress
	Dim cTrade,cSquare,cGroup
	'Dim cDomainBegin,cDomainEnd
	'Dim cSpaceBegin,cSpaceEnd
	'Dim cOprationBegin,cOprationEnd
	Dim cInfo
	cCompany = Trim(Request("company"))
	cLinkman = Trim(Request("linkman"))
	cHomepage = Trim(Request("homepage"))
	If cHomepage = "" Then cHomepage = "&nbsp;"
	cTel = Trim(Request("tel"))
	cEmail = Trim(Request("email"))
	cUser = Trim(Request("user"))
	cArea = Trim(Request("area"))
	cLocal = Trim(Request("local"))
	cType = Trim(Request("type"))
	cStatus = Trim(Request("status"))
	cAddress = Trim(Request("address"))
	cTrade = Trim(Request("trade"))
	cSquare = Trim(Request("square"))
	cGroup = Session("CRM_group")
    'cDomainBegin = Trim(Request("domainBegin"))
    'cDomainEnd = Trim(Request("domainEnd"))
    'cSpaceBegin = Trim(Request("spaceBegin"))
    'cSpaceEnd = Trim(Request("spaceEnd"))
    'cOprationBegin = Trim(Request("oprationBegin"))
    'cOprationEnd = Trim(Request("oprationEnd"))
	cInfo = htmlEncode2(Request("info"))
	Dim rs,cId,flag
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_client Where cCompany = '" & cCompany & "' Or cLinkman = '" & cLinkman & "'",conn,3,1
	If rs.RecordCount > 0 Then
	    Response.Write("<font color=""#FF0000"">�˿ͻ��Ѿ����ڡ�</font><br><br>")
		Response.Write("<input type=""button"" value="" �� �� "" onClick=""history.back();"">")
		Response.End()
	End If
	rs.Close
	''
	rs.Open "Select Top 1 * From baidu_client",conn,3,2
	rs.AddNew
	rs("cCompany") = cCompany
	rs("cLinkman") = cLinkman
	rs("cHomepage") = cHomepage
	rs("cTel") = cTel
	rs("cEmail") = cEmail
	rs("cUser") = cUser
	rs("cArea") = cArea
	rs("cLocal") = cLocal
	rs("cType") = cType
	rs("cStatus") = cStatus
	rs("cAddress") = cAddress
	rs("cTrade") = cTrade
	rs("cSquare") = cSquare
	rs("cGroup") = cGroup
    'If cDomainBegin <> "" Then rs("cDomainBegin") = cDomainBegin
    'If cDomainEnd <> "" Then rs("cDomainEnd") = cDomainEnd
    'If cSpaceBegin <> "" Then rs("cSpaceBegin") = cSpaceBegin
	'If cSpaceEnd <> "" Then rs("cSpaceEnd") = cSpaceEnd
    'If cOprationBegin <> "" Then rs("cOprationBegin") = cOprationBegin
    'If cOprationEnd <> "" Then rs("cOprationEnd") = cOprationEnd
	rs("cInfo") = cInfo
	rs.Update
	cId = rs("cId")
	rs.Close
	Set rs = Nothing
	Response.Redirect("view.asp?cId=" & cId)
End Sub

Sub editform()
    Dim cId
	cId = CInt(ABS(Request("cId")))
	If Not IsNumeric(cId) Or cId <= 0 Then
	    Response.Write("<font color=""#FF0000""><b>�����������</b></font><br><br>")
		Response.Write("<input type=""button"" value="" �� �� "" onClick=""location.replace('listAll.asp');"">")
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_client Where cId = " & cId,conn,3,1
		If rs.RecordCount <> 1 Then		
	        Response.Write("<font color=""#FF0000""><b>�����������</b></font><br><br>")
		    Response.Write("<input type=""button"" value="" �� �� "" onClick=""location.replace('listAll.asp');"">")
			Response.End()
		End If
		Dim cCompany,cLinkman,cHomepage,cTel,cEmail,cUser
	    Dim cArea,cLocal,cType,cStatus,cAddress		
		Dim cTrade,cSquare,cGroup
	    'Dim cDomainBegin,cDomainEnd
	    'Dim cSpaceBegin,cSpaceEnd
	    'Dim cOprationBegin,cOprationEnd
	    Dim cInfo
		cCompany = rs("cCompany")
		cLinkman = rs("cLinkman")
		cHomepage = rs("cHomepage")
		cTel = rs("cTel")
		cEmail = rs("cEmail")
		cUser = rs("cUser")
		cArea = rs("cArea")
		cLocal = rs("cLocal")
		cType = rs("cType")
		cStatus = rs("cStatus")
		cAddress = rs("cAddress")
		cTrade = rs("cTrade")
		cSquare = rs("cSquare")
		cGroup = rs("cGroup")
		'cDomainBegin = rs("cDomainBegin")
		'cDomainEnd = rs("cDomainEnd")
		'cSpaceBegin = rs("cSpaceBegin")
		'cSpaceEnd = rs("cSpaceEnd")
		'cOprationBegin = rs("cOprationBegin")
		'cOprationEnd = rs("cOprationEnd")
		cInfo = htmlEncode3(rs("cInfo"))
		rs.Close
		Set rs = Nothing
%>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form name="editDataForm" action="?action=saveEdit" method="post" onSubmit="return checkInput();">
          <tr> 
            <td width="109" height="25" align="right">��˾���ƣ�</td>
            <td width="134"><input name="company" type="text" id="company" value="<% = cCompany %>" size="24" maxlength="36"> 
              <input name="id" type="hidden" id="id" value="<% = cId %>"></td>
            <td width="107" align="right">��ϵ�ˣ�</td>
            <td width="180"><input name="linkman" type="text" id="linkman" value="<% = cLinkman %>" size="24" maxlength="48"></td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">��ַ��</td>
            <td><input name="homepage" type="text" id="homepage" value="<% = cHomepage %>" size="24" maxlength="36"></td>
            <td width="107" align="right">��ϵ�绰��</td>
            <td><input name="tel" type="text" id="tel" value="<% = cTel %>" size="24" maxlength="36"></td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">�������䣺</td>
            <td><input name="email" type="text" id="email" value="<% = cEmail %>" size="24" maxlength="36"></td>
            <td width="107" align="right">��˾��ַ��</td>
            <td> 
              <input name="address" type="text" id="address" value="<% = cAddress %>" size="36"></td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">������</td>
            <td>
              <% = getList(1,"baidu_area","","areaName","area","ҵ������") %>
            </td>
            <td width="107" align="right">����С����</td>
            <td>
              <% = getList(1,"baidu_square","","squareName","square","ҵ��С��") %>
            </td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">�ͻ����ͣ�</td>
            <td>
              <% = getList(1,"baidu_clientsType","","clientsType","type","�ͻ�����") %>
            </td>
            <td width="107" align="right">��ҵ���ͣ�</td>
            <td>
              <% = getList(1,"baidu_clientsTrade","","clientsTrade","trade","��ҵ����") %>
            </td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">ҵ��Ա��</td>
            <td><input name="user" type="text" id="user" value="<% = cUser %>" size="12" maxlength="16" readonly="true"> 
            </td>
            <td width="107" align="right">ҵ���飺</td>
            <td><input name="group" type="text" id="group" size="16" maxlength="24" value="<% = getGroupName(cGroup) %>" readonly="true"></td>
          </tr>
          <tr> 
            <td width="109" height="25" align="right">�ͻ�������</td>
            <td colspan="3"><textarea name="info" rows="4" id="info" style="width: 80%;"><% = cInfo %></textarea> 
            </td>
          </tr>
          <!--<tr> 
            <td width="100" align="right">�ݷü�¼��</td>
            <td colspan="3"><textarea name="record" rows="8" id="record" style="width: 80%;"><%' = cRecord %></textarea></td>
          </tr>-->
          <tr> 
            <td colspan="4" align="center"> <input type="submit" name="Submit" value=" �� �� "> 
              &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "></td>
          </tr>
          <tr> 
            <td colspan="4"><hr size="1" noshade>              <font color="#FF0000"><span class="emRed">˵����</span></font><br>
              ���ͬһ���������ж�����ݣ����á�|�����ŷָ������磺�ͻ��ж����ϵ�绰���ֱ���80000001 80000002 
            80000003������ϵ�绰һ��Ӧ��������ַ�Ϊ <strong>80000001|80000002|80000003</strong></td>
          </tr>
        </form>
      </table>
      <script language="JavaScript">
<!--
var strType = "<% = cType %>";
for(var i=0;i<document.all.type.options.length;i++){
    if(document.all.type.options[i].value == strType){
	    document.all.type.options[i].selected = true;
	}
}

var strTrade = "<% = cTrade %>";
for(var i=0;i<document.all.trade.options.length;i++){
    if(document.all.trade.options[i].value == strTrade){
	    document.all.trade.options[i].selected = true;
	}
}

var strArea = "<% = cArea %>";
for(var i=0;i<document.all.area.options.length;i++){
    if(document.all.area.options[i].value == strArea){
	    document.all.area.options[i].selected = true;
	}
}

var strSquare = "<% = cSquare %>";
for(var i=0;i<document.all.square.options.length;i++){
    if(document.all.square.options[i].value == strSquare){
	    document.all.square.options[i].selected = true;
	}
}
-->
</script>
<%
    End If
End Sub

Sub saveEditData()
    Dim cId
	cId = CInt(ABS(Request("id")))
	If Not IsNumeric(cId) Or cId <= 0 Then
	    Response.Write("<font color=""#FF0000""><b>�����������</b></font><br><br>")
		Response.Write("<input type=""button"" value="" �� �� "" onClick=""location.replace('listAll.asp');"">")
	Else
        Dim cCompany,cLinkman,cHomepage,cTel,cEmail,cUser
    	Dim cArea,cLocal,cType,cStatus,cAddress
		Dim cTrade,cSquare,cGroup
    	'Dim cDomainBegin,cDomainEnd
    	'Dim cSpaceBegin,cSpaceEnd
    	'Dim cOprationBegin,cOprationEnd
    	Dim cInfo
    	cCompany = Trim(Request("company"))
    	cLinkman = Trim(Request("linkman"))
    	cHomepage = Trim(Request("homepage"))
	    If cHomepage = "" Then cHomepage = "&nbsp;"
    	cTel = Trim(Request("tel"))
    	cEmail = Trim(Request("email"))
    	cUser = Trim(Request("user"))
    	cArea = Trim(Request("area"))
    	cLocal = Trim(Request("local"))
    	cType = Trim(Request("type"))
    	cStatus = Trim(Request("status"))
		cAddress = Trim(Request("address"))
		cTrade = Trim(Request("trade"))
		cSquare = Trim(Request("square"))
		cGroup = CInt(Abs(getGroup(Request("group"))))
    	'cDomainBegin = Trim(Request("domainBegin"))
    	'cDomainEnd = Trim(Request("domainEnd"))
    	'cSpaceBegin = Trim(Request("spaceBegin"))
    	'cSpaceEnd = Trim(Request("spaceEnd"))
    	'cOprationBegin = Trim(Request("oprationBegin"))
    	'cOprationEnd = Trim(Request("oprationEnd"))
    	cInfo = htmlEncode2(Request("info"))
    	Dim rs
    	Set rs = Server.CreateObject("ADODB.Recordset")
	    rs.Open "Select * From baidu_client Where (cCompany = '" & cCompany & "' Or cLinkman = '" & cLinkman & "') And cId <> " & cId,conn,3,1
	    If rs.RecordCount > 0 Then
	        Response.Write("<font color=""#FF0000"">�˿ͻ��Ѿ����ڡ�</font><br><br>")
		    Response.Write("<input type=""button"" value="" �� �� "" onClick=""history.back();"">")
		    Response.End()
	    End If
	    rs.Close
	    ''
    	rs.Open "Select Top 1 * From baidu_client Where cId = " & cId,conn,3,2
		If rs.RecordCount <> 1 Then		
	        Response.Write("<font color=""#FF0000""><b>�����������</b></font><br><br>")
		    Response.Write("<input type=""button"" value="" �� �� "" onClick=""location.replace('listAll.asp');"">")
			Response.End()
		End If
	    rs("cCompany") = cCompany
	    rs("cLinkman") = cLinkman
	    rs("cHomepage") = cHomepage
	    rs("cTel") = cTel
	    rs("cEmail") = cEmail
	    rs("cUser") = cUser
    	rs("cArea") = cArea
    	rs("cLocal") = cLocal
    	rs("cType") = cType
    	rs("cStatus") = cStatus
		rs("cAddress") = cAddress
		rs("cTrade") = cTrade
		rs("cSquare") = cSquare
		rs("cGroup") = cGroup
    	'If cDomainBegin <> "" Then rs("cDomainBegin") = cDomainBegin
    	'If cDomainEnd <> "" Then rs("cDomainEnd") = cDomainEnd
    	'If cSpaceBegin <> "" Then rs("cSpaceBegin") = cSpaceBegin
	    'If cSpaceEnd <> "" Then rs("cSpaceEnd") = cSpaceEnd
    	'If cOprationBegin <> "" Then rs("cOprationBegin") = cOprationBegin
    	'If cOprationEnd <> "" Then rs("cOprationEnd") = cOprationEnd
    	rs("cInfo") = cInfo
    	rs.Update
    	rs.Close
    	Set rs = Nothing
    	Response.Redirect("view.asp?cId=" & cId)
	End If
End Sub
%>
    </td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0053)http://bbs.wj8.net/admincp.php?action=menu&sid=LXPP7l -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>A:link {
	COLOR: #003366; TEXT-DECORATION: none
}
A:visited {
	COLOR: #003366; TEXT-DECORATION: none
}
A:hover {
	TEXT-DECORATION: underline
}
BODY {
	FONT-SIZE: 12px; SCROLLBAR-ARROW-COLOR: #dde3ec; SCROLLBAR-BASE-COLOR: #f8f9fc; BACKGROUND-COLOR: #e9edf7
}
TABLE {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
TEXTAREA {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
INPUT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
OBJECT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
SELECT {
	FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #000000; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #f8f9fc
}
.nav {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-FAMILY: Tahoma, Verdana
}
.header {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/headerbg.gif); COLOR: #ffffff; FONT-FAMILY: Tahoma, Verdana
}
.category {
	FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/catbg.gif); COLOR: #000000; FONT-FAMILY: Tahoma
}
.multi {
	FONT-SIZE: 11px; COLOR: #003366; FONT-FAMILY: Tahoma
}
.smalltxt {
	FONT-SIZE: 11px; FONT-FAMILY: Tahoma
}
.mediumtxt {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
.bold {
	FONT-WEIGHT: bold
}
BLOCKQUOTE {
	BORDER-RIGHT: #dde3ec 1px dashed; PADDING-RIGHT: 5px; BORDER-TOP: #dde3ec 1px dashed; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; BORDER-LEFT: #dde3ec 1px dashed; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BORDER-BOTTOM: #dde3ec 1px dashed; BACKGROUND-COLOR: #ffffff
}
.code {
	PADDING-RIGHT: 5px; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BACKGROUND-COLOR: #ffffff
}
</STYLE>

<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY leftMargin=3 topMargin=3>