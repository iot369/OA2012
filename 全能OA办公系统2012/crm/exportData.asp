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

Dim strCounter,strToPrint

Dim dataItem,beginData,endData
dataItem = Trim(Request("dataItem"))
beginData = Trim(Request("beginData"))
endData = Trim(Request("endData"))

If beginData = "��ʼ����" Then beginData = ""
If endData = "��������" Then endData = ""

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

Function getItem(dataItem)
    If dataItem = "" Then
	    getItem = ""
	Else
	    Select Case dataItem
		Case "cEmail"
		    getItem = "��������"
		Case Else
		    getItem = ""
		End Select
	End If
End Function

Dim flag
flag = 0
If dataItem <> "" Then
    strToPrint = strToPrint & "        <tr>" & VBCrlf
    strToPrint = strToPrint & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
    strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
    strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
    strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">" & getItem(dataItem) & "</td>" & VBCrlf
    If Session("CRM_level") = 9 Then
        strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
    End If
    strToPrint = strToPrint & "        </tr>" & VBCrlf
	
    Dim fso,f,fl
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fl = Server.MapPath("data.CSV")
	If fso.FileExists(fl) Then
	    fso.DeleteFile(fl)
	    Set f = fso.CreateTextFile(fl)
		f.WriteLine("����,�����ʼ���ַ,foxaddrID")
	Else
	    Set f = fso.CreateTextFile(fl)
		f.WriteLine("����,�����ʼ���ַ,foxaddrID")
	End If
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
	
    Select Case dataItem
    Case "cEmail"
        If beginData = "" And endData = "" Then
    	    If Session("CRM_level") = 9 Then
    	        rs.Open "Select * From baidu_client Order By cId Desc",conn,3,1
    		Else
    		    rs.Open "Select * From baidu_client Where cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
    		End If
    	Else
    	    If beginData = "" Then beginData = endData		
      	    If Session("CRM_level") = 9 Then
    	        rs.Open "Select * From baidu_client Where cEmail Like '%" & beginData & "%' Order By cId Desc",conn,3,1
    		Else
    		    rs.Open "Select * From baidu_client Where cEmail Like '%" & beginData & "%' And cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
    		End If
    	End If
    Case Else
    End Select
	
    Do While Not rs.BOF And Not rs.EOF
        strToPrint = strToPrint & "        <tr>" & VBCrlf
        strToPrint = strToPrint & "          <td width=""60"" align=""center"">" & rs("cId") & "</td>" & VBCrlf
        strToPrint = strToPrint & "          <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
        strToPrint = strToPrint & "          <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" & VBCrlf
        strToPrint = strToPrint & "          <td>" & rs(dataItem) & "</td>" & VBCrlf
    	If Session("CRM_level") = 9 Then
            strToPrint = strToPrint & "          <td>" & rs("cUser") & "</td>" & VBCrlf
    	End If
		f.WriteLine(rs("cLinkman") & "," & rs(dataItem) & "," & rs("cId"))
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	f.CLose
	Set f = Nothing
	Set fso = Nothing
	flag = 1
End If
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
    if (document.exportForm.dataItem.value == ""){
	    alert("��ѡ��Ҫ�������������ࡣ");
		document.exportForm.dataItem.focus();
		return false;
	}
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
      <tr>
        <td bgcolor="#88ADDF">&nbsp;</td>
      </tr>
    </table>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="exportForm" action="?" method="post" onSubmit="return checkInput();">
		<tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>��ѡ�񵼳�������Ŀ��
            <select name="dataItem" id="dataItem">
              <option value="">��ѡ��</option>
              <option value="cEmail">��������</option>
            </select>
            <br>
            ��ѡ�񵼳����ݷ�Χ�� 
            <input name="beginData" type="text" id="beginData" value="��ʼ����" size="16" maxlength="36" onFocus="this.value='';">
            -
            <input name="endData" type="text" id="endData" value="��������" size="16" maxlength="36" onFocus="this.value='';">
            <input type="submit" name="Submit" value=" �� �� ">
            <hr size="1" noshade>
            <span class="emRed">˵����</span><br>
            &nbsp;&nbsp;&nbsp;&nbsp;�������ݷ�Χ�����԰������޶�һ�������б�Χ������Ҫ�����������ַ���abc���ĵ������䣬�ڡ���ʼ���ݡ������롰abc�����ɣ����Ҫ����2002��12��ע��������������ڡ���ʼ���ݡ������롰2002-12-01�����ڡ��������ݡ������롰2002-12-31����<br>
            &nbsp;&nbsp;&nbsp;&nbsp;�������ʼ���ݡ��͡��������ݡ������գ�������ȫ�����ݡ�</td>
        </tr>
		</form>
      </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="����ͷ��" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="����ͷ��" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;">
      <% = strCounter %> 
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF"><% = strToPrint %>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
<%
If flag = 1 Then
    Response.Write("<script>window.open('downFile.asp?file=data.CSV','','');</script>")
End If
%>
</body>
</html>
