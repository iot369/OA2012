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

Rem Session("CRM_account") �û��ʺ�
Rem Session("CRM_name") �û���
Rem Session("CRM_level") �û��ȼ�

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

If Session("CRM_level") <> 9 Then Response.Redirect("listAll.asp")

Function getGroupName(gId)
    If Not IsNumeric(gId) Or gId <= 0 Then
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

Function list()
    Dim strToPrint
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From baidu_clientsType",conn,3,1
    Do While Not rs.BOF And Not rs.EOF
    	strToPrint = strToPrint & "        <tr>" & VBCrlf
    '	strToPrint = strToPrint & "          <td align=""center"">" & rs("lId") & "</td>" & VBCrlf
    	strToPrint = strToPrint & "          <td>" & rs("clientsType") & "</td>" & VBCrlf
    	strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&clientsTypeOld=" & rs("clientsType") & """>�޸�</a>] [<a href=""?action=delete&clientsTypeOld=" & rs("clientsType") & """ onClick=""return confirm('ȷ��Ҫɾ���ÿͻ����ͺ͸�\r�����µ����пͻ����ϣ�');"">ɾ��</a>]</td>" & VBCrlf
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	list = strToPrint
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="http://www.web87.9126.com">
<meta name="Date" content="2003-08">
<title>���۹���ϵͳ</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<style type="text/css">
.style7 {color: #2d4865}
.style8 {color: #0d79b3;
	font-weight: bold;
}
</style>
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

function selectSystem(j)
{
    var num = 4;
	for(var i=1;i<=num;i++){
	    if(i != j){
		    document.all["block" + i].style.display = "none";
			document.all["block" + (i + 4)].style.display = "none";
		}
		else{
		    document.all["block" + i].style.display = "block";
		    document.all["block" + (i + 4)].style.display = "block";
		}
	}
}
-->
</script>
</head>

<body style="background-color: menu;" topmargin="0" leftmargin="0">
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
<table width="550" height="680" border="0" align="center" cellpadding="0" cellspacing="0">
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
        <tr id="block1" style="display: block;"> 
          <td>
<%
Dim action
action = Trim(Request.QueryString("action"))
Select Case action
Case "add"
    Call addOrEdit()
Case "save"
    Call saveData()
Case "edit"
    Call addOrEdit()
Case "restore"
    Call restore()
Case "delete"
    Call deleteData()
Case Else
    Call addOrEdit()
End Select

Sub saveData()
    Dim clientsType
	clientsType = Trim(Request.Form("clientsType"))
	If clientsType = "" Then
	    Response.Write("<div align=""center"">�ύ�����ݲ��������뷵��������д��<br>")
		Response.Write("<input name=""back"" type=""button"" value="" �� �� "" onClick=""history.back();""></div>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From baidu_clientsType Where clientsType = '" & clientsType & "'",conn,3,2
	If rs.RecordCount > 0 Then
	    Response.Write("<div align=""center"">�ÿͻ������Ѿ����ڡ�<br>")
		Response.Write("<input name=""back"" type=""button"" value="" �� �� "" onClick=""history.back();""></div>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs("clientsType") = clientsType
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?")
	End If
End Sub

Sub restore()
    Dim clientsTypeOld,clientsType
	clientsTypeOld = Trim(Request.Form("clientsTypeOld"))
	clientsType = Trim(Request.Form("clientsType"))
	If clientsTypeOld = "" Or clientsType = "" Then
	    Response.Write("<div align=""center"">�ύ�����ݲ��������뷵��������д��<br>")
		Response.Write("<input name=""back"" type=""button"" value="" �� �� "" onClick=""history.back();""></div>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From baidu_clientsType Where clientsType <> '" & clientsTypeOld & "'",conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If rs("clientsType") = clientsType Then
	        Response.Write("<div align=""center"">�����ƻ����Ѿ����ڡ�<br>")
		    Response.Write("<input name=""back"" type=""button"" value="" �� �� "" onClick=""history.back();""></div>")
		    rs.Close
		    Set rs = Nothing
		    Exit Sub
		End If
		rs.MoveNext
	Loop
	rs.Close
	
	rs.Open "Select * From baidu_clientsType Where clientsType = '" & clientsTypeOld & "'",conn,3,2
	If rs.RecordCount = 1 Then
	    clientsTypeOld = rs("clientsType")
		rs("clientsType") = clientsType
		rs.Update
		If clientsTypeOld <> clientsType Then
		    Dim rss
			Set rss = Server.CreateObject("ADODB.Recordset")
			rss.Open "Select * From baidu_client Where cType = '" & clientsTypeOld & "'",conn,3,2
			Do While Not rss.BOF And Not rss.EOF
			    rss("cType") = clientsType
				rss.Update
				rss.MoveNext
			Loop
			rss.Close
			Set rss = Nothing
		End If
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub deleteData()
    Dim clientsTypeOld,lIdOld
	clientsTypeOld = Trim(Request("clientsTypeOld"))
	If clientsTypeOld = "" Then
	    Response.Write("<div align=""center"">�ύ�����ݲ��������뷵��������д��<br>")
		Response.Write("<input name=""back"" type=""button"" value="" �� �� "" onClick=""history.back();""></div>")
		Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_clientsType Where clientsType = '" & clientsTypeOld & "'",conn,3,2
	If rs.RecordCount > 0 Then
		rs.Delete
		rs.Update
	End If
	rs.Close
	
	If clientsTypeOld <> "" Then
	    rs.Open "Select * From baidu_client Where cType = '" & clientsTypeOld & "'",conn,3,2
		Do While Not rs.BOF And Not rs.EOF
		    rs.Delete
			rs.Update
			rs.MoveNext
		Loop
		rs.Close
	End If
	Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub addOrEdit()
    Dim clientsType,clientsTypeOld,strOut,strAction
	If action = "edit" Then
	    Dim rs
		clientsTypeOld = Trim(Request("clientsTypeOld"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_clientsType Where clientsType = '" & clientsTypeOld & "'",conn,3,1
		If rs.RecordCount = 1 Then
			clientsType = rs("clientsType")
		End If
		rs.Close
		Set rs = Nothing
		strOut = "�༭�ͻ����ͣ�"
		strAction = "?action=restore"
	Else
	    strOut = "��ӿͻ����ͣ�"
		strAction = "?action=save"
	End If		    
%>
		    <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <form name="clientsTypeForm" action="<% = strAction %>" method="post" onSubmit="return checkInput('clientsTypeForm');">
			  <tr> 
                <td width="60">&nbsp;</td>
                <td><% = strOut %>
                    <% If action = "edit" Then %><input name="clientsTypeOld" type="hidden" id="clientsTypeOld" value="<% = clientsTypeOld %>"><% End If %></td>
              </tr>
              <tr> 
                <td width="60">&nbsp;</td>
                  <td>�ͻ��������ƣ� 
                    <input name="clientsType" type="text" id="clientsType" size="16" maxlength="16" value="<% = clientsType %>" selfValue="�ͻ��������">
                  </td>
              </tr>
              <tr> 
                <td width="60" align="center">&nbsp;</td>
                <td align="center"> <input type="submit" name="Submit" value=" �� �� "> 
                  &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "> 
                </td>
              </tr>
			  </form>
            </table>
<%
End Sub
%>
		  </td>
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
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="40"><div align="center">[<a href="system_level.asp">�û�����</a>] [<a href="system_group.asp">�� �� ��</a> ] [<a href="system_clientsType.asp">�ͻ��ȼ�</a>] [<a href="system_clientsTrade.asp">��ҵ����</a>] [<a href="system_recordsType.asp">�ݷ�����</a>]<br>
[<a href="system_area.asp">ҵ������</a>] [<a href="system_square.asp">ҵ��С��</a>] [<a href="system_del1.asp">�ͻ�ɾ��</a>] [<a href="system_del2.asp">��¼ɾ��</a>] [<a href="system_del3.asp">�ݷ�ɾ��</a>] </div></td>
        </tr>
      </table>
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr> 
          <td align="center" bgcolor="menu">�ͻ���������</td>
          <td width="120" align="center" bgcolor="menu">����</td>
          <% = list() %>
        </tr>
      </table> </td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
