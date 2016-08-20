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

Dim action
action = Trim(Request("action"))
If action = "killSession" Then Session("CRM_sql") = ""

Dim strCounter,strToPrint

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

Dim subAction
subAction = Trim(Request("subAction"))

If subAction = "searchItem" Then
    Dim cCompany,cLinkman,cHomepage
	Dim cTel,cEmail,cAddress
	Dim cArea,cSquare,cType
	Dim cTrade,cUser,cGroup
	Dim arrUser
	cCompany = Trim(Request("company"))
	cLinkman = Trim(Request("linkman"))
	cHomepage = Trim(Request("homepage"))
	cTel = Trim(Request("tel"))
	cEmail = Trim(Request("email"))
	cAddress = Trim(Request("address"))
	cArea = Trim(Request("area"))
	cSquare = Trim(Request("square"))
	cType = Trim(Request("type"))
	cTrade = Trim(Request("trade"))
	cUser = Trim(Request("user"))
	cGroup = Request("group")
	If cGroup <> "" Then
	    cGroup = CInt(Abs(cGroup))
	End If
	
	Dim sql
    sql = ""
    If cCompany <> "" Then
        If sql = "" Then
            sql = sql & " Where cCompany Like '%" & cCompany & "%'"
	    Else
	        sql = sql & " And cCompany Like '%" & cCompany & "%'"
        End If
	End If
	
	
    If cLinkman <> "" Then
        If sql = "" Then
            sql = sql & " Where cLinkman Like '%" & cLinkman & "%'"
	    Else
	        sql = sql & " And cLinkman Like '%" & cLinkman & "%'"
        End If
	End If
	
	
    If cHomepage <> "" Then
        If sql = "" Then
            sql = sql & " Where cHomepage Like '%" & cHomepage & "%'"
	    Else
	        sql = sql & " And cHomepage Like '%" & cHomepage & "%'"
        End If
	End If
	
	If cTel <> "" Then
        If sql = "" Then
            sql = sql & " Where cTel Like '%" & cTel & "%'"
	    Else
	        sql = sql & " And cTel Like '%" & cTel & "%'"
        End If
	End If
	
    If cEmail <> "" Then
        If sql = "" Then
            sql = sql & " Where cEmail Like '%" & cEmail & "%'"
	    Else
	        sql = sql & " And cEmail Like '%" & cEmail & "%'"
        End If
	End If
	
    If cAddress <> "" Then
        If sql = "" Then
            sql = sql & " Where cAddress Like '%" & cAddress & "%'"
	    Else
	        sql = sql & " And cAddress Like '%" & cAddress & "%'"
        End If
	End If
	
    If cArea <> "" Then
        If sql = "" Then
            sql = sql & " Where cArea = '" & cArea & "'"
	    Else
	        sql = sql & " And cArea = '" & cArea & "'"
        End If
	End If
	
    If cSquare <> "" Then
        If sql = "" Then
            sql = sql & " Where cSquare = '" & cSquare & "'"
	    Else
	        sql = sql & " And cSquare = '" & cSquare & "'"
        End If
	End If
	
    If cType <> "" Then
        If sql = "" Then
            sql = sql & " Where cType = '" & cType & "'"
	    Else
	        sql = sql & " And cType = '" & cType & "'"
        End If
	End If
	
    If cTrade <> "" Then
        If sql = "" Then
            sql = sql & " Where cTrade = '" & cTrade & "'"
	    Else
	        sql = sql & " And cTrade = '" & cTrade & "'"
        End If
	End If
		
	If cGroup <> "" And IsNumeric(cGroup) Then
        If sql = "" Then
            sql = sql & " Where cGroup = " & cGroup
	    Else
	        sql = sql & " And cGroup = " & cGroup
        End If
	End If
	
	If Session("CRM_level") < 9 Then
        If cUser <> "" Then
    	    arrUser = Split(getUserList(Session("CRM_level"),Session("CRM_group")),",")
	        Dim k,flag
	        flag = 0
    	    For k = 0 To UBound(arrUser) - 1
    	        If Replace(arrUser(k),"'","") = cUser Then
    		        flag = 1
    		    	Exit For
    		    End If
    	    Next
            If flag = 1 Then
                If sql = "" Then
                    sql = sql & " Where cUser = '" & cUser & "'"
    	        Else
    	            sql = sql & " And cUser = '" & cUser & "'"
                End If
    		Else
    		    If sql = "" Then
                    sql = sql & " Where cUser = 'Ȩ�޲�������û�'"
    	        Else
    	            sql = sql & " And cUser = 'Ȩ�޲�������û�'"
                End If
    	    End If
    	Else
    	    If sql = "" Then
                sql = sql & " Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")"
    	    Else
    	        sql = sql & " And cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")"
            End If
    	End If
	Else
	    If cUser <> "" Then
	        If sql = "" Then
                sql = sql & " Where cUser = '" & cUser & "'"
        	Else
    	        sql = sql & " And cUser = '" & cUser & "'"
            End If
		End If
	End If
End If

If cCompany = "" And cHomepage = "" And cLinkman = "" And cTel = "" And cEmail = "" And cAddress = "" And cArea = "" And cSquare = "" And cType = "" And cTrade = "" And cUser = "" And cGroup = "" Then
    If Session("CRM_sql") <> "" Then
        sql = Session("CRM_sql")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")"
		End If
	End If
Else
    Session("CRM_sql") = sql
End If

'If Session("CRM_level") < 9 Then
'    If sql = "" Then
'	    sql = sql & " Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")"
'	Else
'	    sql = sql & " And cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")"
'	End If
'End If

strToPrint = strToPrint & "        <tr>" & VBCrlf
strToPrint = strToPrint & "          <td width=""60"" align=""center"" bgcolor=""menu"">���</td>" & VBCrlf
strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾����</td>" & VBCrlf
strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">��˾��ַ</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">�����ʼ�</td>" & VBCrlf
'If Session("CRM_level") > 1 Then
    strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">�ͻ��ȼ�</td>" & VBCrlf
    strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
'End If
'If Session("CRM_level") = 9 Then
'    strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">����С��</td>" & VBCrlf
'End If
strToPrint = strToPrint & "        </tr>" & VBCrlf

Dim rs,intTotalRecords,intTotalPages,intCurrentPage,intPageSize
intCurrentPage = CInt(ABS(Request("pageNum")))
If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
intPageSize = 10

Set rs = Server.CreateObject("ADODB.Recordset")
'Response.Write(sql)
'Response.End()
rs.Open "Select * From baidu_client" & sql & " Order By cId",conn,3,1
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
    strToPrint = strToPrint & "          <td width=""60"" align=""center"">" & rs("cId") & "</td>" & VBCrlf
    strToPrint = strToPrint & "          <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
    strToPrint = strToPrint & "          <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" & VBCrlf
    strToPrint = strToPrint & "          <td>" & rs("cEmail") & "</td>" & VBCrlf
    'If Session("CRM_level") > 1 Then
	    strToPrint = strToPrint & "          <td>" & rs("cType") & "</td>" & VBCrlf
        strToPrint = strToPrint & "          <td>" & rs("cUser") & "</td>" & VBCrlf
    'End If
    'If Session("CRM_level") = 9 Then
    '    strToPrint = strToPrint & "          <td>" & rs("cLocal") & "</td>" & VBCrlf
    'End If
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" >
<meta name="Date">
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
-->
</script>
<style type="text/css">
.style7 {color: #2d4865}
.style8 {color: #0d79b3;
	font-weight: bold;
}
</style>
</head>

<body  topmargin="0" leftmargin="0" onCopy="return false;" onSelectStart="return false;" onCut="return false;">
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
          <td>����д������Ŀ<br>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <form name="searchForm" action="?subAction=searchItem" method="post">
                <tr> 
                  <td width="77" height="25" align="right">��˾���ƣ�</td>
                  <td width="144"><input name="company" type="text" id="company" size="24" maxlength="36"></td>
                  <td width="78" align="right">��ϵ�ˣ�</td>
                  <td width="191"> <input name="linkman" type="text" id="linkman" size="12" maxlength="16"></td>
                </tr>
                <tr> 
                  <td width="77" height="25" align="right">��˾��ַ��</td>
                  <td><input name="homepage" type="text" id="homepage2" size="24" maxlength="36"></td>
                  <td width="78" align="right">��ϵ�绰��</td>
                  <td><input name="tel" type="text" id="tel" size="24" maxlength="36"></td>
                </tr>
                <tr> 
                  <td width="77" height="25" align="right">�������䣺</td>
                  <td><input name="email" type="text" id="email" size="24" maxlength="36"></td>
                  <td width="78" align="right">��˾��ַ��</td>
                  <td><input name="address" type="text" id="address" size="36"></td>
                </tr>
                <tr> 
                  <td height="25" align="right">������</td>
                  <td> <% = getList(1,"baidu_area","","areaName","area","ҵ������") %> </td>
                  <td align="right">����С����</td>
                  <td> <% = getList(1,"baidu_square","","squareName","square","ҵ��С��") %> </td>
                </tr>
                <tr> 
                  <td height="25" align="right">�ͻ��ȼ���</td>
                  <td> <% = getList(1,"baidu_clientsType","","clientsType","type","�ͻ�����") %> </td>
                  <td align="right">��ҵ���ͣ�</td>
                  <td> <% = getList(1,"baidu_clientsTrade","","clientsTrade","trade","��ҵ����") %> </td>
                </tr>
                <tr> 
                  <td height="25" align="right">ҵ��Ա��</td>
                  <td> <input name="user" type="text" id="user" size="12" maxlength="16"></td>
                  <td align="right">ҵ���飺</td>
                  <td> <% = getList(2,"baidu_group","gId","gName","group","ҵ����") %> </td>
                </tr>
                <tr align="center"> 
                  <td colspan="4"><input type="submit" name="Submit" value=" �� �� "> 
                    &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" �� �� "></td>
                </tr>
              </form>
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
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><% = strCounter %></td>
          <td align="right">[<a href="?action=killSession">����ȫ���б�</a>]</td>
        </tr>
      </table>
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF"><% = strToPrint %>
	  </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><% Response.Write(Session("CRM_sql")) %><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
