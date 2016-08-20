<%@ CodePage = 936 LCID = 2052 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<%
ewCurSec = 0 ' Initialise

' User levels
Const ewAllowAdd = 1
Const ewAllowDelete = 2
Const ewAllowEdit = 4
Const ewAllowView = 8
Const ewAllowList = 8
Const ewAllowReport = 8
Const ewAllowSearch = 8
Const ewAllowAdmin = 16
%>
<%

' Initialize common variables
x_ID = Null
x_5408540C53F7 = Null
x_5BA26237540D79F0 = Null
x_4EA754C1578B53F7 = Null
x_657091CF = Null
x_4EF7683C = Null
x_91D1989D = Null
x_67085EA6 = Null
x_4EA4671F = Null
x_627F529E = Null
x_59076CE8 = Null
x_53CD998862A5916C = Null
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Buffer = True
sKey = Request.Querystring("key")
If sKey = "" Or IsNull(sKey) Then sKey = Request.Form("key")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
Else

	' Get fields from form
	x_ID = Request.Form("x_ID")
	x_5408540C53F7 = Request.Form("x_5408540C53F7")
	x_5BA26237540D79F0 = Request.Form("x_5BA26237540D79F0")
	x_4EA754C1578B53F7 = Request.Form("x_4EA754C1578B53F7")
	x_657091CF = Request.Form("x_657091CF")
	x_4EF7683C = Request.Form("x_4EF7683C")
	x_91D1989D = Request.Form("x_91D1989D")
	x_67085EA6 = Request.Form("x_67085EA6")
	x_4EA4671F = Request.Form("x_4EA4671F")
	x_627F529E = Request.Form("x_627F529E")
	x_59076CE8 = Request.Form("x_59076CE8")
	x_53CD998862A5916C = Request.Form("x_53CD998862A5916C")
End If
If sKey = "" Or IsNull(sKey) Then Response.Redirect "htlist.asp"

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Get a record to display
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "htlist.asp"
		End If
	Case "U": ' Update
		If EditData(sKey) Then ' Update Record based on key
			Session("ewmsg") = "Update Record Successful for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "htlist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<script type="text/javascript" src="ew.js"></script>
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
if (EW_this.x_91D1989D && !EW_checknumber(EW_this.x_91D1989D.value)) {
	if (!EW_onError(EW_this, EW_this.x_91D1989D, "TEXT", "Incorrect floating point number - 金额"))
		return false; 
}
if (EW_this.x_67085EA6 && !EW_checkinteger(EW_this.x_67085EA6.value)) {
	if (!EW_onError(EW_this, EW_this.x_67085EA6, "TEXT", "Incorrect integer - 月度"))
		return false; 
}
return true;
}
//-->
</script>
<form name="htedit" id="htedit" action="htedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_edit" value="U">
<input type="hidden" name="key" value="<%= sKey %>">
<table width="540" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="B0C8EA">
	
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">合同号</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5408540C53F7" id="x_5408540C53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5408540C53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">客户名称</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5BA26237540D79F0" id="x_5BA26237540D79F0" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5BA26237540D79F0&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品型号</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA754C1578B53F7" id="x_4EA754C1578B53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EA754C1578B53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品数量</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_657091CF" id="x_657091CF" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_657091CF&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品价格</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EF7683C" id="x_4EF7683C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EF7683C&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">成交金额</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_91D1989D" id="x_91D1989D" size="30" value="<%= Server.HTMLEncode(x_91D1989D&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">月度</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_67085EA6" id="x_67085EA6" size="30" value="<%= Server.HTMLEncode(x_67085EA6&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">交期</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA4671F" id="x_4EA4671F" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EA4671F&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">承办业务员</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_627F529E" id="x_627F529E" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_627F529E&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">备注</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_59076CE8" id="x_59076CE8" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_59076CE8&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">业务提成</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_53CD998862A5916C" id="x_53CD998862A5916C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_53CD998862A5916C&"") %>">
</span></td>
	</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="修改">
</form>
<%
conn.Close ' Close Connection
Set conn = Nothing
%>
<%

'-------------------------------------------------------------------------------
' Function LoadData
' - Load Data based on Key Value sKey
' - Variables setup: field variables

Function LoadData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [ht]"
	sSql = sSql & " WHERE [ID] = " & sKeyWrk
	sGroupBy = ""
	sHaving = ""
	sOrderBy = ""
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If	
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If	
	If sOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sOrderBy
	End If	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		x_ID = rs("ID")
		x_5408540C53F7 = rs("合同号")
		x_5BA26237540D79F0 = rs("客户名称")
		x_4EA754C1578B53F7 = rs("产品型号")
		x_657091CF = rs("数量")
		x_4EF7683C = rs("价格")
		x_91D1989D = rs("金额")
		x_67085EA6 = rs("月度")
		x_4EA4671F = rs("交期")
		x_627F529E = rs("承办")
		x_59076CE8 = rs("备注")
		x_53CD998862A5916C = rs("反馈报酬")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function EditData
' - Edit Data based on Key Value sKey
' - Variables used: field variables

Function EditData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Open record
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [ht]"
	sSql = sSql & " WHERE [ID] = " & sKeyWrk
	sGroupBy = ""
	sHaving = ""
	sOrderBy = ""
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If	
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If	
	If sOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sOrderBy
	End If	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If rs.Eof Then
		EditData = False ' Update Failed
	Else

		' Field ID
		' Field 合同号

		sTmp = Trim(x_5408540C53F7)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("合同号") = sTmp

		' Field 客户名称
		sTmp = Trim(x_5BA26237540D79F0)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("客户名称") = sTmp

		' Field 产品型号
		sTmp = Trim(x_4EA754C1578B53F7)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("产品型号") = sTmp

		' Field 数量
		sTmp = Trim(x_657091CF)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("数量") = sTmp

		' Field 价格
		sTmp = Trim(x_4EF7683C)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("价格") = sTmp

		' Field 金额
		sTmp = x_91D1989D
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = cDbl(sTmp)
		End If
		rs("金额") = sTmp

		' Field 月度
		sTmp = x_67085EA6
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("月度") = sTmp

		' Field 交期
		sTmp = Trim(x_4EA4671F)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("交期") = sTmp

		' Field 承办
		sTmp = Trim(x_627F529E)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("承办") = sTmp

		' Field 备注
		sTmp = Trim(x_59076CE8)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("备注") = sTmp

		' Field 反馈报酬
		sTmp = Trim(x_53CD998862A5916C)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("反馈报酬") = sTmp
		rs.Update
		EditData = True ' Update Successful
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
