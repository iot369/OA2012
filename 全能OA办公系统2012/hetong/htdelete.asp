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

' Load Key Parameters
sKey = Request.querystring("key")
If sKey = "" Or IsNull(sKey) Then
	sKey = Request.Form("key_d")
End If
arRecKey = Split(sKey&"", ",")

' Single delete record
If sKey = "" Or IsNull(sKey) Then Response.Redirect "htlist.asp"
sDbWhere = sDbWhere & "[ID]=" & AdjustSql(Trim(sKey)) & ""

' Get action
sAction = Request.Form("a_delete")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
End If

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Display
		If LoadRecordCount(sDbWhere) <= 0 Then
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "htlist.asp"
		End If
	Case "D": ' Delete
		If DeleteData(sDbWhere) Then
			Session("ewmsg") = "Delete Successful For Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "htlist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p align="center"><span class="aspmaker"><br>
  <br><a href="htlist.asp">返回上一页</a></span></p>
<form action="htdelete.asp" method="post">
<p>
<input type="hidden" name="a_delete" value="D">
<input type="hidden" name="key_d" value="<%= sKey %>">
<table border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="6FECFF">
	<tr bgcolor="D8F7FF">
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">ID</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">合同号</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">客户名称</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">产品型号</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">数量</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">价格</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">成交金额</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">月度</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">交期</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">承办人</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">备注</span></div></td>
		<td valign="top"><div align="center" style="color: #0d79b3"><span class="aspmaker" style="">业务提成</span></div></td>
	</tr>
<%
nRecCount = 0
For Each sRecKey In arRecKey
	sRecKey = Trim(sRecKey)
	nRecCount = nRecCount + 1

	' Set row color
	sItemRowClass = " bgcolor=""#FFFFFF"""

	' Display alternate color for rows
	If nRecCount Mod 2 <> 0 Then
		sItemRowClass = " bgcolor=""#F5F5F5"""
	End If
	If LoadData(sRecKey) Then
%>
	<tr bgcolor="#FFFFFF"<%=sItemRowClass%>>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_ID %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_5408540C53F7 %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_5BA26237540D79F0 %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_4EA754C1578B53F7 %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_657091CF %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_4EF7683C %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_91D1989D %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_67085EA6 %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_4EA4671F %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_627F529E %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_59076CE8 %>
        </span></div></td>
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_53CD998862A5916C %>
        </span></div></td>
	</tr>
<%
	End If
Next
%>
</table>
<p align="center">
<input type="submit" name="Action" value="确认删除此条记录">
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
' Function LoadRecordCount
' - Load Record Count based on input sql criteria sqlKey

Function LoadRecordCount(sqlKey)
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sSql = "SELECT * FROM [ht]"
	sSql = sSql & " WHERE " & sqlKey
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
	LoadRecordCount = rs.RecordCount
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function DeleteData
' - Delete Records based on input sql criteria sqlKey

Function DeleteData(sqlKey)
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sSql = "SELECT * FROM [ht]"
	sSql = sSql & " WHERE " & sqlKey
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
	Do While Not rs.Eof
		rs.Delete
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	DeleteData = True
End Function
%>
