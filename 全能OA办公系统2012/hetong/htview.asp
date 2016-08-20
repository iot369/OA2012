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
If sKey = "" Or IsNull(sKey) Then Response.Redirect "htlist.asp"

' Get action
sAction = Request.Form("a_view")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
End If

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
End Select
%>
<!--#include file="header.asp"-->
<script type="text/javascript" src="ew.js"></script>
<p align="center"><span class="aspmaker"><br>
  <br>
<a href="htlist.asp">返回上一页</a>&nbsp;
<a href="<%= "htedit.asp?key=" & Server.URLEncode(sKey) %>">编辑此条记录</a>&nbsp;
<a href="<%= "htadd.asp?key=" & Server.URLEncode(sKey) %>">复制此条记录</a>&nbsp;
<a href="<%= "htdelete.asp?key=" & Server.URLEncode(sKey) %>">删除此条记录</a>&nbsp;
</span></p>
<p>
<form>
<table width="540" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="B0C8EA">
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">ID</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_ID %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">合同号</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_5408540C53F7 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">客户名称</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_5BA26237540D79F0 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品型号</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EA754C1578B53F7 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品数量</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_657091CF %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">产品价格</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EF7683C %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">成交金额</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_91D1989D %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">月度</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_67085EA6 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">交期</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EA4671F %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">承办业务员</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_627F529E %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">备注</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_59076CE8 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">业务提成</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_53CD998862A5916C %>
</span></td>
	</tr>
</table>
</form>
<p>
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
