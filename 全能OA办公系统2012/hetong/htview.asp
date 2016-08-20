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
<a href="htlist.asp">������һҳ</a>&nbsp;
<a href="<%= "htedit.asp?key=" & Server.URLEncode(sKey) %>">�༭������¼</a>&nbsp;
<a href="<%= "htadd.asp?key=" & Server.URLEncode(sKey) %>">���ƴ�����¼</a>&nbsp;
<a href="<%= "htdelete.asp?key=" & Server.URLEncode(sKey) %>">ɾ��������¼</a>&nbsp;
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
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��ͬ��</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_5408540C53F7 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�ͻ�����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_5BA26237540D79F0 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ�ͺ�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EA754C1578B53F7 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_657091CF %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ�۸�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EF7683C %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�ɽ����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_91D1989D %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�¶�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_67085EA6 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_4EA4671F %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�а�ҵ��Ա</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_627F529E %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��ע</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_59076CE8 %>
</span></td>
	</tr>
	<tr>
		<td width="80" bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">ҵ�����</span></div></td>
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
		x_5408540C53F7 = rs("��ͬ��")
		x_5BA26237540D79F0 = rs("�ͻ�����")
		x_4EA754C1578B53F7 = rs("��Ʒ�ͺ�")
		x_657091CF = rs("����")
		x_4EF7683C = rs("�۸�")
		x_91D1989D = rs("���")
		x_67085EA6 = rs("�¶�")
		x_4EA4671F = rs("����")
		x_627F529E = rs("�а�")
		x_59076CE8 = rs("��ע")
		x_53CD998862A5916C = rs("��������")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
