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

' Get action
sAction = Request.Form("a_add")
If (sAction = "" Or IsNull(sAction)) Then
	sKey = Request.Querystring("key")
	If sKey <> "" Then
		sAction = "C" ' Copy record
	Else
		sAction = "I" ' Display blank record
	End If
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

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "C": ' Get a record to display
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "htlist.asp"
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session("ewmsg") = "Add New Record Successful"
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
	if (!EW_onError(EW_this, EW_this.x_91D1989D, "TEXT", "Incorrect floating point number - ���"))
		return false; 
}
if (EW_this.x_67085EA6 && !EW_checkinteger(EW_this.x_67085EA6.value)) {
	if (!EW_onError(EW_this, EW_this.x_67085EA6, "TEXT", "Incorrect integer - �¶�"))
		return false; 
}
return true;
}
//-->
</script>
<form name="htadd" id="htadd" action="htadd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_add" value="A">
<table width="540" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="B0C8EA">
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��ͬ��</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5408540C53F7" id="x_5408540C53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5408540C53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�ͻ�����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5BA26237540D79F0" id="x_5BA26237540D79F0" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5BA26237540D79F0&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ�ͺ�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA754C1578B53F7" id="x_4EA754C1578B53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EA754C1578B53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_657091CF" id="x_657091CF" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_657091CF&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��Ʒ�۸�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EF7683C" id="x_4EF7683C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EF7683C&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�ɽ����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% If IsNull(x_91D1989D) or x_91D1989D = "" Then x_91D1989D = 0 ' Set default value %>
<input type="text" name="x_91D1989D" id="x_91D1989D" size="30" value="<%= Server.HTMLEncode(x_91D1989D&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�¶�</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_67085EA6" id="x_67085EA6" size="30" value="<%= month(date)%><%= Server.HTMLEncode(x_67085EA6&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA4671F" id="x_4EA4671F" size="30" maxlength="50" value="<%= date()%><%= Server.HTMLEncode(x_4EA4671F&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">�а�ҵ��Ա</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_627F529E" id="x_627F529E" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_627F529E&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">��ע</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_59076CE8" id="x_59076CE8" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_59076CE8&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">ҵ�����</span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_53CD998862A5916C" id="x_53CD998862A5916C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_53CD998862A5916C&"") %>">
</span></td>
	</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="ȷ�����">
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
<%

'-------------------------------------------------------------------------------
' Function AddData
' - Add Data
' - Variables used: field variables

Function AddData()
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Add New Record
	sSql = "SELECT * FROM [ht]"
	sSql = sSql & " WHERE 0 = 1"
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
	rs.AddNew

	' Field ��ͬ��
	sTmp = Trim(x_5408540C53F7)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("��ͬ��") = sTmp

	' Field �ͻ�����
	sTmp = Trim(x_5BA26237540D79F0)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("�ͻ�����") = sTmp

	' Field ��Ʒ�ͺ�
	sTmp = Trim(x_4EA754C1578B53F7)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("��Ʒ�ͺ�") = sTmp

	' Field ����
	sTmp = Trim(x_657091CF)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("����") = sTmp

	' Field �۸�
	sTmp = Trim(x_4EF7683C)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("�۸�") = sTmp

	' Field ���
	sTmp = x_91D1989D
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = cDbl(sTmp)
	End If
	rs("���") = sTmp

	' Field �¶�
	sTmp = x_67085EA6
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("�¶�") = sTmp

	' Field ����
	sTmp = Trim(x_4EA4671F)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("����") = sTmp

	' Field �а�
	sTmp = Trim(x_627F529E)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("�а�") = sTmp

	' Field ��ע
	sTmp = Trim(x_59076CE8)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("��ע") = sTmp

	' Field ��������
	sTmp = Trim(x_53CD998862A5916C)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("��������") = sTmp
	rs.Update
	rs.Close
	Set rs = Nothing
	AddData = True
End Function
%>
