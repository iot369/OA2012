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
sAction = Request.Form("a_search")
Select Case sAction
	Case "S": ' Get Search Criteria

	' Build search string for advanced search, remove blank field
	sSrchStr = ""

	' Field ID
	x_ID = Request.Form("x_ID")
	z_ID = Request.Form("z_ID")
	If x_ID <> "" Then
		sSrchFld = x_ID
		sSrchWrk = "x_ID=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_ID=" & Server.URLEncode(z_ID)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 合同号
	x_5408540C53F7 = Request.Form("x_5408540C53F7")
	z_5408540C53F7 = Request.Form("z_5408540C53F7")
	If x_5408540C53F7 <> "" Then
		sSrchFld = x_5408540C53F7
		sSrchWrk = "x_5408540C53F7=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_5408540C53F7=" & Server.URLEncode(z_5408540C53F7)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 客户名称
	x_5BA26237540D79F0 = Request.Form("x_5BA26237540D79F0")
	z_5BA26237540D79F0 = Request.Form("z_5BA26237540D79F0")
	If x_5BA26237540D79F0 <> "" Then
		sSrchFld = x_5BA26237540D79F0
		sSrchWrk = "x_5BA26237540D79F0=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_5BA26237540D79F0=" & Server.URLEncode(z_5BA26237540D79F0)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 产品型号
	x_4EA754C1578B53F7 = Request.Form("x_4EA754C1578B53F7")
	z_4EA754C1578B53F7 = Request.Form("z_4EA754C1578B53F7")
	If x_4EA754C1578B53F7 <> "" Then
		sSrchFld = x_4EA754C1578B53F7
		sSrchWrk = "x_4EA754C1578B53F7=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_4EA754C1578B53F7=" & Server.URLEncode(z_4EA754C1578B53F7)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 数量
	x_657091CF = Request.Form("x_657091CF")
	z_657091CF = Request.Form("z_657091CF")
	If x_657091CF <> "" Then
		sSrchFld = x_657091CF
		sSrchWrk = "x_657091CF=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_657091CF=" & Server.URLEncode(z_657091CF)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 价格
	x_4EF7683C = Request.Form("x_4EF7683C")
	z_4EF7683C = Request.Form("z_4EF7683C")
	If x_4EF7683C <> "" Then
		sSrchFld = x_4EF7683C
		sSrchWrk = "x_4EF7683C=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_4EF7683C=" & Server.URLEncode(z_4EF7683C)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 金额
	x_91D1989D = Request.Form("x_91D1989D")
	z_91D1989D = Request.Form("z_91D1989D")
	If x_91D1989D <> "" Then
		sSrchFld = x_91D1989D
		sSrchWrk = "x_91D1989D=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_91D1989D=" & Server.URLEncode(z_91D1989D)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 月度
	x_67085EA6 = Request.Form("x_67085EA6")
	z_67085EA6 = Request.Form("z_67085EA6")
	If x_67085EA6 <> "" Then
		sSrchFld = x_67085EA6
		sSrchWrk = "x_67085EA6=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_67085EA6=" & Server.URLEncode(z_67085EA6)
		y_67085EA6 = Request.Form("y_67085EA6")
		If y_67085EA6 <> "" Then
			sSrchFld = y_67085EA6
			sSrchWrk = sSrchWrk & "&y_67085EA6=" & Server.URLEncode(sSrchFld)
		End If
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 交期
	x_4EA4671F = Request.Form("x_4EA4671F")
	z_4EA4671F = Request.Form("z_4EA4671F")
	If x_4EA4671F <> "" Then
		sSrchFld = x_4EA4671F
		sSrchWrk = "x_4EA4671F=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_4EA4671F=" & Server.URLEncode(z_4EA4671F)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 承办
	x_627F529E = Request.Form("x_627F529E")
	z_627F529E = Request.Form("z_627F529E")
	If x_627F529E <> "" Then
		sSrchFld = x_627F529E
		sSrchWrk = "x_627F529E=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_627F529E=" & Server.URLEncode(z_627F529E)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 备注
	x_59076CE8 = Request.Form("x_59076CE8")
	z_59076CE8 = Request.Form("z_59076CE8")
	If x_59076CE8 <> "" Then
		sSrchFld = x_59076CE8
		sSrchWrk = "x_59076CE8=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_59076CE8=" & Server.URLEncode(z_59076CE8)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If

	' Field 反馈报酬
	x_53CD998862A5916C = Request.Form("x_53CD998862A5916C")
	z_53CD998862A5916C = Request.Form("z_53CD998862A5916C")
	If x_53CD998862A5916C <> "" Then
		sSrchFld = x_53CD998862A5916C
		sSrchWrk = "x_53CD998862A5916C=" & Server.URLEncode(sSrchFld)
		sSrchWrk = sSrchWrk & "&z_53CD998862A5916C=" & Server.URLEncode(z_53CD998862A5916C)
	Else
		sSrchWrk = ""
	End If
	If sSrchWrk <> "" Then
		If sSrchStr = "" Then
			sSrchStr = sSrchWrk
		Else
			sSrchStr = sSrchStr & "&" & sSrchWrk
		End If
	End If
	If sSrchStr <> "" Then
		Response.Clear
		Response.Redirect "htlist.asp" & "?" & sSrchStr
	End If
End Select

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
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
if (EW_this.x_ID && !EW_checkinteger(EW_this.x_ID.value)) {
	if (!EW_onError(EW_this, EW_this.x_ID, "NO", "Incorrect integer - ID"))
		return false; 
}
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
<form name="htsearch" id="htsearch" action="htsrch.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_search" value="S">
<table width="80%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="B0C8EA">
	

	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">合<span style="">同号</span></span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_5408540C53F7" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5408540C53F7" id="x_5408540C53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5408540C53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">客户名称</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_5BA26237540D79F0" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_5BA26237540D79F0" id="x_5BA26237540D79F0" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_5BA26237540D79F0&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">产品型号</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_4EA754C1578B53F7" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA754C1578B53F7" id="x_4EA754C1578B53F7" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EA754C1578B53F7&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">数量</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_657091CF" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_657091CF" id="x_657091CF" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_657091CF&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">价格</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_4EF7683C" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EF7683C" id="x_4EF7683C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EF7683C&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">金额</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">=
	          <input type="hidden" name="z_91D1989D" value="=,,">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_91D1989D" id="x_91D1989D" size="30" value="<%= Server.HTMLEncode(x_91D1989D&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker">月度</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">=
	          <input type="hidden" name="z_67085EA6" value="=,,">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_67085EA6" id="x_67085EA6" size="30" value="<%= Server.HTMLEncode(x_67085EA6&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">交期</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_4EA4671F" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_4EA4671F" id="x_4EA4671F" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_4EA4671F&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">承办</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_627F529E" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_627F529E" id="x_627F529E" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_627F529E&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">备注</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_59076CE8" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_59076CE8" id="x_59076CE8" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_59076CE8&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="D7E8F8"><div align="center" style="color: #000000"><span class="aspmaker" style="">业务提成</span></div></td>
		<td bgcolor="#FFFFFF"><div align="center"><span class="aspmaker">近似值
	          <input type="hidden" name="z_53CD998862A5916C" value="LIKE,'%,%'">
		  </span></div></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_53CD998862A5916C" id="x_53CD998862A5916C" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_53CD998862A5916C&"") %>">
</span></td>
	</tr>
</table>
<p align="center">
<input type="submit" name="Action" value="进行搜索">
</form>
<%
conn.Close ' Close Connection
Set conn = Nothing
%>
