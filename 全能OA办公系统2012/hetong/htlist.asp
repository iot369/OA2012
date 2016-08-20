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
<%
sExport = Request.QueryString("export") ' Load Export Request
If sExport = "html" Then

	' Printer Friendly
End If
If sExport = "excel" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=ht.xls"
End If
If sExport = "word" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition:", "attachment; filename=ht.doc"
End If
If sExport = "xml" Then
	Response.ContentType = "text/xml"
	Response.AddHeader "Content-Disposition:", "attachment; filename=ht.xml"
End If
If sExport = "csv" Then
	Response.ContentType = "application/csv"
	Response.AddHeader "Content-Disposition:", "attachment; filename=ht.csv"
End If
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<% 
nStartRec = 0
nStopRec = 0
nTotalRecs = 0
nRecCount = 0
nRecActual = 0
sKeyMaster = ""
sDbWhereMaster = ""
sSrchAdvanced = ""
sSrchBasic = ""
sSrchWhere = ""
sDbWhere = ""
sDefaultOrderBy = ""
sDefaultFilter = ""
sWhere = ""
sGroupBy = ""
sHaving = ""
sOrderBy = ""
sSqlMaster = ""
nDisplayRecs = 20
nRecRange = 10

' Set up records per page dynamically
SetUpDisplayRecs()

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Handle Reset Command
ResetCmd()

' Get Search Criteria for Advanced Search
SetUpAdvancedSearch()

' Get Search Criteria for Basic Search
SetUpBasicSearch()

' Build Search Criteria
If sSrchAdvanced <> "" Then
	sSrchWhere = sSrchAdvanced ' Advanced Search
ElseIf sSrchBasic <> "" Then
	sSrchWhere = sSrchBasic ' Basic Search
End If

' Save Search Criteria
If sSrchWhere <> "" Then
	Session("ht_searchwhere") = sSrchWhere

	' Reset start record counter (new search)
	nStartRec = 1
	Session("ht_REC") = nStartRec
Else
	sSrchWhere = Session("ht_searchwhere")
End If

' Build WHERE condition
sDbWhere = ""
If sDbWhereMaster <> "" Then
	sDbWhere = sDbWhere & "(" & sDbWhereMaster & ") AND "
End If
If sSrchWhere <> "" Then
	sDbWhere = sDbWhere & "(" & sSrchWhere & ") AND "
End If
If Len(sDbWhere) > 5 Then
	sDbWhere = Mid(sDbWhere, 1, Len(sDbWhere)-5) ' Trim rightmost AND
End If

' Build SQL
sSql = "SELECT * FROM [ht]"

' Load Default Filter
sDefaultFilter = ""
sGroupBy = ""
sHaving = ""

' Load Default Order
sDefaultOrderBy = "[合同号] DESC"
sWhere = ""
If sDefaultFilter <> "" Then
	sWhere = sWhere & "(" & sDefaultFilter & ") AND "
End If
If sDbWhere <> "" Then
	sWhere = sWhere & "(" & sDbWhere & ") AND "
End If
If Right(sWhere, 5) = " AND " Then sWhere = Left(sWhere, Len(sWhere)-5)
If sWhere <> "" Then
	sSql = sSql & " WHERE " & sWhere
End If
If sGroupBy <> "" Then
	sSql = sSql & " GROUP BY " & sGroupBy
End If	
If sHaving <> "" Then
	sSql = sSql & " HAVING " & sHaving
End If	

' Set Up Sorting Order
sOrderBy = ""
SetUpSortOrder()
If sOrderBy <> "" Then
	sSql = sSql & " ORDER BY " & sOrderBy
End If	

'Session("ewmsg") = sSql ' Uncomment to show SQL for debugging
' Export Data only

If sExport = "xml" Or sExport = "csv" Then
	Call ExportData(sExport, sSql)
	conn.Close ' Close Connection
	Set conn = Nothing
	Response.End
End If
%>
<% If sExport <> "word" And sExport <> "excel" Then %>
<!--#include file="header.asp"-->
<script type="text/javascript" src="ew.js"></script>
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<% End If %>
<%

' Set up Record Set
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open sSql, conn, 1, 2
nTotalRecs = rs.RecordCount
If nDisplayRecs <= 0 Then ' Display All Records
	nDisplayRecs = nTotalRecs
End If
nStartRec = 1
SetUpStartRec() ' Set Up Start Record Position
%>
<p><span class="aspmaker"><% If sExport = "" Then %>
&nbsp;&nbsp;
&nbsp;&nbsp;<a href="htlist.asp?export=excel"><img src="images/exportxls.gif" width="16" height="16" border="0">导入Excel</a>
&nbsp;&nbsp;<a href="htlist.asp?export=word"><img src="images/exportdoc.gif" width="16" height="16" border="0">导入Word</a>
&nbsp;&nbsp;<a href="htlist.asp?export=xml"><img src="images/exportxml.gif" width="16" height="16" border="0">导入Xml</a>
&nbsp;&nbsp;<a href="htlist.asp?export=csv"><img src="images/exportcsv.gif" width="16" height="16" border="0">导入Excel逗号分隔值</a>
<% End If %>
</span></p>
<% If sExport = "" Then %>
<form action="htlist.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td><div align="center"><span class="aspmaker">			  </span><span class="aspmaker">
		  <input type="radio" name="psearchtype" value="" checked>
		  模糊查询&nbsp;&nbsp;
		  <input type="radio" name="psearchtype" value="AND">
		  所有关键字&nbsp;&nbsp;
		  <input type="radio" name="psearchtype" value="OR">
		  单一关键字
		  <input type="text" name="psearch" size="20">
          <input type="Submit" name="Submit" value="查询">
&nbsp;&nbsp; <a href="htlist.asp?cmd=reset"><br>
<br>
显示所有</a>&nbsp;&nbsp; <a href="htsrch.asp">高级查询</a></span></div></td>
	</tr>
</table>
</form>
<% End If %>
<% If sExport = "" Then %>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><span class="aspmaker"><a href="htadd.asp">合同添加</a></span></td>
	</tr>
</table>
<p>
<% End If %>
<%
If Session("ewmsg") <> "" Then
%>
<p><span class="aspmaker" style="color: Red;"><%= Session("ewmsg") %></span></p>
<%
	Session("ewmsg") = "" ' Clear message
End If
%>
<form method="post">
<table border="0"  cellspacing="1" cellpadding="0" width="98%" bgcolor="B0C8EA"  align="center">
<% If nTotalRecs > 0 Then %>
	<!-- Table header -->
	<tr align="center" valign="middle" bgcolor="D7E8F8">
		<td width="65" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
合同号
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("合同号") %>" style="color: #2b486a;">合同号<% If Session("ht_x_5408540C53F7_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_5408540C53F7_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="65" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
客户名称
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("客户名称") %>" style="color: #2b486a;">客户名称<% If Session("ht_x_5BA26237540D79F0_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_5BA26237540D79F0_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="65" height="30" bgcolor="D7E8F8"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
产品型号
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("产品型号") %>" style="color: #2b486a;">产品型号<% If Session("ht_x_4EA754C1578B53F7_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_4EA754C1578B53F7_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="35" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
数量
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("数量") %>" style="color: #2b486a;">数量<% If Session("ht_x_657091CF_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_657091CF_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="65" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
产品价格
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("价格") %>" style="color: #2b486a;">产品价格<% If Session("ht_x_4EF7683C_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_4EF7683C_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="65" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
成交金额
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("金额") %>" style="color: #2b486a;">成交金额<% If Session("ht_x_91D1989D_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_91D1989D_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td width="40" height="30"><div align="center"><span style="color: #2b486a; font-size: 12px;">
  <% If sExport <> "" Then %>
  交期
  <% Else %>
	  <a href="htlist.asp?order=<%= Server.URLEncode("交期") %>" style="color: #2b486a;">交期
	  <% If Session("ht_x_4EA4671F_Sort") = "ASC" Then %>
	  <img src="images/sortup.gif" width="10" height="9" border="0">
	  <% ElseIf Session("ht_x_4EA4671F_Sort") = "DESC" Then %>
	  <img src="images/sortdown.gif" width="10" height="9" border="0">
	  <% End If %>
	  </a>
      <% End If %>
		  </span></div></td>
		<td width="50" height="30"><span style="color: #2b486a; font-size: 12px;">
<% If sExport <> "" Then %>
承办人
<% Else %>
	<a href="htlist.asp?order=<%= Server.URLEncode("承办") %>" style="color: #2b486a;">承办人<% If Session("ht_x_627F529E_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("ht_x_627F529E_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<% If sExport = "" Then %>
<td height="30">&nbsp;</td>
<td height="30">&nbsp;</td>
<td height="30">&nbsp;</td>
<% End If %>
	</tr>
<% End If %>
<%

' Avoid starting record > total records
If CLng(nStartRec) > CLng(nTotalRecs) Then
	nStartRec = nTotalRecs
End If

' Set the last record to display
nStopRec = nStartRec + nDisplayRecs - 1

' Move to first record directly for performance reason
nRecCount = nStartRec - 1
If Not rs.Eof Then
	rs.MoveFirst
	rs.Move nStartRec - 1
End If
Dim tot_x_91D1989D
tot_x_91D1989D = 0 ' Initialise total to zero for aggregation
nRecActual = 0
Do While (Not rs.Eof) And (nRecCount < nStopRec)
	nRecCount = nRecCount + 1
	If CLng(nRecCount) >= CLng(nStartRec) Then 
		nRecActual = nRecActual + 1

	' Set row color
	sItemRowClass = " bgcolor=""#FFFFFF"""

	' Display alternate color for rows
	If nRecCount Mod 2 <> 0 Then
		sItemRowClass = " bgcolor=""#EBF3FC"""
	End If

		' Load Key for record
		sKey = rs("ID")
		x_ID = rs("ID")
		x_5408540C53F7 = rs("合同号")
		x_5BA26237540D79F0 = rs("客户名称")
		x_4EA754C1578B53F7 = rs("产品型号")
		x_657091CF = rs("数量")
		x_4EF7683C = rs("价格")
		x_91D1989D = rs("金额")
		tot_x_91D1989D = tot_x_91D1989D + x_91D1989D ' Accumulate Total
		x_67085EA6 = rs("月度")
		x_4EA4671F = rs("交期")
		x_627F529E = rs("承办")
		x_59076CE8 = rs("备注")
		x_53CD998862A5916C = rs("反馈报酬")
%>
	<!-- Table body -->
	<tr align="center"<%=sItemRowClass%>>
		<!-- ID -->
		<!-- 合同号 -->
		<td><span class="aspmaker">
<a href="<% If Not IsNull(sKey) Then Response.Write "htview.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>"><% Response.Write x_5408540C53F7 %></a>
</span></td>
		<!-- 客户名称 -->
		<td><span class="aspmaker">
<% Response.Write x_5BA26237540D79F0 %>
</span></td>
		<!-- 产品型号 -->
		<td><span class="aspmaker">
<% Response.Write x_4EA754C1578B53F7 %>
</span></td>
		<!-- 数量 -->
		<td><span class="aspmaker">
<% Response.Write x_657091CF %>
</span></td>
		<!-- 价格 -->
		<td><span class="aspmaker">
<% Response.Write x_4EF7683C %>
</span></td>
		<!-- 金额 -->
		<td><span class="aspmaker">
<% Response.Write x_91D1989D %>
</span></td>
		<!-- 月度 -->
		<!-- 交期 -->
		<td><div align="center"><span class="aspmaker">
  <% Response.Write x_4EA4671F %>
        </span></div></td>
		<!-- 承办 -->
		<td><span class="aspmaker">
<% Response.Write x_627F529E %>
</span></td>
		<!-- 备注 -->
		<!-- 反馈报酬 -->
		<% If sExport = "" Then %>
<td><span class="aspmaker"><a href="<% If Not IsNull(sKey) Then Response.Write "htedit.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>"><img src="images/edit.gif" alt="修改" width="16" height="16" border="0"></a></span></td>
<td><span class="aspmaker"><a href="<% If Not IsNull(sKey) Then Response.Write "htadd.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>"><img src="images/copy.gif" alt="复制" width="16" height="16" border="0"></a></span></td>
<td><span class="aspmaker"><a href="<% If Not IsNull(sKey) Then Response.Write "htdelete.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>"><img src="images/delete.gif" alt="删除" width="16" height="16" border="0"></a></span></td>
<% End If %>
	</tr>
<%
	End If
	rs.MoveNext
Loop
%>
<%
x_91D1989D = tot_x_91D1989D
%>
<% If nTotalRecs > 0 Then %>
<!-- Table footer -->
<% End If %>
</table>
</form>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<% If sExport = "" Then %>
<form action="htlist.asp" name="ewpagerform" id="ewpagerform">
<table bgcolor="" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
	<tr>
		<td nowrap>
<%
If nTotalRecs > 0 Then
	rsEof = (nTotalRecs < (nStartRec + nDisplayRecs))
	PrevStart = nStartRec - nDisplayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = nStartRec + nDisplayRecs
	If NextStart > nTotalRecs Then NextStart = nStartRec
	LastStart = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
	%>
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker">当前为&nbsp;</span></td>
<!--first page button-->
	<% If CLng(nStartRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="16" height="16" border="0"></td>
	<% Else %>
	<td><a href="htlist.asp?start=1"><img src="images/first.gif" alt="First" width="16" height="16" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(nStartRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="16" height="16" border="0"></td>
	<% Else %>
	<td><a href="htlist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="16" height="16" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(nStartRec-1)\nDisplayRecs+1%>" size="4"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(nStartRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="16" height="16" border="0"></td>
	<% Else %>
	<td><a href="htlist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="16" height="16" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(nStartRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="16" height="16" border="0"></td>
	<% Else %>
	<td><a href="htlist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="16" height="16" border="0"></a></td>
	<% End If %>
	<td><span class="aspmaker">&nbsp;共 <%=(nTotalRecs-1)\nDisplayRecs+1%>页</span></td>
	</tr></table>
	<% If CLng(nStartRec) > CLng(nTotalRecs) Then nStartRec = nTotalRecs
	nStopRec = nStartRec + nDisplayRecs - 1
	nRecCount = nTotalRecs - 1
	If rsEOF Then nRecCount = nTotalRecs
	If nStopRec > nRecCount Then nStopRec = nRecCount %>
	<span class="aspmaker">第 <%= nStartRec %> 至 <%= nStopRec %> 条合同记录,共<%= nTotalRecs %>条合同记录</span>
    <% Else %>
	<span class="aspmaker">对不起,还没有合同记录</span>
    <% End If %>
		</td>
<% If nTotalRecs > 0 Then %>
		<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td align="right" valign="top" nowrap><span class="aspmaker">每页显示&nbsp;
<select name="RecPerPage" onChange="this.form.submit();" class="aspmaker">
<option value="20"<% If nDisplayRecs = 20 Then response.write " selected" %>>20</option>
<option value="60"<% If nDisplayRecs = 60 Then response.write " selected" %>>60</option>
<option value="ALL"<% If Session("ht_RecPerPage") = -1 Then response.write " selected" %>>所有记录</option>
</select>
		条记录</span></td>
<% End If %>
	</tr>
</table>
</form>	
<% End If %>
<% If sExport <> "word" And sExport <> "excel" Then %>
<% End If %>
<%

'-------------------------------------------------------------------------------
' Function SetUpDisplayRecs
' - Set up Number of Records displayed per page based on Form element RecPerPage
' - Variables setup: nDisplayRecs

Sub SetUpDisplayRecs()
	Dim sWrk
	sWrk = Request.QueryString("RecPerPage")
	If sWrk <> "" Then
		If IsNumeric(sWrk) Then
			nDisplayRecs = CInt(sWrk)
		Else
			If UCase(sWrk) = "ALL" Then ' Display All Records
				nDisplayRecs = -1
			Else
				nDisplayRecs = 20 ' Non-numeric, Load Default
			End If
		End If
		Session("ht_RecPerPage") = nDisplayRecs ' Save to Session

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("ht_REC") = nStartRec
	Else
		If Session("ht_RecPerPage") <> "" Then
			nDisplayRecs = Session("ht_RecPerPage") ' Restore from Session
		Else
			nDisplayRecs = 20 ' Load Default
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpAdvancedSearch
' - Set up Advanced Search parameter based on querystring parameters from Advanced Search Page
' - Variables setup: sSrchAdvanced

Sub SetUpAdvancedSearch()
	Dim arrFldOpr

	' Field ID
	x_ID = Request.QueryString("x_ID")
	z_ID = Request.QueryString("z_ID")
	arrFldOpr = Split(z_ID,",")
	If x_ID <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[ID] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_ID) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 合同号
	x_5408540C53F7 = Request.QueryString("x_5408540C53F7")
	z_5408540C53F7 = Request.QueryString("z_5408540C53F7")
	arrFldOpr = Split(z_5408540C53F7,",")
	If x_5408540C53F7 <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[合同号] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_5408540C53F7) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 客户名称
	x_5BA26237540D79F0 = Request.QueryString("x_5BA26237540D79F0")
	z_5BA26237540D79F0 = Request.QueryString("z_5BA26237540D79F0")
	arrFldOpr = Split(z_5BA26237540D79F0,",")
	If x_5BA26237540D79F0 <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[客户名称] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_5BA26237540D79F0) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 产品型号
	x_4EA754C1578B53F7 = Request.QueryString("x_4EA754C1578B53F7")
	z_4EA754C1578B53F7 = Request.QueryString("z_4EA754C1578B53F7")
	arrFldOpr = Split(z_4EA754C1578B53F7,",")
	If x_4EA754C1578B53F7 <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[产品型号] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_4EA754C1578B53F7) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 数量
	x_657091CF = Request.QueryString("x_657091CF")
	z_657091CF = Request.QueryString("z_657091CF")
	arrFldOpr = Split(z_657091CF,",")
	If x_657091CF <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[数量] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_657091CF) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 价格
	x_4EF7683C = Request.QueryString("x_4EF7683C")
	z_4EF7683C = Request.QueryString("z_4EF7683C")
	arrFldOpr = Split(z_4EF7683C,",")
	If x_4EF7683C <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[价格] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_4EF7683C) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 金额
	x_91D1989D = Request.QueryString("x_91D1989D")
	z_91D1989D = Request.QueryString("z_91D1989D")
	arrFldOpr = Split(z_91D1989D,",")
	If x_91D1989D <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[金额] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_91D1989D) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 月度
	x_67085EA6 = Request.QueryString("x_67085EA6")
	z_67085EA6 = Request.QueryString("z_67085EA6")
	arrFldOpr = Split(z_67085EA6,",")
	If x_67085EA6 <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[月度] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_67085EA6) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If
	y_67085EA6 = Request.QueryString("y_67085EA6")
	If y_67085EA6 <> "" And UBound(arrFldOpr) >=5 Then
		sSrchAdvanced = sSrchAdvanced & "[月度] " & arrFldOpr(3) & " " & arrFldOpr(4) & y_67085EA6 & arrFldOpr(5) & " AND "
	End If

	' Field 交期
	x_4EA4671F = Request.QueryString("x_4EA4671F")
	z_4EA4671F = Request.QueryString("z_4EA4671F")
	arrFldOpr = Split(z_4EA4671F,",")
	If x_4EA4671F <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[交期] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_4EA4671F) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 承办
	x_627F529E = Request.QueryString("x_627F529E")
	z_627F529E = Request.QueryString("z_627F529E")
	arrFldOpr = Split(z_627F529E,",")
	If x_627F529E <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[承办] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_627F529E) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 备注
	x_59076CE8 = Request.QueryString("x_59076CE8")
	z_59076CE8 = Request.QueryString("z_59076CE8")
	arrFldOpr = Split(z_59076CE8,",")
	If x_59076CE8 <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[备注] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_59076CE8) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If

	' Field 反馈报酬
	x_53CD998862A5916C = Request.QueryString("x_53CD998862A5916C")
	z_53CD998862A5916C = Request.QueryString("z_53CD998862A5916C")
	arrFldOpr = Split(z_53CD998862A5916C,",")
	If x_53CD998862A5916C <> "" Then
		sSrchAdvanced = sSrchAdvanced & "[反馈报酬] " ' Add field
		sSrchAdvanced = sSrchAdvanced	& arrFldOpr(0) & " " ' Add operator
		If UBound(arrFldOpr) >= 1 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(1) ' Add search prefix
		End If
		sSrchAdvanced = sSrchAdvanced & AdjustSql(x_53CD998862A5916C) ' Add input parameter
		If UBound(arrFldOpr) >=2 Then
			sSrchAdvanced = sSrchAdvanced & arrFldOpr(2) ' Add search suffix
		End If
		sSrchAdvanced = sSrchAdvanced	& " AND "
	End If
	If Len(sSrchAdvanced) > 4 Then
		sSrchAdvanced = Mid(sSrchAdvanced, 1, Len(sSrchAdvanced)-4)
	End If
End Sub

'-------------------------------------------------------------------------------
' Function BasicSearchSQL
' - Build WHERE clause for a keyword

Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = AdjustSql(Keyword)
	BasicSearchSQL = ""
	If IsNumeric(sKeyword) Then BasicSearchSQL = BasicSearchSQL & "[ID] = " & sKeyword & " OR "
	BasicSearchSQL = BasicSearchSQL & "[合同号] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[客户名称] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[产品型号] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[数量] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[价格] LIKE '%" & sKeyword & "%' OR "
	If IsNumeric(sKeyword) Then BasicSearchSQL = BasicSearchSQL & "[金额] = " & sKeyword & " OR "
	If IsNumeric(sKeyword) Then BasicSearchSQL = BasicSearchSQL & "[月度] = " & sKeyword & " OR "
	BasicSearchSQL = BasicSearchSQL & "[交期] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[承办] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[备注] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[反馈报酬] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

'-------------------------------------------------------------------------------
' Function SetUpBasicSearch
' - Set up Basic Search parameter based on form elements pSearch & pSearchType
' - Variables setup: sSrchBasic

Sub SetUpBasicSearch()
	Dim sSearch, sSearchType, arKeyword, sKeyword
	sSearch = Request.QueryString("psearch")
	sSearchType = Request.QueryString("psearchType")
	If sSearch <> "" Then
		If sSearchType <> "" Then
			While InStr(sSearch, "  ") > 0
				sSearch = Replace(sSearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(sSearch), " ")
			For Each sKeyword In arKeyword
				sSrchBasic = sSrchBasic & "(" & BasicSearchSQL(sKeyword) & ") " & sSearchType & " "
			Next
		Else
			sSrchBasic = BasicSearchSQL(sSearch)
		End If
	End If
	If Right(sSrchBasic, 4) = " OR " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-4)
	If Right(sSrchBasic, 5) = " AND " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-5)
End Sub

'-------------------------------------------------------------------------------
' Function SetUpSortOrder
' - Set up Sort parameters based on Sort Links clicked
' - Variables setup: sOrderBy, Session("Table_OrderBy"), Session("Table_Field_Sort")

Sub SetUpSortOrder()
	Dim sOrder, sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		sOrder = Request.QueryString("order")

		' Field ID
		If sOrder = "ID" Then
			sSortField = "[ID]"
			sLastSort = Session("ht_x_ID_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_ID_Sort") = sThisSort
		Else
			If Session("ht_x_ID_Sort") <> "" Then Session("ht_x_ID_Sort") = ""
		End If

		' Field 合同号
		If sOrder = "合同号" Then
			sSortField = "[合同号]"
			sLastSort = Session("ht_x_5408540C53F7_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_5408540C53F7_Sort") = sThisSort
		Else
			If Session("ht_x_5408540C53F7_Sort") <> "" Then Session("ht_x_5408540C53F7_Sort") = ""
		End If

		' Field 客户名称
		If sOrder = "客户名称" Then
			sSortField = "[客户名称]"
			sLastSort = Session("ht_x_5BA26237540D79F0_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_5BA26237540D79F0_Sort") = sThisSort
		Else
			If Session("ht_x_5BA26237540D79F0_Sort") <> "" Then Session("ht_x_5BA26237540D79F0_Sort") = ""
		End If

		' Field 产品型号
		If sOrder = "产品型号" Then
			sSortField = "[产品型号]"
			sLastSort = Session("ht_x_4EA754C1578B53F7_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_4EA754C1578B53F7_Sort") = sThisSort
		Else
			If Session("ht_x_4EA754C1578B53F7_Sort") <> "" Then Session("ht_x_4EA754C1578B53F7_Sort") = ""
		End If

		' Field 数量
		If sOrder = "数量" Then
			sSortField = "[数量]"
			sLastSort = Session("ht_x_657091CF_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_657091CF_Sort") = sThisSort
		Else
			If Session("ht_x_657091CF_Sort") <> "" Then Session("ht_x_657091CF_Sort") = ""
		End If

		' Field 价格
		If sOrder = "价格" Then
			sSortField = "[价格]"
			sLastSort = Session("ht_x_4EF7683C_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_4EF7683C_Sort") = sThisSort
		Else
			If Session("ht_x_4EF7683C_Sort") <> "" Then Session("ht_x_4EF7683C_Sort") = ""
		End If

		' Field 金额
		If sOrder = "金额" Then
			sSortField = "[金额]"
			sLastSort = Session("ht_x_91D1989D_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_91D1989D_Sort") = sThisSort
		Else
			If Session("ht_x_91D1989D_Sort") <> "" Then Session("ht_x_91D1989D_Sort") = ""
		End If

		' Field 月度
		If sOrder = "月度" Then
			sSortField = "[月度]"
			sLastSort = Session("ht_x_67085EA6_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_67085EA6_Sort") = sThisSort
		Else
			If Session("ht_x_67085EA6_Sort") <> "" Then Session("ht_x_67085EA6_Sort") = ""
		End If

		' Field 交期
		If sOrder = "交期" Then
			sSortField = "[交期]"
			sLastSort = Session("ht_x_4EA4671F_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_4EA4671F_Sort") = sThisSort
		Else
			If Session("ht_x_4EA4671F_Sort") <> "" Then Session("ht_x_4EA4671F_Sort") = ""
		End If

		' Field 承办
		If sOrder = "承办" Then
			sSortField = "[承办]"
			sLastSort = Session("ht_x_627F529E_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_627F529E_Sort") = sThisSort
		Else
			If Session("ht_x_627F529E_Sort") <> "" Then Session("ht_x_627F529E_Sort") = ""
		End If

		' Field 备注
		If sOrder = "备注" Then
			sSortField = "[备注]"
			sLastSort = Session("ht_x_59076CE8_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_59076CE8_Sort") = sThisSort
		Else
			If Session("ht_x_59076CE8_Sort") <> "" Then Session("ht_x_59076CE8_Sort") = ""
		End If

		' Field 反馈报酬
		If sOrder = "反馈报酬" Then
			sSortField = "[反馈报酬]"
			sLastSort = Session("ht_x_53CD998862A5916C_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("ht_x_53CD998862A5916C_Sort") = sThisSort
		Else
			If Session("ht_x_53CD998862A5916C_Sort") <> "" Then Session("ht_x_53CD998862A5916C_Sort") = ""
		End If
		Session("ht_OrderBy") = sSortField & " " & sThisSort
		Session("ht_REC") = 1
	End If
	sOrderBy = Session("ht_OrderBy")
	If sOrderBy = "" Then
		sOrderBy = sDefaultOrderBy
		Session("ht_OrderBy") = sOrderBy
		Session("ht_x_5408540C53F7_Sort") = "DESC"
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpStartRec
' - Set up Starting Record parameters based on Pager Navigation
' - Variables setup: nStartRec

Sub SetUpStartRec()
	Dim nPageNo

	' Check for a START parameter
	If Request.QueryString("start").Count > 0 Then
		nStartRec = Request.QueryString("start")
		Session("ht_REC") = nStartRec
	ElseIf Request.QueryString("pageno").Count > 0 Then
		nPageNo = Request.QueryString("pageno")
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Session("ht_REC") = nStartRec
		Else
			nStartRec = Session("ht_REC")
			If Not IsNumeric(nStartRec) Or nStartRec = "" Then			
				nStartRec = 1 ' Reset start record counter
				Session("ht_REC") = nStartRec
			End If
		End If
	Else
		nStartRec = Session("ht_REC")
		If Not IsNumeric(nStartRec) Or nStartRec = "" Then		
			nStartRec = 1 'Reset start record counter
			Session("ht_REC") = nStartRec
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ResetCmd
' - Clear list page parameters
' - RESET: reset search parameters
' - RESETALL: reset search & master/detail parameters
' - RESETSORT: reset sort parameters

Sub ResetCmd()
	Dim sCmd

	' Get Reset Cmd
	If Request.QueryString("cmd").Count > 0 Then
		sCmd = Request.QueryString("cmd")

		' Reset Search Criteria
		If UCase(sCmd) = "RESET" Then
			sSrchWhere = ""
			Session("ht_searchwhere") = sSrchWhere

		' Reset Search Criteria & Session Keys
		ElseIf UCase(sCmd) = "RESETALL" Then
			sSrchWhere = ""
			Session("ht_searchwhere") = sSrchWhere

		' Reset Sort Criteria
		ElseIf UCase(sCmd) = "RESETSORT" Then
			sOrderBy = ""
			Session("ht_OrderBy") = sOrderBy
			If Session("ht_x_ID_Sort") <> "" Then Session("ht_x_ID_Sort") = ""
			If Session("ht_x_5408540C53F7_Sort") <> "" Then Session("ht_x_5408540C53F7_Sort") = ""
			If Session("ht_x_5BA26237540D79F0_Sort") <> "" Then Session("ht_x_5BA26237540D79F0_Sort") = ""
			If Session("ht_x_4EA754C1578B53F7_Sort") <> "" Then Session("ht_x_4EA754C1578B53F7_Sort") = ""
			If Session("ht_x_657091CF_Sort") <> "" Then Session("ht_x_657091CF_Sort") = ""
			If Session("ht_x_4EF7683C_Sort") <> "" Then Session("ht_x_4EF7683C_Sort") = ""
			If Session("ht_x_91D1989D_Sort") <> "" Then Session("ht_x_91D1989D_Sort") = ""
			If Session("ht_x_67085EA6_Sort") <> "" Then Session("ht_x_67085EA6_Sort") = ""
			If Session("ht_x_4EA4671F_Sort") <> "" Then Session("ht_x_4EA4671F_Sort") = ""
			If Session("ht_x_627F529E_Sort") <> "" Then Session("ht_x_627F529E_Sort") = ""
			If Session("ht_x_59076CE8_Sort") <> "" Then Session("ht_x_59076CE8_Sort") = ""
			If Session("ht_x_53CD998862A5916C_Sort") <> "" Then Session("ht_x_53CD998862A5916C_Sort") = ""
		End If

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("ht_REC") = nStartRec
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ExportData
' - Export Data in Xml or Csv format

Sub ExportData(sExport, sSql)
	Dim oXmlDoc, oXmlTbl, oXmlRec, oXmlFld
	Dim sCsvStr
	Dim rs

	' Set up Record Set
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	nTotalRecs = rs.RecordCount
	nStartRec = 1
	SetUpStartRec() ' Set Up Start Record Position
	If sExport = "xml" Then
		Set oXmlDoc = Server.CreateObject("MSXML.DOMDocument")
		Set oXmlTbl = oXmlDoc.createElement("table")
	End If
	If sExport = "csv" Then
		sCsvStr = sCsvStr & """ID""" & ","
		sCsvStr = sCsvStr & """合同号""" & ","
		sCsvStr = sCsvStr & """客户名称""" & ","
		sCsvStr = sCsvStr & """产品型号""" & ","
		sCsvStr = sCsvStr & """数量""" & ","
		sCsvStr = sCsvStr & """价格""" & ","
		sCsvStr = sCsvStr & """金额""" & ","
		sCsvStr = sCsvStr & """月度""" & ","
		sCsvStr = sCsvStr & """交期""" & ","
		sCsvStr = sCsvStr & """承办""" & ","
		sCsvStr = sCsvStr & """备注""" & ","
		sCsvStr = sCsvStr & """反馈报酬""" & ","
		sCsvStr = Left(sCsvStr, Len(sCsvStr)-1) ' Remove last comma
		sCsvStr = sCsvStr & vbCrLf
	End If

	' Avoid starting record > total records
	If CLng(nStartRec) > CLng(nTotalRecs) Then
		nStartRec = nTotalRecs
	End If

	' Set the last record to display
	If nDisplayRecs < 0 Then
		nStopRec = nTotalRecs
	Else
		nStopRec = nStartRec + nDisplayRecs - 1
	End If

	' Move to first record directly for performance reason
	nRecCount = nStartRec - 1
	If Not rs.Eof Then
		rs.MoveFirst
		rs.Move nStartRec - 1
	End If
	nRecActual = 0
	Do While (Not rs.Eof) And (nRecCount < nStopRec)
		nRecCount = nRecCount + 1
		If CLng(nRecCount) >= CLng(nStartRec) Then 
			nRecActual = nRecActual + 1
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
			If sExport = "xml" Then
				Set oXmlRec = oXmlDoc.createElement("record")
				Call oXmlTbl.appendChild(oXmlRec)

				' Field ID
				Set oXmlFld = oXmlDoc.createElement("ID")
				sTmp = x_ID
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 合同号
				Set oXmlFld = oXmlDoc.createElement("5408540C53F7")
				sTmp = x_5408540C53F7
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 客户名称
				Set oXmlFld = oXmlDoc.createElement("5BA26237540D79F0")
				sTmp = x_5BA26237540D79F0
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 产品型号
				Set oXmlFld = oXmlDoc.createElement("4EA754C1578B53F7")
				sTmp = x_4EA754C1578B53F7
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 数量
				Set oXmlFld = oXmlDoc.createElement("657091CF")
				sTmp = x_657091CF
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 价格
				Set oXmlFld = oXmlDoc.createElement("4EF7683C")
				sTmp = x_4EF7683C
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 金额
				Set oXmlFld = oXmlDoc.createElement("91D1989D")
				sTmp = x_91D1989D
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 月度
				Set oXmlFld = oXmlDoc.createElement("67085EA6")
				sTmp = x_67085EA6
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 交期
				Set oXmlFld = oXmlDoc.createElement("4EA4671F")
				sTmp = x_4EA4671F
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 承办
				Set oXmlFld = oXmlDoc.createElement("627F529E")
				sTmp = x_627F529E
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 备注
				Set oXmlFld = oXmlDoc.createElement("59076CE8")
				sTmp = x_59076CE8
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)

				' Field 反馈报酬
				Set oXmlFld = oXmlDoc.createElement("53CD998862A5916C")
				sTmp = x_53CD998862A5916C
				if IsNull(sTmp) then sTmp = "<Null>"
				oXmlFld.Text = sTmp
				Call oXmlRec.AppendChild(oXmlFld)
				Set oXmlRec = Nothing
			End If
			If sExport = "csv" Then

				' Field ID
				sCsvStr = sCsvStr & """" & Replace(x_ID&"","""","""""") & """" & ","

				' Field 合同号
				sCsvStr = sCsvStr & """" & Replace(x_5408540C53F7&"","""","""""") & """" & ","

				' Field 客户名称
				sCsvStr = sCsvStr & """" & Replace(x_5BA26237540D79F0&"","""","""""") & """" & ","

				' Field 产品型号
				sCsvStr = sCsvStr & """" & Replace(x_4EA754C1578B53F7&"","""","""""") & """" & ","

				' Field 数量
				sCsvStr = sCsvStr & """" & Replace(x_657091CF&"","""","""""") & """" & ","

				' Field 价格
				sCsvStr = sCsvStr & """" & Replace(x_4EF7683C&"","""","""""") & """" & ","

				' Field 金额
				sCsvStr = sCsvStr & """" & Replace(x_91D1989D&"","""","""""") & """" & ","

				' Field 月度
				sCsvStr = sCsvStr & """" & Replace(x_67085EA6&"","""","""""") & """" & ","

				' Field 交期
				sCsvStr = sCsvStr & """" & Replace(x_4EA4671F&"","""","""""") & """" & ","

				' Field 承办
				sCsvStr = sCsvStr & """" & Replace(x_627F529E&"","""","""""") & """" & ","

				' Field 备注
				sCsvStr = sCsvStr & """" & Replace(x_59076CE8&"","""","""""") & """" & ","

				' Field 反馈报酬
				sCsvStr = sCsvStr & """" & Replace(x_53CD998862A5916C&"","""","""""") & """" & ","
				sCsvStr = Left(sCsvStr, Len(sCsvStr)-1) ' Remove last comma
				sCsvStr = sCsvStr & vbCrLf
			End If
		End If
		rs.MoveNext
	Loop

	' Close recordset and connection
	rs.Close
	Set rs = Nothing
	If sExport = "xml" Then
		Response.Write "<?xml version=""1.0"" encoding=""gb2312"" standalone=""yes""?>" & vbcrlf
		Response.Write oXmlTbl.xml
		Set oXmlTbl = Nothing
		Set oXmlDoc = Nothing
	End If
	If sExport = "csv" Then
		Response.Write sCsvStr
	End If
End Sub
%>
