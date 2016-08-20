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

Function getUserList2(intLevel,intGroup)
    Dim rs,strUserList
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
	getUserList2 = strUserList
End Function

Function IsAccessUser(strUserList,strUser)
    Dim arrUserList,k,flag
	flag = 0
	arrUserList = Split(strUserList,",")
	For k = 0 To UBound(arrUserList) - 1
	    If Replace(arrUserList(k),"'","") = strUser Then
		    flag = 1
			Exit For
		End If
	Next
	If flag = 1 Then
	    IsAccessUser = True
	Else
	    IsAccessUser = False
	End If
End Function

Function getGroupUserList(intGroup)
    Dim rs,strUserList
	strUserList = "'" & Session("CRM_name") & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uGroup = " & intGroup,conn,3,1
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
	getGroupUserList = strUserList
End Function

Function getClientsList(strSql)
    Dim rs,strClientsList
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSql,conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If strClientsList = "" Then
		    strClientsList = rs("cId")
		Else
		    strClientsList = strClientsList & "," & rs("uName")
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	getClientsList = strClientsList
End Function

Dim strCounter,strToPrint
strToPrint = strToPrint & "        <tr>" & VBCrlf
strToPrint = strToPrint & "          <td width=""100"" align=""center"" bgcolor=""menu"">�ݷ�����</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">�ݷ�����</td>" & VBCrlf
strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">�ͻ�����</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">�ͻ��ȼ�</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">��ҵ����</td>" & VBCrlf
strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">ҵ��Ա</td>" & VBCrlf
strToPrint = strToPrint & "        </tr>" & VBCrlf

''''''''''''''''''''''''''''''''''''''
Function getClientsItem(cId,s)
    Dim rs,itemValue
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_client Where cId = " & cId,conn,3,1
	If rs.RecordCount = 0 Then
	    itemValue = ""
	Else
	    itemValue = rs(s)
	End If
	rs.Close
	Set rs = Nothing
	getClientsItem = itemValue
End Function

Function listAll(mySql)
    Dim rs,strOut(2),strUserList
	Dim intTotalRecords,intTotalPages,intCurrentPage,intPageSize
	intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 20

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open mySql,conn,3,1
	intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
    strOut(0) = strOut(0) & "�� " & intTotalRecords & " ����¼ "
    strOut(0) = strOut(0) & "�� " & intTotalPages & " ҳ "
    strOut(0) = strOut(0) & "��ǰ�� " & intCurrentPage & " ҳ "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=1""><<��ҳ</a> "
    Else
        strOut(0) = strOut(0) & "<<��ҳ "
    End If
    If intCurrentPage > 1 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
    Else
        strOut(0) = strOut(0) & "<��һҳ "
    End If
    If intCurrentPage < intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
    Else
        strOut(0) = strOut(0) & "��һҳ> "
    End If
    If intCurrentPage <> intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
    Else
        strOut(0) = strOut(0) & "βҳ>>"
    End If
	
	Dim k
	k = 0
	Do While Not rs.BOF And Not rs.EOF
	    k = k + 1
	    strOut(1) = strOut(1) & "        <tr>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td align=""center"">" & rs("rDate") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & rs("rType") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & getClientsItem(rs("cId"),"cCompany") & "</a></td>" &  VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cType") & "</td>" & VBCrlf
		strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cTrade") & "</td>" & VBCrlf
		'If Session("CRM_level") = 9 Then
	        strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cUser") & "</td>" & VBCrlf
		'End If
	    strOut(1) = strOut(1) & "        </tr>" & VBCrlf
		If k >= intPageSize Then Exit Do
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	listAll = strOut
End Function

Function listAllAtDate(mySql,intDay)
    Dim rs,strOut(2),strUserList
	Dim intTotalRecords,intTotalPages,intCurrentPage,intPageSize
	intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 20

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open mySql,conn,3,1
	intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
	strOut(0) = strOut(0) & "�� " & intTotalRecords & " ����¼ "
	If intDay <> "" Then
	    strOut(0) = strOut(0) & "��" & intDay & " �� "
	    strOut(0) = strOut(0) & "ƽ�� " & FormatNumber((intTotalRecords / intDay),1,-1) & " ����¼/ �� "
	End If
    strOut(0) = strOut(0) & "�� " & intTotalPages & " ҳ "
    strOut(0) = strOut(0) & "��ǰ�� " & intCurrentPage & " ҳ "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=1""><<��ҳ</a> "
    Else
        strOut(0) = strOut(0) & "<<��ҳ "
    End If
    If intCurrentPage > 1 Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage - 1 & """><��һҳ</a> "
    Else
        strOut(0) = strOut(0) & "<��һҳ "
    End If
    If intCurrentPage < intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intCurrentPage + 1 & """>��һҳ></a> "
    Else
        strOut(0) = strOut(0) & "��һҳ> "
    End If
    If intCurrentPage <> intTotalPages Then
        strOut(0) = strOut(0) & "<a href=""?pageNum=" & intTotalPages & """>βҳ>></a>"
    Else
        strOut(0) = strOut(0) & "βҳ>>"
    End If
	
	Dim k
	k = 0
	Do While Not rs.BOF And Not rs.EOF
	    k = k + 1
	    strOut(1) = strOut(1) & "        <tr>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td align=""center"">" & rs("rDate") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & rs("rType") & "</td>" & VBCrlf
	    strOut(1) = strOut(1) & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & getClientsItem(rs("cId"),"cCompany") & "</a></td>" &  VBCrlf
	    strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cType") & "</td>" & VBCrlf
		strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cTrade") & "</td>" & VBCrlf
		'If Session("CRM_level") = 9 Then
	        strOut(1) = strOut(1) & "        <td>" & getClientsItem(rs("cId"),"cUser") & "</td>" & VBCrlf
		'End If
	    strOut(1) = strOut(1) & "        </tr>" & VBCrlf
		If k >= intPageSize Then Exit Do
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	listAllAtDate = strOut
End Function

Dim actionUrl
If Session("CRM_level") >= 9 Then
    actionUrl = "?action=adminAction"
Else
    actionUrl = "?action=userAction"
End If

Dim action,selectItems,rUser,arrList
Dim sql1,intDay
action = Trim(Request.QueryString("action"))
selectItems = Trim(Request.Form("selectItems"))
rUser = Trim(Request.Form("rUser"))

If action <> "" Then Session("CRM_sql1") = ""

If action = "userAction" Then
    Select Case selectItems
	Case "rTime"
	    Dim rTimeBegin,rTimeEnd
		rTimeBegin = Trim(Request.Form("rTimeBegin"))
		rTimeEnd = Trim(Request.Form("rTimeEnd"))
		If rTimeBegin = "" And rTimeEnd = "" Then Response.Redirect("?errMsg=1")
		If rTimeBegin <> "" And rTimeEnd <> "" Then
		    ''���ݲ�����
		    If Not IsDate(rTimeBegin) Or Not IsDate(rTimeEnd) Or rTimeBegin > rTimeEnd Then Response.Redirect("?errMsg=1")
			intDay = DateDiff("d",rTimeBegin,rTimeEnd)
			If intDay = 0 Then intDay = 1
			If rUser = "" Then
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
                arrList = listAllAtDate(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			Else
			    ''û��Ȩ��
			    If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
                arrList = listAllAtDate(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			End If
		ElseIf rTimeBegin <> "" Then
		    ''���ݲ�����
		    If Not IsDate(rTimeBegin) Then Response.Redirect("?errMsg=1")
			If rUser = "" Then
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And rDate >= #" & rTimeBegin & "#"
                arrList = listAll(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			Else
			    ''û��Ȩ��
			    If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "#"
                arrList = listAll(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			End If
		Else
		    ''���ݲ�����
		    If Not IsDate(rTimeEnd) Then Response.Redirect("?errMsg=1")
			If rUser = "" Then
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And rDate <= #" & rTimeEnd & "#"
                arrList = listAll(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			Else
			    ''û��Ȩ��
			    If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate <= #" & rTimeEnd & "#"
                arrList = listAll(sql1)
            	strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
			End If
		End If
	Case "rType"
	    Dim rType
		rType = Trim(Request.Form("rType"))
		If rType = "" Then Response.Redirect("?errMsg=1")
		If rUser = "" Then
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And cId In (Select cId From baidu_client Where cType = '" & rType & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		Else
			''û��Ȩ��
			If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")		
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And cId In (Select cId From baidu_client Where cType = '" & rType & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		End If
	Case "rTrade"
	    Dim rTrade
		rTrade = Trim(Request.Form("rTrade"))
		If rTrade = "" Then Response.Redirect("?errMsg=1")
		If rUser = "" Then
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And cId In (Select cId From baidu_client Where cTrade = '" & rTrade & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		Else
			''û��Ȩ��
			If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")		
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And cId In (Select cId From baidu_client Where cTrade = '" & rTrade & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		End If
	Case "rRecordsType"
	    Dim rRecordsType
		rRecordsType = Trim(Request.Form("rRecordsType"))
		If rRecordsType = "" Then Response.Redirect("?errMsg=1")
		If rUser = "" Then
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList(Session("CRM_level"),Session("CRM_group")) & ")) And rType  = '" & rRecordsType & "'"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		Else
			''û��Ȩ��
			If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")		
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rType = '" & rRecordsType & "'"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
		End If
	End Select
ElseIf action = "adminAction" Then
    Dim rGroup
	rGroup = Request.Form("rGroup")
    If rGroup = "" Then
    	Select Case selectItems
    	Case "rTime"
    	    'Dim rTimeBegin,rTimeEnd
    		rTimeBegin = Trim(Request.Form("rTimeBegin"))
    		rTimeEnd = Trim(Request.Form("rTimeEnd"))
    		If rTimeBegin = "" And rTimeEnd = "" Then Response.Redirect("?errMsg=1")
    		If rTimeBegin <> "" And rTimeEnd <> "" Then
	    	    ''���ݲ�����
	    	    If Not IsDate(rTimeBegin) Or Not IsDate(rTimeEnd) Or rTimeBegin > rTimeEnd Then Response.Redirect("?errMsg=1")
				intDay = DateDiff("d",rTimeBegin,rTimeEnd)
				If intDay = 0 Then intDay = 1
	    		If rUser = "" Then
	    		    sql1 = "Select * From baidu_records Where rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
                    arrList = listAllAtDate(sql1,intDay)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
	    		Else
	    		    ''û��Ȩ��
	    		    'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
	    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
                    arrList = listAllAtDate(sql1,intDay)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
	    		End If
	    	ElseIf rTimeBegin <> "" Then
	    	    ''���ݲ�����
	    	    If Not IsDate(rTimeBegin) Then Response.Redirect("?errMsg=1")
	    		If rUser = "" Then
    			    sql1 = "Select * From baidu_records Where rDate >= #" & rTimeBegin & "#"
                    arrList = listAll(sql1)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
    			Else
    			    ''û��Ȩ��
    			    'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
    			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "#"
                    arrList = listAll(sql1)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
    			End If
    		Else
    		    ''���ݲ�����
    		    If Not IsDate(rTimeEnd) Then Response.Redirect("?errMsg=1")
    			If rUser = "" Then
    			    sql1 = "Select * From baidu_records Where rDate <= #" & rTimeEnd & "#"
                    arrList = listAll(sql1)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
    			Else
    			    ''û��Ȩ��
    			    'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
    			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate <= #" & rTimeEnd & "#"
                    arrList = listAll(sql1)
                	strToPrint = strToPrint & arrList(1)
                    strCounter = arrList(0)
    			End If
    		End If
    	Case "rType"
    	    'Dim rType
    		rType = Trim(Request.Form("rType"))
    		If rType = "" Then Response.Redirect("?errMsg=1")
    		If rUser = "" Then
    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cType = '" & rType & "')"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		Else		
    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "' And cType = '" & rType & "')"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		End If
    	Case "rTrade"
    	    'Dim rTrade
    		rTrade = Trim(Request.Form("rTrade"))
    		If rTrade = "" Then Response.Redirect("?errMsg=1")
    		If rUser = "" Then
    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cTrade = '" & rTrade & "')"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		Else		
    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "' And cTrade = '" & rTrade & "')"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		End If
    	Case "rRecordsType"
    	    'Dim rRecordsType
    		rRecordsType = Trim(Request.Form("rRecordsType"))
    		If rRecordsType = "" Then Response.Redirect("?errMsg=1")
    		If rUser = "" Then
    		    sql1 = "Select * From baidu_records Where rType  = '" & rRecordsType & "'"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		Else		
    		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rType = '" & rRecordsType & "'"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		End If
    	End Select
    Else
	    If CInt(Abs(rGroup)) <= 0 Then Response.Redirect("?errMsg=1")
    	Select Case selectItems
    	Case "rTime"
    	    'Dim rTimeBegin,rTimeEnd
    		rTimeBegin = Trim(Request.Form("rTimeBegin"))
    		rTimeEnd = Trim(Request.Form("rTimeEnd"))
    		If rTimeBegin = "" And rTimeEnd = "" Then Response.Redirect("?errMsg=1")
    		If rTimeBegin <> "" And rTimeEnd <> "" Then
	    	    ''���ݲ�����
	    	    If Not IsDate(rTimeBegin) Or Not IsDate(rTimeEnd) Or rTimeBegin > rTimeEnd Then Response.Redirect("?errMsg=1")
				intDay = DateDiff("d",rTimeBegin,rTimeEnd)
				If intDay = 0 Then intDay = 1
	    		''û��Ȩ��
	    		'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
	    '		sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "# And rDate <= #" & rTimeEnd & "#"
                arrList = listAllAtDate(sql1,intDay)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
	    	ElseIf rTimeBegin <> "" Then
	    	    ''���ݲ�����
	    	    If Not IsDate(rTimeBegin) Then Response.Redirect("?errMsg=1")
    			''û��Ȩ��
    			'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And rDate >= #" & rTimeBegin & "#"
    			'sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate >= #" & rTimeBegin & "#"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		Else
    		    ''���ݲ�����
    		    If Not IsDate(rTimeEnd) Then Response.Redirect("?errMsg=1")
    			''û��Ȩ��
    			'If Not IsAccessUser(getUserList(Session("CRM_level"),Session("CRM_group")),rUser) Then Response.Redirect("?errMsg=2")
			    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And rDate <= #" & rTimeEnd & "#"
    			'sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rDate <= #" & rTimeEnd & "#"
                arrList = listAll(sql1)
                strToPrint = strToPrint & arrList(1)
                strCounter = arrList(0)
    		End If
    	Case "rType"
    	    'Dim rType
    		rType = Trim(Request.Form("rType"))
    		If rType = "" Then Response.Redirect("?errMsg=1")
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And cId In (Select cId From baidu_client Where cType = '" & rType & "')"
    		'sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And cId In (Select cId From baidu_client Where cType = '" & rType & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
    	Case "rTrade"
    	    'Dim rTrade
    		rTrade = Trim(Request.Form("rTrade"))
    		If rTrade = "" Then Response.Redirect("?errMsg=1")
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And cId In (Select cId From baidu_client Where cTrade = '" & rTrade & "')"
    		'sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And cId In (Select cId From baidu_client Where cTrade = '" & rTrade & "')"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
    	Case "rRecordsType"
    	    'Dim rRecordsType
    		rRecordsType = Trim(Request.Form("rRecordsType"))
    		If rRecordsType = "" Then Response.Redirect("?errMsg=1")
		    sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser In (" & getUserList2(9,rGroup) & ")) And rType  = '" & rRecordsType & "'"
    		'sql1 = "Select * From baidu_records Where cId In (Select cId From baidu_client Where cUser = '" & rUser & "') And rType = '" & rRecordsType & "'"
            arrList = listAll(sql1)
            strToPrint = strToPrint & arrList(1)
            strCounter = arrList(0)
    	End Select
    End If
End If
If sql1 <> "" Then
    Session("CRM_sql1") = sql1
	If intDay <> "" Then Session("CRM_intDay") = intDay
Else
    Dim pageNum
	pageNum = Request.QueryString("pageNum")
	If pageNum <> "" Then
        arrList = listAllAtDate(Session("CRM_sql1"),Session("CRM_intDay"))
        strToPrint = strToPrint & arrList(1)
        strCounter = arrList(0)
	End If
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ͻ���ϵ����ϵͳ</title>
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

function changeItem()
{
    var items01 = "��ʼ���ڣ�";
    items01 = items01 + " <input name=\"rTimeBegin\" type=\"text\" id=\"rTimeBegin\" value=\"<% = Date() %>\" size=\"12\" maxlength=\"12\" onClick=\"this.value='';\">";
    items01 = items01 + "----�������ڣ�";
    items01 = items01 + " <input name=\"rTimeEnd\" type=\"text\" id=\"rTimeEnd\" value=\"<% = Date() %>\" size=\"12\" maxlength=\"12\" onClick=\"this.value='';\">";
    var items02 = "�ͻ��ȼ���";
    var items02 = items02 + " <% = Replace(getList(1,"baidu_clientsType","","clientsType","rType","�ͻ��ȼ�"),"""","\""") %>";
    var items03 = "��ҵ���ͣ�";
    var items03 = items03 + " <% = Replace(getList(1,"baidu_clientsTrade","","clientsTrade","rTrade","��ҵ����"),"""","\""") %>";
	var items04 = "�ݷ����ͣ�";
    items04 = items04 + " <% = Replace(getList(1,"baidu_recordsType",,"recordsType","rRecordsType","�ݷ�����"),"""","\""") %>";

    var items = document.all.selectItems.value;
	switch(items){
	case "rTime":
	    document.all.dataFormItems.innerHTML = items01;
		document.all.items.value = items;
		return;
	case "rType":
	    document.all.dataFormItems.innerHTML = items02;		
		document.all.items.value = items;
		return;
	case "rTrade":
	    document.all.dataFormItems.innerHTML = items03;
		document.all.items.value = items;
		return;
	case "rRecordsType":
	    document.all.dataFormItems.innerHTML = items04;
		document.all.items.value = items;
		return;
	case "":
	    alert("��ѡ�񱨱��������͡�");
		document.all.selectItems.focus();
		return;
	}
	//if(items != ""){
	//    if(items == "rTime"){
	//	    document.all.dataFormItems.innerHTML = items01;
	//	}
	//	else{
	//	    if(items == "rType"){
	//		    document.all.dataFormItems.innerHTML = items02;
	//		}
	//		else{
	//		    document.all.dataFormItems.innerHTML = items03;
	//		}
	//	}
	//}
	//else{
	//    alert("��ѡ�񱨱��������͡�");
	//	document.all.selectItems.focus();
	//	return false;
	//}
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
<%' Response.Write(sql1 & selectItems & rTimeBegin & rTimeEnd) %>
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
        <form name="dataForm" action="<% = actionUrl %>" method="post">
		<tr> 
          <td align="right">&nbsp;</td>
          <td>�������ݣ� <select name="selectItems" id="selectItems" onChange="return changeItem();">
              <option value="">��ѡ��</option>
              <option value="rTime">�ݷ�ʱ��</option>
              <option value="rType">�ͻ��ȼ�</option>			  
              <option value="rTrade">��ҵ����</option>
              <option value="rRecordsType">�ݷ�����</option>
            </select>
            <input name="items" type="hidden" id="items"></td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td id="dataFormItems">&nbsp; </td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;�û���
            <input name="rUser" type="text" id="rUser" size="16" maxlength="16">
              <% If Session("CRM_level") >= 9 Then %>ҵ���飺<% = getList(2,"baidu_group","gId","gName","rGroup","ҵ����") %><% End If %>
              ���û�������ΪȨ���������û��� </td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td align="center"><input type="submit" name="Submit" value="�ύ"> &nbsp;&nbsp; 
            <input name="Reset" type="reset" id="Reset" value="����"></td>
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
</body>
</html>
