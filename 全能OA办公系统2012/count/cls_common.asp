<%
class Cls_CuteCounter
	Public BaseUrl
	Public WebName,WebNameE,WebUrl,SysName,SysNameE,SysVersion
	Public rs


	Private Sub Class_Initialize()
		BaseUrl = "http://"&LCase(Replace(Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL"),Split(Request.ServerVariables("SCRIPT_NAME"),"/")(ubound(Split(Request.ServerVariables("SCRIPT_NAME"),"/"))),""))
		WebName="IT学习者"
		WebNameE="ITlearner"
		WebUrl="http://www.itlearner.com"
		SysName="网站访问统计系统"		
		SysNameE="CuteCounter"
		SysVersion="V1.6"

		if Application.Contents(CacheName & "_isStart")="" then Application.Contents(CacheName & "_isStart")=1	
	End Sub

	Private Sub class_terminate()
		If IsObject(Conn) Then 
			Conn.Close
			Set Conn = Nothing
		End If 
	End Sub

	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase	
		On Error Resume Next
		Set Execute = Conn.Execute(Command)
		If Err Then
			If IsDeBug = 1 Then
				Response.Write Lang.item("g_054") & Command
				Response.Write Lang.item("g_055") & Err.description
			Else
				Response.Write Lang.item("g_056")
			End If
			Err.Clear
			Response.End
		End If	
	End Function
	
	Public Function Getrs(Sql,num1,num2)
		On Error Resume Next
		If Not IsObject(Conn) Then ConnectionDatabase
		Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open Sql,Conn,num1,num2
		Set Getrs = rs
		If Err Then
			If IsDeBug = 1 Then
				Response.Write Lang.item("g_054") & Command
				Response.Write Lang.item("g_055") & Err.description
			Else
				Response.Write Lang.item("g_056")
			End If
			Err.Clear
			Response.End
		End If	
	End Function

	Public Sub ShowFooter()
		dim Endtime,Runtime,OutStr
		Endtime=timer()
		dim WebName,WebUrl,rs
		set rs=Execute("select WebName,WebUrl from Webinfo where ID=1")
		OutStr = "<p align=center>"
		OutStr = OutStr & "Copyright &copy; " &Year(Date())& "  <a href="&rs(1)&">"&rs(0)&"</a> All Rights Reserved <br>"
		set rs=nothing

		Runtime=FormatNumber((endtime-startime)*1000,2) 
		if Runtime>0 then
			if Runtime>=1000 then
				OutStr = OutStr & Lang.item("g_019") & FormatNumber(runtime/1000,2) & Lang.item("g_021")
			else
				OutStr = OutStr & Lang.item("g_019") & Runtime & Lang.item("g_020")
			end if	
		end if
		OutStr = OutStr & "&nbsp;&nbsp;"
		OutStr = OutStr & "<a href=""http://www.it" + "learner.com/CuteCounter/"" target=_blank>ITlearner CuteCounter "&SysVersion&"</a>"				
		OutStr = OutStr & "</p>"
		Response.Write(OutStr)
	End Sub

	Public Function twonum(num)
		if len(num)=1 then
			twonum="0"&num
		else
			twonum=num
	   	end if
	End Function

	Public Function Checkstr(Str,length)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		'	CheckStr = trim(Replace(Str,"'","''"))
		'if instr(Str,"%27") then
		'	CheckStr = trim(Replace(Str,"%27","''"))
		'End if
		CheckStr = server.HTMLEncode(Str)
		if length>0 and strlength(CheckStr)>length then
			CheckStr= Strleft(CheckStr,length)
		End if
	End Function


	Public Function htmlencode2(str)
		htmlencode2=replace(str,chr(10),"&nbsp;")
		htmlencode2=replace(htmlencode2,chr(13),"&nbsp;")
		htmlencode2=replace(htmlencode2,chr(32),"&nbsp;")
	End Function
	
	Public Function Strlength(Str)
		dim Temp_Str,I,Test_Str
		Temp_Str=Len(Str)
		For I=1 To Temp_Str
			Test_Str=(Mid(Str,I,1))
			If Asc(Test_Str)>0 Then
				Strlength=Strlength+1
			Else
				Strlength=Strlength+2
			End If
		Next
	End Function
	
	Public Function Strleft(Str,L)
		dim Temp_Str,I,lens,Test_Str
		Temp_Str=Len(Str)
		For I=1 To Temp_Str
			Test_Str=(Mid(Str,I,1))
			Strleft=Strleft&Test_Str
			If Asc(Test_Str)>0 Then
				lens=lens+1
			Else
				lens=lens+2
			End If
				If lens>=L Then Exit For
		Next
	End Function
	
	Public Function OutStr(Str,L)
		if Strlength(Str)>L then
			OutStr=StrLeft(Str,L)
			OutStr=OutStr & ".."
		else
			OutStr=Str
		end if
	End Function

	Public Function GetSearchKeyword(RefererUrl)	'搜索关键词
		on error resume next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		Dim a,b,j
		'模糊查找关键词，此方法速度较快，范围也较大
		re.Pattern = "(word=([^&]*)|q=([^&]*)|p=([^&]*)|query=([^&]*)|name=([^&]*)|_searchkey=([^&]*)|wd=([^&]*)|baidu.*?w=([^&]*))"
		Set a = re.Execute(RefererUrl)
		If a.Count>0 then
			Set b = a(a.Count-1).SubMatches
			For j=1 to b.Count
				If Len(b(j))>0 then GetSearchKeyword=b(j) : Exit Function
			Next
		End If
		if err then
		err.clear
		GetSearchKeyword = RefererUrl
		else
		GetSearchKeyword = ""		
		end if		
	End Function
End class
%>