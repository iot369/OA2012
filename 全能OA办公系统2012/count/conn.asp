<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
Dim Startime
Dim CacheName
Dim hx
Dim db
		db="data/#ITlearner.asp"
Dim SqlNowString,SqlDateString,TimeDiff(3),Conn
'定义数据库类别，1为SQL数据库，0为Access数据库
Const IsSqlDataBase = 0
'定义运行模式，测试的时候设置1，正常运行的时候设置为0,不输出错误信息有利于安全，
Const IsDeBug = 1
'缓存名称，根据程序放置路径自动生成
CacheName=Request.ServerVariables("url")
CacheName=left(CacheName,instrRev(CacheName,"/")-1)
'记录页面开始执行时间
Startime = Timer()

Set hx = New Cls_CuteCounter
If IsSqlDataBase = 1 Then
	SqlNowString = "GetDate()"
	TimeDiff(0)="n"
	TimeDiff(1)="hh"	
	SqlDateString = "'"&Date()&"'"
Else
	SqlNowString = "Now()"
	TimeDiff(0)="'n'"
	TimeDiff(1)="'h'"
	SqlDateString = "Date()"	
End If

Sub ConnectionDatabase
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		'sql数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
    connstr = "Provider=SQLOLEDB.1;Password='';Persist Security Info=True;User ID='';Initial Catalog='';Data Source=''"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "Error:数据库连接出错!"
		Response.End
	End If
End Sub
%>
<!-- #include file="cls_common.asp"-->
<!-- #include file="config.asp"-->