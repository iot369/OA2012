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
'�������ݿ����1ΪSQL���ݿ⣬0ΪAccess���ݿ�
Const IsSqlDataBase = 0
'��������ģʽ�����Ե�ʱ������1���������е�ʱ������Ϊ0,�����������Ϣ�����ڰ�ȫ��
Const IsDeBug = 1
'�������ƣ����ݳ������·���Զ�����
CacheName=Request.ServerVariables("url")
CacheName=left(CacheName,instrRev(CacheName,"/")-1)
'��¼ҳ�濪ʼִ��ʱ��
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
		'sql���ݿ����Ӳ��������ݿ������û����롢�û�������������������local�������IP��
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
		Response.Write "Error:���ݿ����ӳ���!"
		Response.End
	End If
End Sub
%>
<!-- #include file="cls_common.asp"-->
<!-- #include file="config.asp"-->