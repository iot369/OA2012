<!--#include file="uploadx.asp"-->
<!--#include file='connect.asp'-->
<html>
<head>
<title>�ϴ��ļ���.....</title>
</head>
<body>
<%
Dim filePath
Dim fileName
Dim fileExt
Dim file_subject
Dim Sql
Dim msg
Dim errflag
Dim errnumber
Dim SavePath
Dim maxfilesize

SavePath = "SavePath"									'����·��(���治Ҫ��"/"����)
maxfilesize = 1*5120									'��СΪ5M

Errflag=false
filePath = SavePath										'ʹ������·�����и�ֵ,��"/www"��"www"��
filePath = Server.MapPath(filePath)						'������·��ת��Ϊ����·��
file_subject = GetFormVal("filename")					'ȡ���ļ�����
fileext = GetFormVal("fileExt")							'ȡ���ļ�����
errnumber = GetFormVal("errnumber")						'ȡ�ñ���ʽ
errnumber = cint(errnumber)

if len(trim(file_subject))=0 then
	Response.Write "�ļ����ⲻ��Ϊ��"
	Response.End
end if
if len(trim(fileext))=0 then
	fileext = "�޼��"
end if

filename = SaveFile("fruit",filePath,maxfilesize,errnumber,1)	'���沢ȡ���ļ���
																'	0,1			Ψһ�ļ�����ʽ�������ͬ�����Զ�������
																'	1,1			����ʽ�������ͬ�������
																'	2,[0|1]		���Ƿ�ʽ�������ͬ���򸲸�ԭ�����ļ�

if len(trim(filename))>0 then
	Dim PerFnN
	Dim intPerFnN
	Dim PerFsize

	PerFnN=split(filename,"|")
	intPerFnN=Ubound(PerFnN)
	Select Case intPerFnN
		Case 1
			FileName=Trim(PerFnN(0))
			PerFsize=Csng(Trim(PerFnN(1)))
		Case 0
			FileName=Trim(PerFnN(0))
			PerFsize=0
	End Select

	select case Trim(filename)
		case "pathError"
			msg="����: ָ����·��������"
			errflag=true
		case "refileError"
			msg="����: �ļ��Ѿ�����"
			errflag=true
		case "sizeError"
			msg="����: �ļ�����ָ����С"
			errflag=true
		case "modeError"
			msg="�����ڲ�֧��Fsoģʽ�²��ܲ���Ψһ�򱨴�ʽ�ϴ��ļ�"
			errflag=true
		case "fileError"
			msg="�������ϴ����ļ���ʽ"
			errflag=true
		case else
			msg=""
			errflag=false
	end select
	if not errflag then
		Sql = "insert into upfile_table (subject,expit,filepath,filename,filesize) values"
		Sql = Sql& " ('"& file_subject &"','"& fileext &"','"& SavePath &"','"& filename &"',"& PerFsize &")"
		conn.execute(sql)
	end if
end if
conn.close
set conn=nothing

Response.Write "<script language='Jscript'>"&vbcrlf
Response.Write "<!--"&vbcrlf
if errflag then
	Response.Write "alert('"& msg &"');"&vbcrlf
end if
Response.Write "window.open('default.asp','_self');"
Response.Write "//-->"&vbcrlf
Response.Write "</script>"&vbcrlf

%>
</body>
</html>