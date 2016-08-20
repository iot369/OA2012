<!--#include file="uploadx.asp"-->
<!--#include file='connect.asp'-->
<html>
<head>
<title>上传文件中.....</title>
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

SavePath = "SavePath"									'虚拟路径(后面不要加"/"符号)
maxfilesize = 1*5120									'大小为5M

Errflag=false
filePath = SavePath										'使用虚拟路径进行赋值,如"/www"或"www"等
filePath = Server.MapPath(filePath)						'将虚拟路径转换为磁盘路径
file_subject = GetFormVal("filename")					'取得文件标题
fileext = GetFormVal("fileExt")							'取得文件介绍
errnumber = GetFormVal("errnumber")						'取得报错方式
errnumber = cint(errnumber)

if len(trim(file_subject))=0 then
	Response.Write "文件主题不能为空"
	Response.End
end if
if len(trim(fileext))=0 then
	fileext = "无简介"
end if

filename = SaveFile("fruit",filePath,maxfilesize,errnumber,1)	'保存并取得文件名
																'	0,1			唯一文件名方式，如果有同名则自动改名；
																'	1,1			报错方式，如果有同名则出错；
																'	2,[0|1]		覆盖方式，如果有同名则覆盖原来的文件

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
			msg="错误: 指定的路径不存在"
			errflag=true
		case "refileError"
			msg="错误: 文件已经存在"
			errflag=true
		case "sizeError"
			msg="错误: 文件超出指定大小"
			errflag=true
		case "modeError"
			msg="主机在不支持Fso模式下不能采用唯一或报错方式上传文件"
			errflag=true
		case "fileError"
			msg="被限制上传的文件格式"
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