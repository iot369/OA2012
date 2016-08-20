<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<html>
<head>
<title>文件上传,咨询邮箱  ,联系qq:客户关系管理系统联系邮箱：  ,客户服务qq:客户关系管理系统---咨询电话:(短信)</title>
</head>
<body>
<%

dim upload,file,formName,formPath,iCount,filename,fileExt,ranNum
set upload=new upload_5xsoft ''建立上传对象
 session("flname")="" 

if upload.form("filepath")="" then   ''得到上传目录
 HtmEnd "请输入要上传至的目录!"
 set upload=nothing
 response.end
else
 formPath=upload.form("filepath")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 
end if

iCount=0
for each formName in upload.objForm ''列出所有form数据
next
for each formName in upload.objFile ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
  if file.filesize<100 then
  response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
 response.end
 end if
  
 if file.filesize>5200000 then
  response.write "<font size=2>文件大小超过了限制 5M　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
 response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".zip" and fileEXT<>".xls" and fileEXT<>".rar" and  fileEXT<>".doc" and  fileEXT<>".txt"  then
  response.write "<font size=2>文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
 response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
 
' filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(FileName)   ''保存文件
           response.write("已上传文件:")
		   response.write(FileName)
 		   session("flname")=CStr(filename)
      iCount=iCount+1
 else
  response.write "<font size=2 color=#0000ff>请选择你要上传的文件!![<a href=""javascript:history.back();"">返回</a>]</font>"
 end if

 set file=nothing
next
set upload=nothing  ''删除此对象

sub HtmEnd(Msg)
 set upload=nothing
 response.write "<br>"&Msg&" <br>"
 response.end
end sub
%></body></html>