<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<html>
<head>
<title>�ļ��ϴ�,��ѯ����  ,��ϵqq:�ͻ���ϵ����ϵͳ��ϵ���䣺  ,�ͻ�����qq:�ͻ���ϵ����ϵͳ---��ѯ�绰:(����)</title>
</head>
<body>
<%

dim upload,file,formName,formPath,iCount,filename,fileExt,ranNum
set upload=new upload_5xsoft ''�����ϴ�����
 session("flname")="" 

if upload.form("filepath")="" then   ''�õ��ϴ�Ŀ¼
 HtmEnd "������Ҫ�ϴ�����Ŀ¼!"
 set upload=nothing
 response.end
else
 formPath=upload.form("filepath")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 
end if

iCount=0
for each formName in upload.objForm ''�г�����form����
next
for each formName in upload.objFile ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
  if file.filesize<100 then
  response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
 response.end
 end if
  
 if file.filesize>5200000 then
  response.write "<font size=2>�ļ���С���������� 5M��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
 response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".zip" and fileEXT<>".xls" and fileEXT<>".rar" and  fileEXT<>".doc" and  fileEXT<>".txt"  then
  response.write "<font size=2>�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
 response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
 
' filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(FileName)   ''�����ļ�
           response.write("���ϴ��ļ�:")
		   response.write(FileName)
 		   session("flname")=CStr(filename)
      iCount=iCount+1
 else
  response.write "<font size=2 color=#0000ff>��ѡ����Ҫ�ϴ����ļ�!![<a href=""javascript:history.back();"">����</a>]</font>"
 end if

 set file=nothing
next
set upload=nothing  ''ɾ���˶���

sub HtmEnd(Msg)
 set upload=nothing
 response.write "<br>"&Msg&" <br>"
 response.end
end sub
%></body></html>