<!--#include file="upload.inc"-->

<html>
<head>
<title>�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	background-color: #CCCCCC;
}
-->
</style>
<link href="../common/style.css" rel="stylesheet" type="text/css">
</head>
<body topmargin=0>
<div align="left">
  <%
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����

 formPath=upload.form("filepath")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body>"

iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
     response.write "��ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
    response.end
 end if
     
 'if file.filesize>100*99000 then
  '   response.write "�ļ���С����������100K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
   ' response.end
' end if

 fileExt=lcase(right(file.filename,4))
 uploadsuc=false
Forum_upload="doc,xls,gif,jpg,png,swf,wav,mp3"
 Forumupload=split(Forum_upload,",")
 for i=0 to ubound(Forumupload)
    if fileEXT="."&trim(Forumupload(i)) then
    uploadsuc=true
    exit for
    else
    uploadsuc=false
    end if
 next
 if uploadsuc=false then
     response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
    response.end
 end if

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(FileName)   ''�����ļ�
    for i=0 to ubound(Forumupload)
        if fileEXT="."&trim(Forumupload(i)) then

        exit for
        end if
    next
 iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���

Htmend iCount&" ���ļ��ϴ�����!"

sub HtmEnd(Msg)
 set upload=nothing
response.write right(FileName,22)
response.end
end sub

%>
</div>
</body>
</html>
