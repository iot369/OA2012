<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include FILE="upload_5xsoft.inc"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")

filepath="file/"
FileMaxSize="2500000"
fileweb="1"
nameset ="1"
pathset ="0"

function makefilename(fname)
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
  makefilename=fname
end function 
%>

<html>
<head>
<title>公文发送成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css" type="text/css">
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {font-size: 12px}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7 style8">公文传阅</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br><br><br><div align="center">
  <%
dim upload,file,formName,iCount
Dim FixFileExt
    Dim intfnN
	Dim FileExtName
    Dim FixFnN
	Dim intFix
	i=0
set upload=new upload_5xSoft ''建立上传对象
FixFileExt="asp|aspx|asa|asax|ascx|ashx|asmx|axd|cdx|cer|config|cs|csproj|licx|rem|resx|shtml|shtm|soap|stm|vb|vbproj|webinfo|cgi|pl|php|phtml|php3|jpg|gif|txt|"		'限制为只有这些文件可以上传(用"|"号格开)
iCount=0
for each formName in upload.file 
 set file=upload.file(formName)  
 if file.FileSize>0 then        
 if file.FileSize<FileMaxSize then 
FixFnN = Split(FixFileExt,"|")
intFix = Ubound(FixFnN)
			for i = 0 to intFix
				if GetExtendName(file.FileName) = lcase(trim(FixFnN(i))) then
  	Response.write "不支持您所上传的文件类型："
	Response.write GetExtendName(file.FileName)
	Response.write "<br>"
i=1
 exit for
 end if
next
if i=1 then
  	Response.write "附件传送失败!!"
 exit for
 end if 
 TypeFlag = 1
 
  if TypeFlag = 1 then 
  vfname = makefilename(now())
     if nameset = 1 then
  fname = vfname & iCount & "." & GetExtendName(file.FileName)
       elseif nameset =2 then
	   fname = file.FileName
	     elseif nameset = 3 then
		 fname = vfname & iCount & file.FileName
		 end if
		 
	Upfilepath= "file/"	 
  'response.write Upfilepath
  'response.write Server.mappath(UpFilePath&fname)
  file.SaveAs Server.mappath(UpFilePath&fname)  
  'response.write file.FilePath&file.FileName&" ("&file.FileSize&") => 上传附件成功! <br>"
  iCount=iCount+1
  FileNameStr = UpFilePath&fname
  
  if linkpath = "" then
  linkpath = fname
  else
  linkpath = request.form("file1")
  end if

end if

 else
 response.write "单个附件大小超出限制，您最多可以上传 "& FileMaxSize &"个字节的文件数据"
 exit for
 end if
 end if
 set file=nothing
 
next

sub HtmEnd(Msg)
 set upload=nothing
end sub


function GetExtendName(FileName)
dim ExtName
ExtName = LCase(FileName)
ExtName = right(ExtName,3)
ExtName = right(ExtName,3-Instr(ExtName,"."))
GetExtendName = ExtName
end function
%>
  <%
title=upload.form("title")
classify=upload.form("classify")
documenttype=upload.form("lb")
content=upload.form("content")
secretRB=upload.form("secretRB")
Degree=upload.form("Degree")
sendto=upload.form("sendto")


dim mysendto
mysendto=split(sendto,"|",-1,1)
for each sendtoinf in mysendto
userdeptpoint=InStr(sendtoinf,":")
if userdeptpoint>0 then
sendtoinflen=len(sendtoinf)
recipientusername=right(sendtoinf,sendtoinflen-userdeptpoint)
if recipientusername="所有人" then
recipientusername="所有人"
else
usernamepoint=Instr(recipientusername,"(")
usernamelen=len(recipientusername)
recipientusername=left(recipientusername,usernamelen-1)
recipientusername=right(recipientusername,usernamelen-1-usernamepoint)
end if
recipientuserdept=left(sendtoinf,userdeptpoint-1)
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("ADODB.recordset") 
sql = "select * from senddate"
rs.Open sql,conn,1,3
rs.addnew 
rs("title")=title
Rs("documenttype")=documenttype
Rs("content")=content
rs("sender")=oabusyusername
rs("recipientusername")=recipientusername
rs("recipientuserdept")=recipientuserdept
rs("inputdate")=date
rs("filename")=upload.form("file1")
rs.update 
rs.close 
set rs=nothing 
set conn=nothing 
end if
next

set upload=nothing  

%>
</div>
<table width="95%" height="27" align="center">
  <tr>
    <td height="23" align="center">
      <p>&nbsp;</p>
      <p>&nbsp;</p>
	  <%
	  if iCount>0 then
	  %>
	        <p align="center"><font color="#FF0000">公文发送成功！</font>
	  <%
	response.write "<br>成功地上传了附件!"
	  %></p>
	  <%
	  else
	response.write "<br>发送成功!"
    end if
	  %>
    </td>
  </tr>
</table>
</body>
</html>
