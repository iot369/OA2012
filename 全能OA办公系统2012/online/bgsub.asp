<%
imgpath="./"
sub bghead()
%>

<head>
</head>

<body bgcolor="#EBEBEB" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" height="5">
  <tr> 
    <td bgcolor="#ffffff" width="100%" height="100%"></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td height=67>
        <table border="0" cellpadding="0" cellspacing="0" width="610" height="67">
          <tr>
            <td width=63><img border="0" src="<%=imgpath%>images/j1.gif"></td>
            <td background="<%=imgpath%>images/j2.gif" align=center>
<%
end sub
%>

<%
sub bgmid()
%>
</td>
            <td width="60"><img border="0" src="<%=imgpath%>images/j3.gif"></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td valign=top>
        <table border="0" cellpadding="0" cellspacing="0" width="610">
          <tr>
            <td background="<%=imgpath%>images/j4.gif" width=17 height=280></td>
            
          <td valign=top height=100%> 
            <%
end sub
%>
            <%
sub bgback()
%>
  <!--        </td>
            
 
          </tr>
        </table>
    </tr>
    <tr>
      <td height=19>
        <table border="0" cellpadding="0" cellspacing="0" width="610" height="19">
          <tr>
            
           
            
           
            
           
          </tr>
		  <tr>
          <td width="100%" height=20 colspan="3"><center>-----<font color="#808080">JZUND-OA &copy; <a href="http://www.ie37.com" target="_blank">COMPANY.JZUD.COM</a></font></a>-----</center></td>
		  </tr>
        </table>
      </td>
    </tr>
  </table>-->

<%
end sub
'�����ַ�����ʵ�ʳ���
function strlength(inputstr)
	dim length,i
	length=0
	for i=1 to len(inputstr)
		if asc(mid(inputstr,i,1))<0 then
			length=length+2
		else
			length=length+1
		end if
	next
	strlength=length
end function
'����ҳ��ʽ��ʾ������Ϣ��ȷ����ִ�к��˲���
Sub DispErrorInfo1(ErrorInfoStr)
	Response.Write("<div align=""center"">")
	Response.Write("<br><br>")
	Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""0"">")
    Response.Write("<tr>")
	Response.Write("<td><img border=""0"" src=""/Electron_Doc/Images/errorico.gif"" align=""absmiddle"" width=""32"" height=""32"">��<font size=""3"" color=""#FF0000"">"&ErrorInfoStr&"</font></td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	Response.Write("<br><br>")
	Response.Write("<input type=""button"" Value="" �� �� "" onclick=""javascript:history.go(-1);"">")
	Response.Write("<br><br>")
	Response.Write("</div>")
End Sub
%>
