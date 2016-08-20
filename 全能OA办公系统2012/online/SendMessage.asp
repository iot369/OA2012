<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Conn.asp"-->
<%

	  Content=request.form("Content")	
	  ToUserId=Request("ToUserId")  
	  '将消息加入消息表中
		strSQL="insert into Msg (Send,Receive,Content,DateAndTime) values('" & Request("FromUserId") & "','"&ToUserId&"','"&Content&"',now())"
		conn.execute strsql   
      response.write "<script language=JavaScript>{window.close();}</script>"

%>
