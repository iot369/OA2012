<!--#INCLUDE FILE=inc_dbconn.asp-->
<%
my_conn.execute ("insert into function (show,cmd,xiang) values ('慌慌张张','/!!!','var_who慌慌张张的对var_to说：“好象会出人命了呀！”')")
my_conn.close
set my_conn=nothing
%>