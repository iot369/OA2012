<!--#INCLUDE FILE=inc_dbconn.asp-->
<%
my_conn.execute ("insert into function (show,cmd,xiang) values ('�Ż�����','/!!!','var_who�Ż����ŵĶ�var_to˵����������������ѽ����')")
my_conn.close
set my_conn=nothing
%>