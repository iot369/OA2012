<!--#INCLUDE FILE="inc_dbconn.asp" -->
<%
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_manager&"=2  where "&dbfield_user_username&"='admin'")
my_conn.close
set my_conn=nothing
%>
