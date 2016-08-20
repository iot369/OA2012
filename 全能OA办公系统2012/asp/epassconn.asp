<%
'DB_name="Data Source=epass;User ID=epassmanager; Password="
DB_name="Data Source=epass;User ID=; Password="
set conn=server.createobject("ADODB.CONNECTION")
conn.open DB_name
%>
