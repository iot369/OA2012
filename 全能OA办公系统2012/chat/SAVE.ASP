<!--#INCLUDE FILE="inc_dbconn.asp" -->
<!--#INCLUDE FILE="inc_config.asp" -->
<% if session("user")<>"" then
      i=session("id")
      if application("intime"&i)<>"" then
         oldtime=cdate(application("intime"&i))
      else
         oldtime=now
      end if
      addrate=int(datediff("s",oldtime,now())/60)
      my_conn.execute "update "&dbtable_user&" set "&dbfield_user_rate&"='"&session("oldrate")+addrate&"' where "&dbfield_user_username&"='"&session("user")&"'"
      my_conn.close
      set my_conn=nothing
      application("intime"&i)=now
      session("oldrate")=session("oldrate")+addrate
%>
<html>
<head>
<title>保存积分</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>
<body bgcolor="#cccccc" text="<%=m_text2%>" style="font-size:10.4pt;border:0">
<br><center>您的积分存入 <%=addrate%> 分<br><br>
<input type="button" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" value="关闭" onclick="javascript:window.close();"></center>
</body>
</html>
<% end if %>