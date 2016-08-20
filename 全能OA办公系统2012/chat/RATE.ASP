<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title><%=r_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body      {  font-size: 10.4pt }
td        {  font-size: 10.4pt }
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>

<body bgcolor="#cccccc" text="<%=m_text2%>" style="margin:0">
<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr bgcolor="<%=m_bg1%>"> 
      
    <td colspan="3" style="color:<%=m_text1%>">&nbsp;&nbsp;
      <div align="center"><font class=3dfont color=#0000ff><%=r_title%> 排 行 榜</font> 
        &nbsp; 
        <select onchange="location=this.options[this.selectedIndex].value" style="font-size:10pt">
          <option>请选择</option>
          <option value="rate.asp?toop=10">前10名</option>
          <option value="rate.asp?toop=20">前20名</option>
          <option value="rate.asp?toop=30">前30名</option>
          <option value="rate.asp?toop=40">前40名</option>
          <option value="rate.asp?toop=50">前50名</option>
        </select>
      </div>
    </td>
  </tr>
    <tr>
    <td colspan="3" height="14"></td>
  </tr>
  <tr>
    <td width="40" align="center">名次</td>
	<td width="100" align="center">姓名</td>
	<td width="200"align="center">积分</td>
  </tr>  
<%
i=0
set rs=my_conn.execute("select * from "&dbtable_user&" order by "&dbfield_user_rate&" desc")
    rs.movefirst
    raterules=rs(dbfield_user_rate)
    toop=request("toop")
    if toop="" then toop=10
    do while not rs.eof
       i=i+1
       if toop+1=i then exit do
       nowrate=rs(dbfield_user_rate)
       if raterules=0 then raterules=1
       displayrate=nowrate*100/raterules
       %>
  <tr>
    <td align="center">〖<%=i%>〗&nbsp;</td>
	<td align="center"><%=rs(dbfield_user_username)%></td>
	<td><img src="img/X.jpg" width="<%=(displayrate/100)*200%>" height="10"><%=rs(dbfield_user_rate)%></td>
  </tr>
       <%
       rs.movenext
    loop
rs.close
set rs=nothing
my_conn.close
set my_conn=nothing
%>
  </tr>
    <tr>
    <td colspan="3" height="14"></td>
  </tr>
  <tr align="right"> 
    <td colspan="3" bgcolor="<%=m_bg1%>">
        <input type="button" value="关闭窗口" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()">
    </td>
  </tr>
 </table>
</body>
</html>
