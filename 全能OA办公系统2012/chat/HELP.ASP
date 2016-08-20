<!--#INCLUDE FILE="inc_config.asp"-->
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
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="3">
  <tr align="center" bgcolor="<%=m_bg1%>"> 
    <td style="color:<%=m_text1%>"><%=r_title%> 使 用 帮 助</td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td><br>
      　欢迎您来到『<%=r_title%>』，本聊天室有等级、用户资料修改、查看用户资料、给其它人留言、点歌等功能。<br>
      <br>
      　等级的划分及相应的权利：<br>
      <br>
      <table width="85%" border="1" bordercolor="#EFEFEF" cellspacing="0" align="center" bordercolordark="#FFFFFF">
        <tr> 
          <td bgcolor="<%=m_bg3%>" width="30">级别</td>
          <td bgcolor="<%=m_bg3%>">积分</td>
          <td bgcolor="<%=m_bg3%>">称号</td>
        </tr>
        <tr> 
          <td>1级</td>
          <td>1-300</td>
          <td>私聊，选择颜色，选择表情，选择趣语，自动清除<br>约会留言，查找聊友,修改信息，保存积分</td>
        </tr>
        <tr> 
          <td>2级</td>
          <td>300-600</td>
          <td>在一级的基础上增加点播MIDI和发送图案的功能</td>
        </tr>
        <tr> 
          <td>3级</td>
          <td>600-1200</td>
          <td>在二级的基础上增加发送贴图的功能，增加看用户级别功能</td>
        </tr>
        <tr> 
          <td>4级</td>
          <td>1200-2400</td>
          <td>在三级的基础上增加动作列表的功能</td>
        </tr>
        <tr> 
          <td>5级</td>
          <td>2400-3600</td>
          <td>在四级的基础上增加字效的功能</td>
        </tr>
        <tr> 
          <td>6级</td>
          <td>3600-6000</td>
          <td>在五级的基础上增加看见用户IP的权限</td>
        </tr>
        <tr> 
          <td>7级</td>
          <td>6000-10000</td>
          <td>在六级的基础上增加踢人，HTML权限，<br>快速HTML框，修改标题</td>
        </tr>
        <tr> 
          <td>8级</td>
          <td>10000以上</td>
          <td>在七级的基础上增加用户管理，清信息区，<br>清、锁标题，留言管理(需申请)</td>
        </tr>
        <tr> 
          <td>9级</td>
          <td>无</td>
          <td>最高管理员(需申请)</td>
        </tr>
      </table>
      <br>
      　等级功能： 聊天室按时间加分，每分钟加1分。 <br>
      <br>
      　注意：离开聊天室时一定要按[离开]，否则积分不会加。<br>
	  <br>
      　如有任何问题请联系<a href="mailto:<%=r_masteremail%>">管理员</a></td>
  </tr>
  <tr align="right" bgcolor="<%=m_bg1%>"> 
    <td><input type="button" value="关 闭 窗 口" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()"></td>
  </tr>
</table>
</body>
</html>
