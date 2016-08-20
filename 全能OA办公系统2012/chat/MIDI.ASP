<!--#INCLUDE FILE="inc_config.asp"-->
<html>
<head>
<title>点歌台</title>
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

<body text="<%=m_text2%>" bgcolor="#cccccc">
<%if session("user")<>"" and session("userlevel")>1 then%>
<table width="280" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999" align="center">
  <form method=POST action="midisend.asp">
      <tr bgcolor="<%=m_bg1%>"> 
      <td colspan="2" align="center" style="color:<%=m_text1%>">点 歌 台</td>
      </tr>
      <tr>         
      <td align="center"> 歌　名： </td>
        <td> 
          <select name="midi">
            <option value="noname" selected>无名</option>
            <option value="lowheart">心太软</option>
            <option value="tian">甜蜜蜜</option>
            <option value="love2000">恋曲2000</option>
            <option value="crysea">泪海</option>
            <option value="huanzu">还珠</option>
            <option value="wind">风中奇缘</option>
            <option value="heart">动心</option>
            <option value="donotask">别问我是谁</option>
            <option value="think">想说爱你不容易</option>
            <option value="thank">谢谢你的爱</option>
            <option value="little">有一点点动心</option>
            <option value="dream">半梦半醒</option>
            <option value="hero">真心英雄</option>
            <option value="birthday">祝你生日快乐</option>
            <option value="foryou">特别的爱给特别的你</option>
            <option value="letme">让我欢喜让我忧</option>
            <option value="promise">诺言（李翊君）</option>
            <option value="nextlife">来生缘</option>
            <option value="choose">选择</option>
            <option value="yihan">遗憾</option>
            <option value="realy">真的有你</option>
            <option value="know">知不知</option>
          </select>
        </td>
      </tr>
      <tr> 
      <td align="center"> 点歌给： </td>
        <td> 
          <select name="name">
            <option value="大家" selected>大家</option>
            <% for i=1 to 100
               if application("user"&i)<>"" then %>
            <option value="<% =application("user"&i) %>"><% =application("user"&i) %></option>
            <% end if
            next %>
          </select>
        </td>
      </tr>
      <tr> 
      <td colspan="2" align="center"> 
        <input type="submit" value="发送音乐" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        <input type="button" value="关　　闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="window.close();">
        </td>
      </tr>
    </form>
   </table>
   <%end if%>
</body>
</html>