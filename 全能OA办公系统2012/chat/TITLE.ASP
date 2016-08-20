<!--#INCLUDE FILE="inc_config.asp"-->
<html>
<head>
<title>更改标题</title>
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

<body bgcolor="#cccccc" text="<%=m_text2%>">
<br>
<% if session("userlevel")<7 then
      response.end
   end if
      if session("userlevel")<8 and application("titlock")="lock" then %>
         <br><br><center><strong>标题被锁定!<br><br></strong>
         <input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()"></center>
         <% response.end 
      end if
      if request("do")="lockit" and session("userlevel")>7 then
         application("titlock")="lock"
         if application("title")<>"" then
            application("title")=application("title")+"[已锁定]"
         end if %>
      <script language="javascript">
      alert("标题已锁定!");
      </script>
      <% end if
         if request("do")="unlock" and session("userlevel")>7 then
         application("titlock")=""
'application("title")=right(application("title"),12)
         if right(application("title"),5)="[已锁定]" then
           application("title")=left(application("title"),len(application("title"))-5)
         end if %>
      <script language="javascript">
      alert("标题已解锁!");
      </script>
      <% end if
         if request("do")="empty" and session("userlevel")>7 then
         application("title")="" %>
      <script language="javascript">
      alert("标题已清空!");
      </script>
      <% end if

      if trim(request("bt"))<>"" then
         if session("userlevel")<9 then
            bt=Server.HTMLEncode(trim(request("bt")))
         else
            bt=trim(request("bt"))
         end if
         application("title")="<font  face='"&request("font")&"' color='"&request("color")&"'>" +bt+"</font>" %>
      <script language="javascript">
      alert("标题已更改!");
      </script>
      <% end if %>
<table border="1" align="center" cellpadding="3" cellspacing="0" width="280">
  <form method="POST" action="title.asp">
    <tr bgcolor="<%=m_bg1%>"> 
      <td colspan="2" align="center" style="color:<%=m_text1%>">更改标题 </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 标题：<input type="text" name="bt" size="27" maxlength=200 style="font-size:9pt">
      </td>
    </tr>
    <tr align="center">
      <td colspan="2">颜色: 
        <select name="color" style="font-size:9pt">
          <option style="color: #000000" value="000000">绝对黑色</option>
          <option style="color: #000088" value="000088">深蓝忧郁</option>
          <option style="color: #0000ff" value="0000ff">草原之蓝</option>
          <option style="color: #008800" value="008800">橄榄树绿</option>
          <option style="color: #008888" value="008888">灰蓝种族</option>
          <option style="color: #0088ff" value="0088ff">海洋之蓝</option>
          <option style="color: #00a010" value="00a010">绿色回忆</option>
          <option style="color: #1100ff" value="1100ff">蓝色月光</option>
          <option style="color: #111111" value="111111">夜幕低垂</option>
          <option style="color: #333333" value="333333">灰色轨迹</option>
          <option style="color: #50b000" value="50b000">春草青青</option>
          <option style="color: #880000" value="880000">暗夜兴奋</option>
          <option style="color: #8800ff" value="8800ff">发亮蓝紫</option>
          <option style="color: #888800" value="888800">卡其制服</option>
          <option style="color: #888888" value="888888">伦敦灰雾</option>
          <option style="color: #8888ff" value="8888ff">兴份过蓝</option>
          <option style="color: #aa00cc" value="aa00cc">紫的拘谨</option>
          <option style="color: #aaaa00" value="aaaa00">流金岁月</option>
          <option style="color: #ccaa00" value="ccaa00">卡布其诺</option>
          <option style="color: #ff0000" value="ff0000">正宗喜红</option>
          <option style="color: #ff0088" value="ff0088">爱的暗示</option>
          <option style="color: #ff00ff" value="ff00ff">红的发紫</option>
          <option style="color: #ff8800" value="ff8800">黄金岁月</option>
          <option style="color: #ff0005" value="ff0005">红袍飘飘</option>
          <option style="color: #ff88ff" value="ff88ff">紫金绣帖</option>
          <option style="color: #ee0005" value="ee0005" selected>焚心似火</option>
          <option style="color: #ee01ff" value="ee01ff">红粉佳人</option>
          <option style="color: #3388aa" value="3388aa">我不知道</option>
        </select>
        字体: 
        <select name="font" style="font-size:9pt">
          <option value="宋体" selected>宋体</option>
          <option value="楷体">楷体</option>
          <option value="黑体">黑体</option>
          <option value="幼圆">幼圆</option>
          <option value="隶书">隶书</option>
        </select>
      </td>
    </tr>
    <tr align="center">
      <td colspan="2"><% if session("userlevel")>7 then %> 
        <input type="button" value="锁定标题" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'" onclick="location.href='title.asp?do=lockit';">
        <input type="button" value="解除锁定" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'" onclick="location.href='title.asp?do=unlock';">
        <input type="button" value="清空标题" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'" onclick="location.href='title.asp?do=empty';">
        <% end if %> </td>
    </tr>
    <tr align="center">
      <td colspan="2">
        <input type="submit" value="更 改" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'">
        <input type="reset" value="复 原" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'">
        <input type="button" value="关 闭" class="normal" onMouseOver="this.className='over'" onMouseDown="this.className='down'" onMouseOut="this.className='normal'" onClick="window.close()">
      </td>
    </tr>
  </form>
</table>
</body>
</html>