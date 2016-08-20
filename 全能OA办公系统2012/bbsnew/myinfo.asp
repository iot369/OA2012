<%picnum=20%>
<!--#include file="up.asp"--><!--#include file="fun.asp"--><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<br><br><%sty="<P style='MARGIN: 8px'>"
t1="<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% background=pic/"&sp&"3.gif height=25 bgcolor="&c1&">&nbsp;<img border=0 src=pic/fl.gif> <font color=#FFFFFF><b>"
t2="</b></font></td></tr>"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
set canl=myconn.execute("select*from user where name='"&lgname&"' and password='"&lgpwd&"'")
if canl.eof or canl.bof then
%>
<%=t1%>错 误 信 息<%=t2&d1&sty%>·你还没有登陆或者你登陆的用户名或密码错误！·<%=d2%><%
response.end
end if
%>
<form method="POST" action="chinfo.asp" name="form">
<%=t1%>我的个人资料<%=t2%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-width:1; border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
<tr>
<td colspan="2" height="22" background="pic/4.gif">&nbsp;<b><font color="<%=c1%>">●必填资料：</font></b></td>
</tr>
<tr>
<td>
<%=sty%><b>用户名</b>：<br>注册用户名不能超过20个字符（10个汉字）</td>
<td width="50%">&nbsp;<b><%=kbbs(lgname)%></b></td>
</tr>
<tr>
<td><%=sty%>
<b>密码(最多16位)</b>：<br>请不要使用任何类似 '*'、' ' 或 HTML 字符</td>
<td>&nbsp;<input type="password" name="newpwd" size="30" maxlength="20" value="<%=canl("password")%>"></td>
</tr>
<tr>
<td><%=sty%><b>重复密码(最多16位)</b>：<br>请再输一遍确认</td>
<td>&nbsp;<input type="password" name="repwd" size="30" maxlength="20" value="<%=canl("password")%>"></td>
</tr>
<tr>
<td><%=sty%><b>Email地址</b>：<br>请输入有效的邮件地址，这将使您能用到论坛中的所有功能</td>
<td>&nbsp;<input type="text" name="email" size="30" maxlength="30" value="<%=canl("email")%>"></td>
</tr>
<tr>
<td><%=sty%><b>备用密码：</b><br>请牢记！忘记密码时可以用这个来充当密码！</td>
<td>&nbsp;<input type="text" name="anhao" size="30" maxlength="30" value="<%=canl("anhao")%>"></td>
</tr>
<tr>
<td colspan="2" height="22" background="pic/4.gif"><b>
    &nbsp;<font color="<%=c1%>">●选填资料：</font></b></td>
</tr>
<tr>
<td><%=sty%><b>性别：</b></td>
<td>&nbsp;<select size="1" name="sex" style="font-size: 9pt; border: 1px solid <%=c1%>; background-color: #FFFFEC">
        <%if canl("sex")="1" then%><option selected value="1">男</option>
        <option value="2">女</option><%else%><option selected value="2">女</option>
        <option value="1">男</option><%end if%></select></td>
</tr>
<tr>
<td><%=sty%><b>生日：</b>（请按照1900-1-1格式填写）</td>
<td>&nbsp;<input type="text" name="burn" size="21" maxlength="10" value="<%=canl("burn")%>"></td>
</tr>
<tr>
<td><%=sty%><b>主页：</b><br>填写你的个人主页，让大家见识见识！</td>
<td>&nbsp;<input type="text" name="home" size="30" maxlength="255" value="<%=canl("home")%>"></td>
</tr>
<tr>
<td><%=sty%><B>OICQ号码</B>：<BR>填写您的QQ地址，方便与他人的联系</td>
<td>&nbsp;<input type="text" name="qq" size="16" maxlength="10" value="<%=canl("qq")%>"></td>
</tr>
<tr>
<td valign="top"><%=sty%><b>我的头像：</b><br>使用论坛自带的图像</td>
<td>
<p style="margin-top: 3; margin-bottom: 3">&nbsp;<select name=headpic size=1 onChange="showimage()" style="font-size: 9pt">
<%for i=1 to picnum%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select><img src="headpic/1.gif" name="tus"><script>function showimage(){document.images.tus.src="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";document.form.mypic.value="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";document.form.ch.value="40";document.form.ku.value="40";}</script> <br>
</td>
</tr>
<tr>
<td><%=sty%><B>自定义头像</B>：<br>如果图像位置中有连接图片将以自定义的为主</td>
<td>&nbsp;<input name="mypic" size=38 maxlength=100 value="<%=canl("toupic")%>"> 完整Url地址<br>&nbsp;图像宽度：<input type="text" name="ku" size="6" value="<%=canl("ku")%>"> 高度： <input type="text" name="ch" size="6" value="<%=canl("ch")%>">（都不能大于120）</td>
</tr>
<tr>
<td valign="top" height="80"><%=sty%><B>个性签名</B>：<BR>最多255个字符<BR>文字将出现在您发表的文章的结尾处。体现您的个性。</td>
<td>&nbsp;<TEXTAREA name=gxqm rows=5 wrap=PHYSICAL cols=60><%=canl("gxqm")%></TEXTAREA></td>
</tr>
<tr>
<td colspan="2" height="30" align="center" background="pic/<%=sp%>3.gif">
<input type=submit value=我修改了，现在提交！ name=Submit>&nbsp;&nbsp; <input type=reset value=不行，还是重写吧！ name=Submit2></td>
</tr>
</table></center>
</div></form>
<br><!--#include file="down.asp"-->