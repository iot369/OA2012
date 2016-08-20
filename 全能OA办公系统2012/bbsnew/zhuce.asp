<%picnum=20%>
<%sty1="<p style='line-height: 150%; margin-left: 4; margin-top: 4'>"%>
<!--#include file="up.asp"--><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<br><br><%
action=request.querystring("action")
if action<>"agree" then
noyes="服务条款和声明"
mes="<b>继续注册前请先阅读论坛协议</b><br>欢迎您加入本站点参加交流和讨论，本站点为公共论坛，为维护网上公共秩序和社会稳定，请您自觉遵守以下条款：<br><br>一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利用本站制作、复制和传播下列信息： <br>（一）煽动抗拒、破坏宪法和法律、行政法规实施的；<br>（二）煽动颠覆国家政权，推翻社会主义制度的；<br>（三）煽动分裂国家、破坏国家统一的；<br>（四）煽动民族仇恨、民族歧视，破坏民族团结的；<br>（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的；<br>（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；<br>（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br>（八）损害国家机关信誉的；<br>（九）其他违反宪法和法律行政法规的；<br>（十）进行商业广告行为的。<br>二、互相尊重，对自己的言论和行为负责。<form method=POST action='?action=agree'><center><input type=submit value=' 我 同 意 ' name=B1></center></form>"
%><!--#include file="mes.asp"--><%else%>
<form method="POST" action="reg.asp" name="form">
<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>新 用 户 注 册</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="94%">
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-width:1; border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
<tr>
<td colspan="2" height="24" background="pic/4.gif">&nbsp;<b><font color="<%=c1%>">●必填资料：</font></b></td>
</tr>
<tr>
<td>
<%=sty1%><b>用户名</b>：<br>注册用户名不能超过16个字符（8个汉字）</td>
<td width="50%">&nbsp;<input type=text name=name size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%>
<b>密码(最多16位)</b>：<br>请不要使用任何类似 '*'、' ' 或 HTML 字符</td>
<td>&nbsp;<input type=password name=password size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%><b>重复密码(最多16位)</b>：<br>请再输一遍确认</td>
<td>&nbsp;<input type=password name=repassword size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%><b>Email地址</b>：<br>请输入有效的邮件地址，这将使您能用到论坛中的所有功能</td>
<td>&nbsp;<input type=text name=email size=30 maxlength=30></td>
</tr>
<tr>
<td><%=sty1%><b>备用密码：</b><br>请牢记！忘记密码时可以用这个来充当密码！</td>
<td>&nbsp;<input type=text name=anhao size=30 maxlength=30></td>
</tr>
<tr>
<td colspan="2" height="24" background="pic/4.gif"><b>
    &nbsp;<font color="<%=c1%>">●选填资料：</font></b></td>
</tr>
<tr>
<td><%=sty1%><b>性别：</b></td>
<td>&nbsp;<select size=1 name=sex style='font-size: 9pt; border: 1px solid ; background-color: #FFFFEC'><option selected value=1>男</option><option value=2>女</option></select></td>
</tr>
<tr>
<td><%=sty1%><b>生日：</b>（请按照1900-1-1格式填写）</td>
<td>&nbsp;<input type=text name=burn size=21 maxlength=10 value=不告诉你></td>
</tr>
<tr>
<td><%=sty1%><b>主页：</b><br>填写你的个人主页，让大家见识见识！</td>
<td>&nbsp;<input type=text name=home size=30 maxlength=255></td>
</tr>
<tr>
<td><%=sty1%><B>OICQ号码</B>：<BR>填写您的QQ地址，方便与他人的联系</td>
<td>&nbsp;<input type=text name=qq size=16 maxlength=10></td>
</tr>
<tr>
<td valign="top"><%=sty1%><b>我的头像：</b><br>使用论坛自带的图像</td>
<td>
<p style="margin-top: 3; margin-bottom: 3">&nbsp;<select name=headpic size=1 onChange="showimage()" style="font-size: 9pt">
<%for i=1 to picnum%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select><img src="headpic/1.gif" name="tus"><script>function showimage(){document.images.tus.src="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";}</script> <br>
</td>
</tr>
<tr>
<td><%=sty1%><B>自定义头像</B>：<br>如果图像位置中有连接图片将以自定义的为主</td>
<td>&nbsp;<input name=mypic size=38 maxlength=100> 完整Url地址<br>&nbsp;图像宽度：<input type=text name=ku size=6> 高度： <input type=text name=ch size=6>（都不能大于120）</td>
</tr>
<tr>
<td valign="top" height="80"><%=sty1%><B>个性签名</B>：<BR>最多255个字符<BR>文字将出现在您发表的文章的结尾处。体现您的个性。</td>
<td>&nbsp;<TEXTAREA name=gxqm rows=5 wrap=PHYSICAL cols=60></TEXTAREA></td>
</tr>
<tr>
<td colspan="2" height="30" align="center" background="pic/<%=sp%>3.gif"><input type=submit value=我填好了，现在注册！ name=Submit>&nbsp;&nbsp; <input type=reset value=不行，还是重写吧！ name=Submit2></td>
</tr>
</table></center>
</div></form><%end if%><br><!--#include file="down.asp"-->