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
<%=t1%>�� �� �� Ϣ<%=t2&d1&sty%>���㻹û�е�½�������½���û�����������󣡡�<%=d2%><%
response.end
end if
%>
<form method="POST" action="chinfo.asp" name="form">
<%=t1%>�ҵĸ�������<%=t2%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-width:1; border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
<tr>
<td colspan="2" height="22" background="pic/4.gif">&nbsp;<b><font color="<%=c1%>">��������ϣ�</font></b></td>
</tr>
<tr>
<td>
<%=sty%><b>�û���</b>��<br>ע���û������ܳ���20���ַ���10�����֣�</td>
<td width="50%">&nbsp;<b><%=kbbs(lgname)%></b></td>
</tr>
<tr>
<td><%=sty%>
<b>����(���16λ)</b>��<br>�벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ�</td>
<td>&nbsp;<input type="password" name="newpwd" size="30" maxlength="20" value="<%=canl("password")%>"></td>
</tr>
<tr>
<td><%=sty%><b>�ظ�����(���16λ)</b>��<br>������һ��ȷ��</td>
<td>&nbsp;<input type="password" name="repwd" size="30" maxlength="20" value="<%=canl("password")%>"></td>
</tr>
<tr>
<td><%=sty%><b>Email��ַ</b>��<br>��������Ч���ʼ���ַ���⽫ʹ�����õ���̳�е����й���</td>
<td>&nbsp;<input type="text" name="email" size="30" maxlength="30" value="<%=canl("email")%>"></td>
</tr>
<tr>
<td><%=sty%><b>�������룺</b><br>���μǣ���������ʱ������������䵱���룡</td>
<td>&nbsp;<input type="text" name="anhao" size="30" maxlength="30" value="<%=canl("anhao")%>"></td>
</tr>
<tr>
<td colspan="2" height="22" background="pic/4.gif"><b>
    &nbsp;<font color="<%=c1%>">��ѡ�����ϣ�</font></b></td>
</tr>
<tr>
<td><%=sty%><b>�Ա�</b></td>
<td>&nbsp;<select size="1" name="sex" style="font-size: 9pt; border: 1px solid <%=c1%>; background-color: #FFFFEC">
        <%if canl("sex")="1" then%><option selected value="1">��</option>
        <option value="2">Ů</option><%else%><option selected value="2">Ů</option>
        <option value="1">��</option><%end if%></select></td>
</tr>
<tr>
<td><%=sty%><b>���գ�</b>���밴��1900-1-1��ʽ��д��</td>
<td>&nbsp;<input type="text" name="burn" size="21" maxlength="10" value="<%=canl("burn")%>"></td>
</tr>
<tr>
<td><%=sty%><b>��ҳ��</b><br>��д��ĸ�����ҳ���ô�Ҽ�ʶ��ʶ��</td>
<td>&nbsp;<input type="text" name="home" size="30" maxlength="255" value="<%=canl("home")%>"></td>
</tr>
<tr>
<td><%=sty%><B>OICQ����</B>��<BR>��д����QQ��ַ�����������˵���ϵ</td>
<td>&nbsp;<input type="text" name="qq" size="16" maxlength="10" value="<%=canl("qq")%>"></td>
</tr>
<tr>
<td valign="top"><%=sty%><b>�ҵ�ͷ��</b><br>ʹ����̳�Դ���ͼ��</td>
<td>
<p style="margin-top: 3; margin-bottom: 3">&nbsp;<select name=headpic size=1 onChange="showimage()" style="font-size: 9pt">
<%for i=1 to picnum%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select><img src="headpic/1.gif" name="tus"><script>function showimage(){document.images.tus.src="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";document.form.mypic.value="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";document.form.ch.value="40";document.form.ku.value="40";}</script> <br>
</td>
</tr>
<tr>
<td><%=sty%><B>�Զ���ͷ��</B>��<br>���ͼ��λ����������ͼƬ�����Զ����Ϊ��</td>
<td>&nbsp;<input name="mypic" size=38 maxlength=100 value="<%=canl("toupic")%>"> ����Url��ַ<br>&nbsp;ͼ���ȣ�<input type="text" name="ku" size="6" value="<%=canl("ku")%>"> �߶ȣ� <input type="text" name="ch" size="6" value="<%=canl("ch")%>">�������ܴ���120��</td>
</tr>
<tr>
<td valign="top" height="80"><%=sty%><B>����ǩ��</B>��<BR>���255���ַ�<BR>���ֽ�����������������µĽ�β�����������ĸ��ԡ�</td>
<td>&nbsp;<TEXTAREA name=gxqm rows=5 wrap=PHYSICAL cols=60><%=canl("gxqm")%></TEXTAREA></td>
</tr>
<tr>
<td colspan="2" height="30" align="center" background="pic/<%=sp%>3.gif">
<input type=submit value=���޸��ˣ������ύ�� name=Submit>&nbsp;&nbsp; <input type=reset value=���У�������д�ɣ� name=Submit2></td>
</tr>
</table></center>
</div></form>
<br><!--#include file="down.asp"-->