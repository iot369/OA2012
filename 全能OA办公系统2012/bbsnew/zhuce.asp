<%picnum=20%>
<%sty1="<p style='line-height: 150%; margin-left: 4; margin-top: 4'>"%>
<!--#include file="up.asp"--><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<br><br><%
action=request.querystring("action")
if action<>"agree" then
noyes="�������������"
mes="<b>����ע��ǰ�����Ķ���̳Э��</b><br>��ӭ�����뱾վ��μӽ��������ۣ���վ��Ϊ������̳��Ϊά�����Ϲ������������ȶ��������Ծ������������<br><br>һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ�� <br>��һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ�<br>������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<br>������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<br>���ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br>���壩�������������ʵ��ɢ��ҥ�ԣ������������ģ�<br>����������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<br>���ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<br>���ˣ��𺦹��һ��������ģ�<br>���ţ�����Υ���ܷ��ͷ�����������ģ�<br>��ʮ��������ҵ�����Ϊ�ġ�<br>�����������أ����Լ������ۺ���Ϊ����<form method=POST action='?action=agree'><center><input type=submit value=' �� ͬ �� ' name=B1></center></form>"
%><!--#include file="mes.asp"--><%else%>
<form method="POST" action="reg.asp" name="form">
<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>�� �� �� ע ��</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="94%">
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-width:1; border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
<tr>
<td colspan="2" height="24" background="pic/4.gif">&nbsp;<b><font color="<%=c1%>">��������ϣ�</font></b></td>
</tr>
<tr>
<td>
<%=sty1%><b>�û���</b>��<br>ע���û������ܳ���16���ַ���8�����֣�</td>
<td width="50%">&nbsp;<input type=text name=name size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%>
<b>����(���16λ)</b>��<br>�벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ�</td>
<td>&nbsp;<input type=password name=password size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%><b>�ظ�����(���16λ)</b>��<br>������һ��ȷ��</td>
<td>&nbsp;<input type=password name=repassword size=30 maxlength=20></td>
</tr>
<tr>
<td><%=sty1%><b>Email��ַ</b>��<br>��������Ч���ʼ���ַ���⽫ʹ�����õ���̳�е����й���</td>
<td>&nbsp;<input type=text name=email size=30 maxlength=30></td>
</tr>
<tr>
<td><%=sty1%><b>�������룺</b><br>���μǣ���������ʱ������������䵱���룡</td>
<td>&nbsp;<input type=text name=anhao size=30 maxlength=30></td>
</tr>
<tr>
<td colspan="2" height="24" background="pic/4.gif"><b>
    &nbsp;<font color="<%=c1%>">��ѡ�����ϣ�</font></b></td>
</tr>
<tr>
<td><%=sty1%><b>�Ա�</b></td>
<td>&nbsp;<select size=1 name=sex style='font-size: 9pt; border: 1px solid ; background-color: #FFFFEC'><option selected value=1>��</option><option value=2>Ů</option></select></td>
</tr>
<tr>
<td><%=sty1%><b>���գ�</b>���밴��1900-1-1��ʽ��д��</td>
<td>&nbsp;<input type=text name=burn size=21 maxlength=10 value=��������></td>
</tr>
<tr>
<td><%=sty1%><b>��ҳ��</b><br>��д��ĸ�����ҳ���ô�Ҽ�ʶ��ʶ��</td>
<td>&nbsp;<input type=text name=home size=30 maxlength=255></td>
</tr>
<tr>
<td><%=sty1%><B>OICQ����</B>��<BR>��д����QQ��ַ�����������˵���ϵ</td>
<td>&nbsp;<input type=text name=qq size=16 maxlength=10></td>
</tr>
<tr>
<td valign="top"><%=sty1%><b>�ҵ�ͷ��</b><br>ʹ����̳�Դ���ͼ��</td>
<td>
<p style="margin-top: 3; margin-bottom: 3">&nbsp;<select name=headpic size=1 onChange="showimage()" style="font-size: 9pt">
<%for i=1 to picnum%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select><img src="headpic/1.gif" name="tus"><script>function showimage(){document.images.tus.src="headpic/"+document.form.headpic.options[document.form.headpic.selectedIndex].value+".gif";}</script> <br>
</td>
</tr>
<tr>
<td><%=sty1%><B>�Զ���ͷ��</B>��<br>���ͼ��λ����������ͼƬ�����Զ����Ϊ��</td>
<td>&nbsp;<input name=mypic size=38 maxlength=100> ����Url��ַ<br>&nbsp;ͼ���ȣ�<input type=text name=ku size=6> �߶ȣ� <input type=text name=ch size=6>�������ܴ���120��</td>
</tr>
<tr>
<td valign="top" height="80"><%=sty1%><B>����ǩ��</B>��<BR>���255���ַ�<BR>���ֽ�����������������µĽ�β�����������ĸ��ԡ�</td>
<td>&nbsp;<TEXTAREA name=gxqm rows=5 wrap=PHYSICAL cols=60></TEXTAREA></td>
</tr>
<tr>
<td colspan="2" height="30" align="center" background="pic/<%=sp%>3.gif"><input type=submit value=������ˣ�����ע�ᣡ name=Submit>&nbsp;&nbsp; <input type=reset value=���У�������д�ɣ� name=Submit2></td>
</tr>
</table></center>
</div></form><%end if%><br><!--#include file="down.asp"-->