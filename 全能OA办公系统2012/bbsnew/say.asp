<!--#include file="conn.asp"--><%
name=Request.Cookies(cn)("lgname")
pwd=Request.Cookies(cn)("lgpwd")
if name="" then
Response.Redirect"login.asp"
response.end
end if
application(name)=1
myconn.close
set myconn=nothing%>
<!--#include file="up.asp"-->
<%
function ubbs(str)
dim re
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
re.Pattern="(\[showtoname=(.[^\[]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")	
re.Pattern="(\[showtoreply\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")	
re.Pattern="(\[showtograde=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")
re.Pattern="(\[smoney=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")
re.Pattern="(\[smeili=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")
re.Pattern="(\[sjingyan=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]���� �������ݲ������� ����[/color][enter]")
	
str = replace(str, ">", "&gt;")
str = replace(str, "<", "&lt;")
	
set re=Nothing
ubbs=str
end function
%>
<SCRIPT>
function emoticon(theSmilie){
document.kbbs.body.value +=theSmilie + '';
document.kbbs.body.focus();
}
</SCRIPT><SCRIPT src="ybbcode.js"></SCRIPT>
<SCRIPT>var i=0;
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('�������ڷ����������ĵȴ���');return false;}this.document.kbbs.submit();}}
</SCRIPT><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%bdlogin(2)
%>
<%
session("upsize")=upsize
session("upnum")=upnum
pagenum=request.querystring("pagenum")
quoteid=request.querystring("quoteid")
re=request.querystring("re")
if re="no" then
action="save.asp?bd="&bd&"&re=no&pagenum="&pagenum&""
zhuti=""%>
<SCRIPT>function showvote(){
if (document.kbbs.voteyn.checked == true) {
vote.style.display = "";
}else{
vote.style.display = "none";
}
}
</SCRIPT>
<%
elseif re="yes" then
action="save.asp?bd="&bd&"&id="&id&"&re=yes&pagenum="&pagenum&""
zhuti="�ظ�ID:"&id&""
set lock=myconn.execute("select type from min where id="&id&"")
if lock("type")<>4 or canre="yes" then
else
response.end
end if
set lock=nothing
end if
%>
<form method=POST name=kbbs action=<%=action%>><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b> �� �� �� ��</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
    <tr>
  <td width="24%" height="30">	<b>&nbsp;�û�����</b></td>
  <td width="76%" height="30"> &nbsp;<input type=text name=name size=20 value="<%=name%>">
  <font color="<%=c1%>">*</font> <a href="zhuce.asp">û��ע�᣿</a></td>    </tr>
    <tr>
      <td width="24%" height="30">&nbsp;<b>��&nbsp; �룺</b></td>
      <td width="76%">&nbsp;<input type=password name=password size=20 value="<%=pwd%>">
      <font color="<%=c1%>">*</font> <a href="getpwd.asp">�������룿</a></td>
    </tr>
    <tr>
      <td width="24%" height="30">&nbsp;<b>�������⣺</b></td>
      <td width="76%">&nbsp;<input type=text name=zhuti size=80 maxlength=59 value="<%=zhuti%>">
  <font color="<%=c1%>">*</font></td>
    </tr>
    <tr>
      <td width="24%">
      <p style="line-height: 150%; margin-left: 5; margin-top: 5"><b>��ı��飺</b> <br>����������ǰ�档</td>
      <td width="76%">
  <p style="margin: 4"><input type=radio value=face1 name=face checked> 
  <img border=0 src=face/face1.gif width="16" height="16"> <input type=radio value=face2 name=face> 
  <img border=0 src=face/face2.gif width="16" height="16"> <input type=radio value=face3 name=face> 
  <img border=0 src=face/face3.gif width="16" height="16"> <input type=radio value=face4 name=face> 
  <img border=0 src=face/face4.gif width="16" height="16"> <input type=radio value=face5 name=face> 
  <img border=0 src=face/face5.gif width="16" height="16"> <input type=radio value=face6 name=face> 
  <img border=0 src=face/face6.gif width="16" height="16"> <input type=radio value=face7 name=face> 
  <img border=0 src=face/face7.gif width="16" height="16"> <input type=radio value=face8 name=face> 
  <img border=0 src=face/face8.gif width="16" height="16"> <input type=radio value=face9 name=face> 
  <img border=0 src=face/face9.gif width="16" height="16"><br>
    <input type=radio value=face10 name=face> 
  <img border=0 src=face/face10.gif width="16" height="16"> <input type=radio value=face11 name=face> 
  <img border=0 src=face/face11.gif width="16" height="16"> <input type=radio value=face12 name=face> 
  <img border=0 src=face/face12.gif width="16" height="16"> <input type=radio value=face13 name=face> 
  <img border=0 src=face/face13.gif width="16" height="16"> <input type=radio value=face14 name=face> 
  <img border=0 src=face/face14.gif width="16" height="16"> <input type=radio value=face15 name=face> 
  <img border=0 src=face/face15.gif width="16" height="16"> <input type=radio value=face16 name=face> 
  <img border=0 src=face/face16.gif width="16" height="16"> <input type=radio value=face17 name=face> 
  <img border=0 src=face/face17.gif width="16" height="16"> <input type=radio value=face18 name=face> 
  <img border=0 src=face/face18.gif width="16" height="16"></td>
    </tr>
    <%if re="no" then%><tr>
      <td width="24%" height="28">
      <p style="margin: 5"><b>�������ͣ�</b></td>
      <td width="76%">
<input type="radio" name="gonggao" value="0" checked>��ͨ����<a title="�����ڹ���������ʾ���⣬ʹ�ø�����������Ҫ�۳�����Ǯ1000������200������200��"> <input type="radio" name="gonggao" value="4"></a>�������� <%if admin="yes" then%><input type="radio" name="gonggao" value="1">ϵͳ����<%end if%></td>
    </tr><%end if%>
<tr>
      <td width="24%" valign="top">
<p style="line-height: 150%; margin-left: 5; margin-top: 5">
<b>�������ݣ�</b><br>
      ��HTML��ǩ�� ������<br>
��UBB��ǩ�� ����<br>
����ͼ��ǩ�� ����<br>
��Flash��ǩ��������<br>
�������ַ�ת����������<br>
���ϴ�ͼƬ������<br>
�����15KB<br><b>�������ݣ�</b><br>
��<a href="javascript:grade()" title='��ʽ��[showtograde=1]����[/s]			��ʾֻ�еȼ��ڡ�1�����ϲ�����������ݡ�			��1���͡����ݡ������Զ���				ʹ�ø���������'>�ȼ��ɼ�����</a><br>��<a href="javascript:reply()" title='��ʽ��[showtoreply]����[/s]			��ʾֻ�лظ��߲�����������ݡ���		�����ݡ������Զ���				ʹ�ø���������'>�ظ��ɼ�����</a><br>��<a href="javascript:name()" title='��ʽ��[showtoname=zym]����[/s]			��ʾֻ�С�zym��������������ݡ�			��zym���͡����ݡ������Զ���				ʹ�ø���������'>ָ����������</a><br>��<a href="javascript:smoney()" title='��ʽ��[smoney=1000]����[/s]			��ʾֻ�н�Ǯ��С�ڡ�1000��������������ݡ���1000���͡����ݡ������Զ���ʹ�ø���������'>��Ǯ</a>��<a href="javascript:smeili()" title='��ʽ��[smeili=1]����[/s]			��ʾֻ�������ڡ�1000�����ϲ�����������ݡ�		��1000���͡����ݡ������Զ���ʹ�ø���������'>����</a>��<a href="javascript:sjingyan()" title='��ʽ��[sjingyan=1000]����[/s]			��ʾֻ�о��鲻С�ڡ�1000��������������ݡ�		��1000���͡����ݡ������Զ���ʹ�ø���������'>����</a>��
<p style="line-height: 150%; margin-left: 5; margin-top: 5">
<%if re="no" then%><input type="checkbox" onclick=showvote() name="voteyn" value="1"> ��ʾͶƱѡ��<br><%end if%><br>
</td>      <td width="76%" valign="top">
      <p style="margin-left: 4; margin-top: 4">
        <p>
<IFRAME name=ad src="upload.asp" frameBorder=0 
            width="100%" scrolling=no height=25></IFRAME><br>&nbsp;���ܰ�ť��<IMG onclick=fly() alt=������ src="pic/fly.gif" border=0> 
        <IMG onclick=move() alt=�ƶ��� 
      src="pic/move.gif" border=0> 
        <IMG 
      onclick=light() alt=������ src="pic/glow.gif" border=0> 
        <IMG onclick=ying() alt=��Ӱ�� 
      src="pic/shadow.gif" border=0> 
        <IMG 
      onclick=image() alt=ͼƬ src="pic/image.gif" border=0> 
        <IMG onclick=Cswf() alt=FlashͼƬ 
      src="pic/swf.gif" border=0> 
        <IMG onclick=Cdir() alt=Shockwave�ļ� src="pic/Shockwave.gif" border=0> 
        <IMG onclick=Crm() alt=realplay��Ƶ�ļ� 
      src="pic/rm.gif" border=0> 
        <IMG onclick=Cwmv() alt="Media Player��Ƶ�ļ�" src="pic/mp.gif" border=0>
        <img onclick=center() alt="����" border="0" src="pic/center.gif">
        <img onclick=javascript:emoticon('[url=������]�����ı�[/url]') alt="������" border="0" src="pic/url1.gif"><br><br>&nbsp;���ִ�С��<select onchange=ybbsize(this.options[this.selectedIndex].value) name=a style="font-size: 9pt"><OPTION value=1>1</OPTION><OPTION value=2>2</OPTION><OPTION value=3>3</OPTION><OPTION value=4>4</OPTION></SELECT> <span lang=zh-cn>��ɫ��<select onchange=COLOR(this.options[this.selectedIndex].value) name="111" style="font-size: 9pt"><option style='COLOR:000000;BACKGROUND-COLOR:000000' value=000000>��ɫ</option><option style='COLOR:FFFFFF;BACKGROUND-COLOR:FFFFFF' value=FFFFFF>��ɫ</option><option style='COLOR:008000;BACKGROUND-COLOR:008000' value=008000>��ɫ</option><option style='COLOR:800000;BACKGROUND-COLOR:800000' value=800000>��ɫ</option><option style='COLOR:808000;BACKGROUND-COLOR:808000' value=808000>���ɫ</option><option style='COLOR:000080;BACKGROUND-COLOR:000080' value=000080>����ɫ</option><option style='COLOR:800080;BACKGROUND-COLOR:800080' value=800080>��ɫ</option><option style='COLOR:808080;BACKGROUND-COLOR:808080' value=808080>��ɫ</option><option style='COLOR:FFFF00;BACKGROUND-COLOR:FFFF00' value=FFFF00>��ɫ</option><option style='COLOR:00FF00;BACKGROUND-COLOR:00FF00' value=00FF00>ǳ��ɫ</option><option style='COLOR:00FFFF;BACKGROUND-COLOR:00FFFF' value=00FFFF>ǳ��ɫ</option><option style='COLOR:FF00FF;BACKGROUND-COLOR:FF00FF' value=FF00FF>�ۺ�ɫ</option><option style='COLOR:C0C0C0;BACKGROUND-COLOR:C0C0C0' value=C0C0C0>����ɫ</option><option style='COLOR:FF0000;BACKGROUND-COLOR:FF0000' value=FF0000>��ɫ</option><option style='COLOR:0000FF;BACKGROUND-COLOR:0000FF' value=0000FF>��ɫ</option><option style='COLOR:008080;BACKGROUND-COLOR:008080' value=008080>����ɫ</option></select><br><br>
  &nbsp;
<%if quoteid<>"" then
set quote=myconn.execute("select top 1 name,riqi,body,bd,gonggao from min where id="&quoteid&"")
qqbody=quote("body")
%>
<textarea onkeydown=presskey(); rows=8 name=body cols=79 Title='�� Ctrl+Enter ֱ�ӷ���'>
[quote][color=<%=c1%>]���������� <%=quote("name")%> �� <%=quote("riqi")%> ��������ݣ�[/color]
<%=ubbs(qqbody)%>
[/quote]
</textarea>
<%
set quote=nothing
else%>
<textarea onkeydown=presskey(); rows=8 name=body cols=79 Title='�� Ctrl+Enter ֱ�ӷ���'></textarea><%end if%><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="30">
          <tr>
            <td>
            &nbsp;���ͼ�������飺</td>
            <td>
            <A href=javascript:emoticon('{f1)')><IMG alt=Ц�� src=face/xl.gif border=0></A> <A href=javascript:emoticon('{f2)')><IMG alt=����Ц�� src=face/kk.gif border=0></A> <A href=javascript:emoticon('{f3)')><IMG alt=���ȵ�Ц�� src=face/jy.gif border=0></A> <A href=javascript:emoticon('{f4)')><IMG alt=����Ц�� src=face/ts.gif border=0></A> <A href=javascript:emoticon('{f5)')><IMG alt=գ��΢Ц src=face/zy.gif border=0></A> <A href=javascript:emoticon('{f6)')><IMG alt=�ѹ����� src=face/ng.gif border=0></A> <A href=javascript:emoticon('{f7)')><IMG alt=�����Ц�� src=face/kh.gif border=0></A> <A href=javascript:emoticon('{f8)')><IMG alt=ʧ������ src=face/sw.gif border=0></A> <a href=javascript:emoticon('{f9)')><IMG alt=���ε�Ц�� src=face/gg.gif border=0></a>

            </td>
          </tr>
        </table>
</td>
    </tr>  
    <tr id=vote style="DISPLAY: none">
      <td valign="top">
<p style="margin: 5; ; line-height:150%"><b>ͶƱ��Ŀ</b><br>��������Ŀ�ûس�����<br>
<input type="radio" name="votetype" value="1" checked>��ѡ&nbsp;
<input type="radio" name="votetype" value="2">��ѡ</p>
<p style="margin: 5; ; line-height:150%"><b>����ʱ�䣺</b><select size="1" name="outtime" style="font-size: 9pt">
<option value="1">һ��</option>
<option value="3">����</option>
<option value="7">һ��</option>
<option value="15">�����</option>
<option value="31">һ����</option>
<option value="93">������</option>
<option value="365">һ��</option>
<option value="10000" selected>������</option>
</select><br></p>
</td>      <td valign="top">
      <p style="margin: 5"> &nbsp;<textarea onkeydown=presskey(); rows=6 name=vote cols=79 Title='�� Ctrl+Enter ֱ�ӷ���'></textarea></td>
    </tr>
    <tr>
      <td width="100%" height="40" colspan="2" align="center" bgcolor="<%=c2%>">
      &nbsp;<input type=submit value=OK_���ˣ��������� name=B1> <input type=reset value=NO_���У���Ҫ��д name=B2> [��  Ctrl+Enter ֱ�ӷ���]</td>      
    </tr>
  </table>
  </center>
</div></form><!--#include file="down.asp"-->