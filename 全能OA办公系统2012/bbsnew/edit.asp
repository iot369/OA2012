<!--#include file="up.asp"--><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%
function hell(str)
dim re
Set re=new RegExp
re.IgnoreCase=true
re.Global=True
re.Pattern=""&chr(10)&""&chr(10)&""&chr(10)&"(\[right\])(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])(\[\/right\])"
str=re.Replace(str,"")	
str = replace(str, ">", "&gt;")
str = replace(str, "<", "&lt;")
set re=Nothing
hell=str
end function
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% ><P style='MARGIN: 15px'>"
d2="</td></tr></table></center></div>"
id=request.querystring("id")
set edit=myconn.execute("select*from min where id="&id&"")
nnn=edit("name")
if lgname<>nnn then
abc="no"
end if
set edit=nothing
set che=myconn.execute("select name from user where name='"&lgname&"' and password='"&lgpwd&"'")
if che.eof then
abc="no"
end if
set che=nothing
set adok=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if not adok.eof then
abc="yes"
end if
set adok=nothing
set adok=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='"&bd&"'")
if not adok.eof then
abc="yes"
end if
set adok=nothing
if abc="no" then
%><br><br>
<%=t1%>�� �� �� Ϣ<%=t2&d1%>���㲻�Ǹ��������߻�ð���İ��񣬲��ܱ༭�����ӡ�<%=d2%>
<%
response.end
end if
set edit=nothing
%>
<%
id=request.querystring("id")
ed=request.querystring("ed")
reid=request.querystring("reid")
bd=request.querystring("bd")
set edit=myconn.execute("select*from min where id="&id&"")
%>
<%if ed=1 then
add=edit("zhuti")
else
add="re"
end if
zhuti=Replace(Request.Form("zhuti"),"'","''")
body=Replace(Request.Form("body"),"'","''")
body=""&body&chr(10)&chr(10)&chr(10)&"[right][color="&c1&"]�������ӱ� "&lgname&" �� "&now&" �༭����[/color][/right]"
face=Replace(Request.Form("face"),"'","''")
if zhuti="" or body="" or face="" then
%><br>
<SCRIPT>
function emoticon(theSmilie){
document.kbbs.body.value +=theSmilie + '';
document.kbbs.body.focus();
}
</SCRIPT><SCRIPT src="ybbcode.js"></SCRIPT>
<SCRIPT>var i=0;
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('�������ڷ����������ĵȴ���');return false;}this.document.kbbs.submit();}}
</SCRIPT><form method=POST name=kbbs>

<%=t1%>�� �� �� ��<%=t2%>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
        <tr>
      <td width="24%" height="30">&nbsp;<b>�������⣺</b></td>
      <td width="76%">&nbsp;<input type=text name=zhuti size=80 maxlength=59 value="<%=add%>">
  <font color="<%=c1%>">*</font></td>
    </tr>
    <tr>
      <td width="24%" height="50">
      <p style="line-height: 150%; margin-left: 5; margin-top: 5"><b>��ı��飺</b> <br>����������ǰ�档</td>
      <td width="76%"><input type=radio value=face1 name=face checked> 
  <img border=0 src=face/face1.gif width="16" height="16"> <input type=radio value=face2 name=face> 
  <img border=0 src=face/face2.gif width="16" height="16"> <input type=radio value=face3 name=face> 
  <img border=0 src=face/face3.gif width="16" height="16"> <input type=radio value=face4 name=face> 
  <img border=0 src=face/face4.gif width="16" height="16"> <input type=radio value=face5 name=face> 
  <img border=0 src=face/face5.gif width="16" height="16"> <input type=radio value=face6 name=face> 
  <img border=0 src=face/face6.gif width="16" height="16"> <input type=radio value=face7 name=face> 
  <img border=0 src=face/face7.gif width="16" height="16"> <input type=radio value=face8 name=face> 
  <img border=0 src=face/face8.gif width="16" height="16"> <input type=radio value=face9 name=face> 
  <img border=0 src=face/face9.gif width="16" height="16"><p style='margin-top: 2; margin-bottom: 2'>
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
<tr>
      <td width="24%" height="296" valign="top">
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
<br></td>      <td width="76%" valign="top">
      <p style="margin-left: 4; margin-top: 4">
        <p>
        <IFRAME name=ad 
            src="upload.asp" frameBorder=0 
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
        <IMG onclick=Cwmv() alt="Media Player��Ƶ�ļ�" src="pic/mp.gif" border=0><br><br>&nbsp;���ִ�С��<select onchange=ybbsize(this.options[this.selectedIndex].value) name=a style="font-size: 9pt"><OPTION value=1>1</OPTION><OPTION value=2>2</OPTION><OPTION value=3>3</OPTION><OPTION value=4>4</OPTION></SELECT> <span lang=zh-cn>��ɫ��<select onchange=COLOR(this.options[this.selectedIndex].value) name="111" style="font-size: 9pt"><option style='COLOR:000000;BACKGROUND-COLOR:000000' value=000000>��ɫ</option><option style='COLOR:FFFFFF;BACKGROUND-COLOR:FFFFFF' value=FFFFFF>��ɫ</option><option style='COLOR:008000;BACKGROUND-COLOR:008000' value=008000>��ɫ</option><option style='COLOR:800000;BACKGROUND-COLOR:800000' value=800000>��ɫ</option><option style='COLOR:808000;BACKGROUND-COLOR:808000' value=808000>���ɫ</option><option style='COLOR:000080;BACKGROUND-COLOR:000080' value=000080>����ɫ</option><option style='COLOR:800080;BACKGROUND-COLOR:800080' value=800080>��ɫ</option><option style='COLOR:808080;BACKGROUND-COLOR:808080' value=808080>��ɫ</option><option style='COLOR:FFFF00;BACKGROUND-COLOR:FFFF00' value=FFFF00>��ɫ</option><option style='COLOR:00FF00;BACKGROUND-COLOR:00FF00' value=00FF00>ǳ��ɫ</option><option style='COLOR:00FFFF;BACKGROUND-COLOR:00FFFF' value=00FFFF>ǳ��ɫ</option><option style='COLOR:FF00FF;BACKGROUND-COLOR:FF00FF' value=FF00FF>�ۺ�ɫ</option><option style='COLOR:C0C0C0;BACKGROUND-COLOR:C0C0C0' value=C0C0C0>����ɫ</option><option style='COLOR:FF0000;BACKGROUND-COLOR:FF0000' value=FF0000>��ɫ</option><option style='COLOR:0000FF;BACKGROUND-COLOR:0000FF' value=0000FF>��ɫ</option><option style='COLOR:008080;BACKGROUND-COLOR:008080' value=008080>����ɫ</option></select><br><br>
  &nbsp;<textarea onkeydown=presskey(); rows=10 name=body cols=79 Title='�� Ctrl+Enter ֱ�ӷ���'><%=hell(edit("body"))%></textarea><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="30">
          <tr>
            <td>
            &nbsp;���ͼ�������飺</td>
            <td>
            <A href=javascript:emoticon('{f1)')><IMG alt=Ц�� src=face/xl.gif border=0></A> <A href=javascript:emoticon('{f2)')><IMG alt=����Ц�� src=face/kk.gif border=0></A> <A href=javascript:emoticon('{f3)')><IMG alt=���ȵ�Ц�� src=face/jy.gif border=0></A> <A href=javascript:emoticon('{f4)')><IMG alt=����Ц�� src=face/ts.gif border=0></A> <A href=javascript:emoticon('{f5)')><IMG alt=գ��΢Ц src=face/zy.gif border=0></A> <A href=javascript:emoticon('{f6)')><IMG alt=�ѹ����� src=face/ng.gif border=0></A> <A href=javascript:emoticon('{f7)')><IMG alt=�����Ц�� src=face/kh.gif border=0></A> <A href=javascript:emoticon('{f8)')><IMG alt=ʧ������ src=face/sw.gif border=0></A> <a href=javascript:emoticon('{f9)')><IMG alt=���ε�Ц�� src=face/gg.gif border=0></a>

            </td>
          </tr>
        </table>
<br>
&nbsp;<input type=submit value=OK_���ˣ��޸����� name=B1>&nbsp; <input type=reset value=NO_���У���Ҫ��д name=B2> [��  Ctrl+Enter ֱ�ӷ���]<br><br>
        
      </td>
    </tr>  </table>
  </center>
</div></form><%else
riqi=now+timeset/24
myconn.execute("update [min] set zhuti='"&zhuti&"',body='"&body&"',face='"&face&"',orders='"&riqi&"' where id="&id&"")
%><br><br><%=t1%>�� �� �� ��<%=t2&d1%>���޸ĳɹ� <a href="show.asp?id=<%=reid%>&bd=<%=bd%>">
<font color="<%=c1%>">�ص�����ҳ�桤</font></a><%=d2%><%
end if%><br><!--#include file="down.asp"-->