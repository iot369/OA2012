<!--#include file="up.asp"-->
<SCRIPT>
function showtb(tbnum)
{
whichEl = eval("tbtype" + tbnum);
if (whichEl.style.display == "none")
{
eval("tbtype" + tbnum + ".style.display=\"\";");
}
else
{
eval("tbtype" + tbnum + ".style.display=\"none\";");
}
}
</SCRIPT>
<%
'ʮ����
set fw=myconn.execute("select top 10 qian,name from user order by qian desc")
response.write"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<a onclick=showtb(1)><SPAN style='CURSOR: hand' ><img border=0 src=pic/fle.gif></span></a> <font color='#FFFFFF'><b>��̳ʮ����</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table id=tbtype1 onclick=showtb(1) border=1 cellpadding=0 cellspacing=0 style='display:none;border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=4% align=center height=24 background=pic/1.gif >���</td><td width=11% align=center background=pic/1.gif >�û���</td><td width=85% align=center background=pic/1.gif >��Ǯ�����Լ�����</td></tr>"
if fw.eof then
else
do while not fw.eof
qianall=qianall+fw("qian")*1
fw.movenext
loop
end if
i=1
fw.movefirst
do while not fw.eof
width=fw("qian")/qianall*90
response.write"<tr><td width=4% ><p style='margin: 4'><b>"&i&"</b>��</td><td width=11% ><p style='margin: 4'><a href='userinfo.asp?name="&kbbs(fw("name"))&"'>"&kbbs(fw("name"))&"</a></td><td width=85% ><p style='margin: 4'><img border=0 src=pic/line.gif width="&width&"% height=8> <b>"&fw("qian")&"</b></td></tr>"
i=i+1
if i=11 then exit do
fw.movenext
loop
fw.close
set fw=nothing
response.write"</table></center></div>"
%>
<%'ʮ��������ʿ
set ml=myconn.execute("select top 10 meili,name from user order by meili desc")
response.write"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<a onclick=showtb(2)><SPAN style='CURSOR: hand' ><img border=0 src=pic/fle.gif></span></a> <font color='#FFFFFF'><b>��̳ʮ��������ʿ</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table id=tbtype2 onclick=showtb(2) border=1 cellpadding=0 cellspacing=0 style='display:none;border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=4% align=center height=24 background=pic/1.gif >���</td><td width=11% align=center background=pic/1.gif >�û���</td><td width=85% align=center background=pic/1.gif >���������Լ�����</td></tr>"
if ml.eof then
else
do while not ml.eof
meiliall=meiliall+ml("meili")*1
ml.movenext
loop
end if
i=1
ml.movefirst
do while not ml.eof
width=ml("meili")/meiliall*90
response.write"<tr><td width=4% ><p style='margin: 4'><b>"&i&"</b>��</td><td width=11% ><p style='margin: 4'><a href='userinfo.asp?name="&kbbs(ml("name"))&"'>"&kbbs(ml("name"))&"</a></td><td width=85% ><p style='margin: 4'><img border=0 src=pic/line.gif width="&width&"% height=8> <b>"&ml("meili")&"</b></td></tr>"
i=i+1
if i=11 then exit do
ml.movenext
loop
ml.close
set ml=nothing
response.write"</table></center></div>"
%>
<%'ʮ�����о�����ʿ
set ml=myconn.execute("select top 10 jingyan,name from user order by jingyan desc")
response.write"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<a onclick=showtb(3)><SPAN style='CURSOR: hand' ><img border=0 src=pic/fle.gif></span></a> <font color='#FFFFFF'><b>��̳ʮ������ʿ</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table id=tbtype3 onclick=showtb(3) border=1 cellpadding=0 cellspacing=0 style='display:none;border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=4% align=center height=24 background=pic/1.gif >���</td><td width=11% align=center background=pic/1.gif >�û���</td><td width=85% align=center background=pic/1.gif >��������Լ�����</td></tr>"
if ml.eof then
else
do while not ml.eof
jingyanall=jingyanall+ml("jingyan")*1
ml.movenext
loop
end if
i=1
ml.movefirst
do while not ml.eof
width=ml("jingyan")/jingyanall*90
response.write"<tr><td width=4% ><p style='margin: 4'><b>"&i&"</b>��</td><td width=11% ><p style='margin: 4'><a href='userinfo.asp?name="&kbbs(ml("name"))&"'>"&kbbs(ml("name"))&"</a></td><td width=85% ><p style='margin: 4'><img border=0 src=pic/line.gif width="&width&"% height=8> <b>"&ml("jingyan")&"</b></td></tr>"
i=i+1
if i=11 then exit do
ml.movenext
loop
ml.close
set ml=nothing
response.write"</table></center></div>"
%>
<%'ʮ��������������
set ml=myconn.execute("select top 10 hits,zhuti,bd,id,name from min order by hits desc")
response.write"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<a onclick=showtb(4)><SPAN style='CURSOR: hand' ><img border=0 src=pic/fle.gif></span></a> <font color='#FFFFFF'><b>��̳ʮ��������������</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table id=tbtype4 onclick=showtb(4) border=1 cellpadding=0 cellspacing=0 style='display:none;border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=4% align=center height=24 background=pic/1.gif >���</td><td width=11% align=center background=pic/1.gif >��������</td><td width=85% align=center background=pic/1.gif >���������Լ�����ָ��</td></tr>"
if ml.eof then
else
do while not ml.eof
hitsall=hitsall+ml("hits")*1
ml.movenext
loop
end if
i=1
ml.movefirst
do while not ml.eof
width=ml("hits")/hitsall*90
response.write"<tr><td width=4% ><p style='margin: 4'><b>"&i&"</b>��</td><td width=11% ><p style='margin: 4'><a href='userinfo.asp?name="&kbbs(ml("name"))&"'>"&kbbs(ml("name"))&"</a></td><td width=85% ><p style='margin: 4'><img border=0 src=pic/fl.gif> <a href='show.asp?id="&ml("id")&"&bd="&ml("bd")&"'>"&kbbs(ml("zhuti"))&"</a><br><img border=0 src=pic/line.gif width="&width&"% height=8> <b>"&ml("hits")&"</b></td></tr>"
i=i+1
if i=11 then exit do
ml.movenext
loop
ml.close
set ml=nothing
response.write"</table></center></div>"
%>
<br><br><!--#include file="down.asp"-->