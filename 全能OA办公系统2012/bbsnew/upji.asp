<%
if 0<=q1 and 0<=m1 and 0<=j1 then
dj=0
end if
if 999<q1 and 299<m1 and 299<j1 then
dj=1
end if 
if 1999<q1 and 499<m1 and 499<j1 then
dj=2
end if 
if 2999<q1 and 699<m1 and 699<j1 then
dj=3
end if 
if 3999<q1 and 899<m1 and 899<j1 then
dj=4
end if 
if 4999<q1 and 1090<m1 and 1090<j1 then
dj=5
end if 
if 5999<q1 and 1299<m1 and 1299<j1 then
dj=6
end if 
if 6999<q1 and 1499<m1 and 1499<j1 then
dj=7
end if 
if 7999<q1 and 1699<m1 and 1699<j1 then
dj=8
end if 
if 8999<q1 and 1899<m1 and 1899<j1 then
dj=9
end if 
if 9999<q1 and 2990<m1 and 2990<j1 then
dj=10
end if 
if 10999<q1 and 2299<m1 and 2299<j1 then
dj=11
end if 
if 11999<q1 and 2499<m1 and 2499<j1 then
dj=12
end if 
if 12999<q1 and 2699<m1 and 2699<j1 then
dj=13
end if 
if 13999<q1 and 2899<m1 and 2899<j1 then
dj=14
end if
if sqltype="lg" then
sql1="select name from admin where name='"&lgname&"' and bd='70767766'"
sql="select name from admin where name='"&lgname&"'"
elseif sqltype="my" then
sql1="select name from admin where name='"&myname&"' and bd='70767766'"
sql="select name from admin where name='"&myname&"'"
end if
set mn=myconn.execute(sql)
if not mn.eof then
dj=15
end if
set mn=nothing
set mn1=myconn.execute(sql1)
if not mn1.eof then
dj=16
end if
set mn1=nothing
%>
<%
select case dj
case 0
dd="������·"
case 1
dd="��̳����"
case 2
dd="ҵ������"
case 3
dd="ְҵ����"
case 4
dd="��֮����"
case 5
dd="������"
case 6
dd="֩����"
case 7
dd="�����"
case 8
dd="С����"
case 9
dd="������"
case 10
dd="�����"
case 11
dd="��������"
case 12
dd="�ޱ���"
case 13
dd="��ʥ"
case 14
dd="�������"
case 15
dd="����"
case 16
dd="����Ա"
end select
%>