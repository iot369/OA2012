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
dd="新手上路"
case 1
dd="论坛游侠"
case 2
dd="业余侠客"
case 3
dd="职业侠客"
case 4
dd="侠之大者"
case 5
dd="蝙蝠侠"
case 6
dd="蜘蛛侠"
case 7
dd="青蜂侠"
case 8
dd="小飞侠"
case 9
dd="蒙面侠"
case 10
dd="火箭侠"
case 11
dd="城市猎人"
case 12
dd="罗宾汉"
case 13
dd="侠圣"
case 14
dd="超级贵宾"
case 15
dd="版主"
case 16
dd="管理员"
end select
%>