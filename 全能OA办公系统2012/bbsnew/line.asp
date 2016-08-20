<%myconn.execute("delete*from online where ltime<now-0.005")
lgname=Request.Cookies(cn)("lgname")
ip=request.servervariables("remote_addr")
if lgname="" then
set jilu=myconn.execute("select ip from online where ip='"&ip&"'")
if jilu.eof then
set jilu=nothing
myconn.execute("insert into online (ip,ltime)VALUES('"&ip&"','"&now&"')")
else
myconn.execute("update online set ltime='"&now&"' where ip='"&ip&"'")
end if
end if
if lgname<>"" then
set ujilu=myconn.execute("select name from online where name='"&lgname&"'")
if ujilu.eof then
set ujilu=nothing
myconn.execute("delete*from online where ip='"&ip&"'")
myconn.execute("insert into online (name,ltime)VALUES('"&lgname&"','"&now&"')")
else
myconn.execute("update online set ltime='"&now&"' where name='"&name&"'")
end if
end if
usno=myconn.execute("Select count(ltime)from online where name<>''")(0)
lineno=myconn.execute("Select count(ltime)from online")(0)
if lineno>mostonline then
myconn.execute("update bbsinfo set mostonline='"&lineno&"'")
mostonline=lineno
end if
nusno=lineno-usno
%>