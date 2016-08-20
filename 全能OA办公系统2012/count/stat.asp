<!-- #include file="conn.asp"-->
<!-- #include file="cls_stat.asp"-->
<%
Response.Expires = 0
dim hxstat
set hxstat = New cls_stat

If Application.Contents(CacheName & "_isStart")=0 then
Call hxstat.OutPut
else
Call hxstat.StartCount
End If

set hxstat = nothing
set hx = nothing
%>