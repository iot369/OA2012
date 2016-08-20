<!--#include file="conn.asp"-->
<%response.buffer=true
url=Request.ServerVariables("HTTP_REFERER")
if instr(url,"bd=")>0 then
bd=request.querystring("bd")
sty="bbs"&bd
else
sty="all"
end if
skin=request.querystring("skin")
select case skin
case"a"
c1="#5E8ACA"
c2="#F9FBFD"
case"b"
c1="#5E8ACA"
c2="#F9FBFD"
case"c"
c1="#5E8ACA"
c2="#F9FBFD"
case"d"
c1="#5E8ACA"
c2="#F9FBFD"
case"e"
c1="#5E8ACA"
c2="#F9FBFD"
end select
Response.Cookies(cn&"1")(sty)=skin
Response.Cookies(cn&"1")(sty&"c1")=c1
Response.Cookies(cn&"1")(sty&"c2")=c2
Response.Cookies(cn&"1").Expires=date+365
response.redirect url
%>