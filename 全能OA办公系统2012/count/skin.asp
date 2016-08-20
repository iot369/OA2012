<%
dim skinid,cookieskinid
skinid=request("skinid")
cookieskinid=request.cookies(CacheName&"_skin")("id")
if cookieskinid="" and skinid="" then skinid=Skin
if skinid="" or not isnumeric(skinid) then skinid=cookieskinid
if skinid>4 or skinid<0 then skinid=cookieskinid
if Cint(skinid)<>Cint(cookieskinid) then
	response.cookies(CacheName&"_skin")("id")=skinid
	response.cookies(CacheName&"_skin").expires=Dateadd("yyyy",1,now())
end if
%>