<!-- #include file="conn.asp"-->
<%
dim str
dim p1,p2,p3,p4
str = server.HTMLEncode(request("str"))
'ͼƬ����������ʾ,0��ʾ���֣�1��ʾͼƬ
p1 = request("p1")
'ͼƬ��ʾ��������ʾ��ʾλ����������6λ������ʾ000234������Ϊ0��Ϊ234����Χ0~10
p2 = request("p2")
'ͼƬ��ʾ��������ʾʹ����һ��ͼƬ(ȡֵ��Χ1-3)
p3 = request("p3")
'�Ƿ��ָ��ͳ����ʾҳ�������
p4 = request("p4")

if isnumeric(p1) then
    if p1<>0 and p1<>1 then p1=0
else
    p1=0
end if

if isnumeric(p2) then
	if p2>10 or p2<0 then p2=0
else
	p2=0
end if
if isnumeric(p3) then
	if p3>3 or p3<1 then p3=1
else
	p3=1
end if
if p4<>"" and isnumeric(p4) then
	if p4>1 or p4<0 then p4=1
else
	p4=1
end if


Call OutPut(str)

Sub OutPut(str)	
	dim outstr		
	dim style,scount
	dim AllVisitor,AllPageView
	dim TodayPageView,TodayVisitor
	dim YestodayPageView,YestodayVisitor
	
	if instr(str,"$AllVisitor") then
		AllVisitor=hx.execute("select Sum(Visitor) from CC_D")(0)
		if YVisitor > 0 then
		AllVisitor = AllVisitor + YVisitor
		end if
		str=replace(str,"$AllVisitor",ShowPic(AllVisitor,p1,p2,p3))
	end if
	
	if instr(str,"$AllPageView") then
		AllPageView=hx.execute("select Sum(PageView) from CC_D")(0)
		if YPageView > 0 then
		AllPageView = AllPageView + YPageView
		end if
		str=replace(str,"$AllPageView",ShowPic(AllPageView,p1,p2,p3))
	end if
	
	dim ors
	if instr(str,"$TodayPageView") or instr(str,"$TodayVisitor") then
		If IsSqlDataBase = 1 Then
			Dim Date1
			Date1=Date()
			set ors=hx.execute("select Visitor,PageView from CC_D where CDate='"&Date1&"'")
		else
			set ors=hx.execute("select Visitor,PageView from CC_D where CDate=date()")	
		end if
		if ors.eof then
			TodayVisitor=0
			TodayPageView=0
		else
			TodayVisitor=ors(0)
			TodayPageView=ors(1)
		end if
		set ors=nothing
		str=replace(str,"$TodayPageView",ShowPic(TodayPageView,p1,p2,p3))
		str=replace(str,"$TodayVisitor",ShowPic(TodayVisitor,p1,p2,p3))
	end if
	
	if instr(str,"$YestodayPageView") or instr(str,"$YestodayVisitor") then
		If IsSqlDataBase = 1 Then
			Dim Date2
			Date2=DateAdd("d",-1,Date())
			set ors=hx.execute("select Visitor,PageView from CC_D where CDate='"&Date2&"'")
		else
			set ors=hx.execute("select Visitor,PageView from CC_D where CDate=DateAdd('d',-1,Date())")	
		end if
		if ors.eof then
			YestodayVisitor=0
			YestodayPageView=0
		else
			YestodayVisitor=ors(0)
			YestodayPageView=ors(1)
		end if
		set ors=nothing
		str=replace(str,"$YestodayPageView",ShowPic(YestodayPageView,p1,p2,p3))
		str=replace(str,"$YestodayVisitor",ShowPic(YestodayVisitor,p1,p2,p3))
	end if
	
	if p4 = 1 then
		outstr = "<a href="&hx.baseurl&"show.asp target=_blank>" & str & "</a>"
	else
		outstr = str
	end if
	
	response.write "document.write("& chr(34) & outstr & chr(34) &")"
End Sub

Function ShowPic(scount,p1,p2,p3)
	if isnull(scount) then
		scount = 0
	else
		scount = CStr(scount)
	end if
  for i=len(scount) to p2-1
  	scount = "0" & scount
  next

	if p1=0 then
	ShowPic=scount
	else
	Dim i
	For i = 1 to Len(scount)
		ShowPic = ShowPic & "<IMG SRC="&hx.baseurl&"images/"&p3&"/" & Mid(scount, i, 1) & ".gif border=0>"
	Next
	end if
End Function

function longnum(innum)
  longnum=cstr(innum)
  if numlong <> 0 then
    for i=numlong-1 to 1 step -1
      if innum < 10^i then longnum = "0" & longnum
    next
  end if
end function
%>