<!--#include file="conn.asp"-->
<!--#include file="cls_show.asp"-->
<!--#include file="skin.asp"-->
<%
Dim hxshow
set hxshow = New cls_show

Dim Action
Action=Request.QueryString("action")
Dim rs,IsPublic,ViewPass
set rs=hx.execute("select IsPublic,ViewPass from WebInfo where ID=1")
IsPublic = rs(0)
ViewPass = rs(1)
set rs=nothing
%>
<!--#include file="languages.asp"-->
<html>
<head>
<meta http-equiv="refresh" content="60">
<meta http-equiv="Content-Type" content="text/html; charset=<%=Lang.item("charset")%>">
<link href="style/style.css" rel="stylesheet" type="text/css">
<link href="style/style<%=skinid%>.css" rel="stylesheet" type="text/css">
<title><%=Lang.item("g_009")%></title></head>
<body leftmargin="0">
<%
if IsPublic=0 and Session(CacheName & "_ViewPass")<>"OK" then
	if Request.form("viewpass")="" then
		Call ShowForm
	elseif  IsPublic=0 and Request.form("viewpass")<>"" then
		if Request.form("viewpass")=viewpass then
			Session(CacheName & "_ViewPass")="OK"
		end if
		Response.redirect "show.asp"			
	end if
hx.ShowFooter
set hx=nothing
Response.end
end if

%>
<script src="js/fadeTicker.js" type="text/javascript"></script>
<script type="text/javascript">
function initFadeTicker()
{
	popupLoad.style.visibility='hidden'  //隐藏提示窗口

}
onload = initFadeTicker
</script>
<SCRIPT language=javascript src="js/mt_dropdownC.js"></SCRIPT>
<table width="768" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
              <tr align="center" class="TopLighNav1"> 
                <td width="11%"><a id=menu0 href="?action=main"><%=Lang.item("m_00")%></a></td>
                <td width="11%"><a id=menu1 href="?action=D"><%=Lang.item("m_10")%></a></td>
                <td width="11%"><a id=menu2 href="?action=PV"><%=Lang.item("m_20")%></a></td>
                <td width="11%"><a id=menu3 href="?action=R"><%=Lang.item("m_30")%></a></td>
                <td width="11%"><a id=menu4 href="?action=O"><%=Lang.item("m_40")%></a></td>
                <td width="11%"><a id=menu5 href="?action=Q"><%=Lang.item("m_50")%></a></td>
                <td width="11%"><a id=menu6 href="?action=I"><%=Lang.item("m_60")%></a></td>
                <td width="11%"><a id=menu7 href="?action=C"><%=Lang.item("m_70")%></a></td>
				<td width="11%"><a id=menu8 href="?action=<%=action%>&skinid=0"><%=Lang.item("m_80")%></a></td>
              </tr>
</table>
<SCRIPT language=javascript>

	if (mtDropDown.isSupported()) {

		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 0, mtDropDown.reference.bottomLeft);

		// menu : 0
		var menu0 = ms.addMenu(document.getElementById("menu0"));

		menu0.addItem(" <%=Lang.item("m_01")%>", "show.asp?action=main");
		menu0.addItem(" <%=Lang.item("m_02")%>", "show.asp?action=V");	
		menu0.addItem(" <%=Lang.item("m_03")%>", "show.asp?action=PV&ordernum=1");	
		//menu0.addItem(" 在线访客", "show.asp?action=Online");

		//var subMenu0 = menu0.addMenu(menu0.items[0]);
		//subMenu0.addItem(" 二级菜单","http://www.itlearner.com");


		// menu : 1
		var menu1 = ms.addMenu(document.getElementById("menu1"));
		menu1.addItem(" <%=Lang.item("m_11")%>", "show.asp?action=H");
		menu1.addItem(" <%=Lang.item("m_12")%>", "show.asp?action=D");
		menu1.addItem(" <%=Lang.item("m_13")%>", "show.asp?action=W");
		menu1.addItem(" <%=Lang.item("m_14")%>", "show.asp?action=M");
		menu1.addItem(" <%=Lang.item("m_15")%>", "show.asp?action=Y");

		// menu : 2
		var menu2 = ms.addMenu(document.getElementById("menu2"));
		menu2.addItem(" <%=Lang.item("m_21")%>", "show.asp?action=VT");
		menu2.addItem(" <%=Lang.item("m_22")%>", "show.asp?action=PV");

		// menu : 3
		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem(" <%=Lang.item("m_31")%>", "show.asp?action=R");
		menu3.addItem(" <%=Lang.item("m_32")%>", "show.asp?action=S");
		menu3.addItem(" <%=Lang.item("m_33")%>", "show.asp?action=Where");		
			
		// menu : 4
		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem(" <%=Lang.item("m_41")%>", "show.asp?action=O");
		menu4.addItem(" <%=Lang.item("m_42")%>", "show.asp?action=Width");
		menu4.addItem(" <%=Lang.item("m_43")%>", "show.asp?action=B");

		// menu : 5
		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem(" <%=Lang.item("m_51")%>", "show.asp?action=Q")
		menu5.addItem(" <%=Lang.item("m_52")%>", "show.asp?action=SQ2");
		menu5.addItem(" <%=Lang.item("m_53")%>", "show.asp?action=SQ1");
		menu5.addItem(" <%=Lang.item("m_54")%>", "show.asp?action=SQ3");

		// menu : 6
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem(" <%=Lang.item("m_61")%>", "show.asp?action=I");
		menu6.addItem(" <%=Lang.item("m_62")%>", "show.asp?action=SI3");
		menu6.addItem(" <%=Lang.item("m_63")%>", "show.asp?action=SI2");
		menu6.addItem(" <%=Lang.item("m_64")%>", "show.asp?action=SI1");
		
		// menu : 7
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		menu7.addItem(" <%=Lang.item("m_71")%>", "show.asp?action=C");
		menu7.addItem(" <%=Lang.item("m_72")%>", "http://www.itlearner.com/cutecounter/");
		menu7.addItem(" <%=Lang.item("m_73")%>", "http://www.itlearner.com/guestbook/");
		menu7.addItem(" <%=Lang.item("m_74")%>", "http://www.itlearner.com/");
		
		// menu : 8
		var menu8 = ms.addMenu(document.getElementById("menu8"));
		menu8.addItem(" <%=Lang.item("m_81")%>", "?action=<%=action%>&skinid=0");		
		menu8.addItem(" <%=Lang.item("m_82")%>", "?action=<%=action%>&skinid=1");
		menu8.addItem(" <%=Lang.item("m_83")%>", "?action=<%=action%>&skinid=2");
		menu8.addItem(" <%=Lang.item("m_84")%>", "?action=<%=action%>&skinid=3");
		menu8.addItem(" <%=Lang.item("m_85")%>", "?action=<%=action%>&skinid=4");
			
		mtDropDown.renderAll();
	}
init();
</SCRIPT>
<div name="loadpop" id="popupLoad" class="divcenter"> 
  <table width="200" border="0" cellspacing="0" cellpadding="5" class=popload>
    <tr> 
      <td align=center><p><font color="#000000"><%=Lang.item("g_010")%></p>
        <p><a href="http://www.itlearner.com" target="_blank">ITlearner</a> <a href="http://www.itlearner.com/cutecounter/" target="_blank">CuteCounter</a></p></td>
    </tr>
  </table>
</div>
<%

Dim query
query=hx.checkstr(Request("query"),25)

	'分页信息
	dim PageNo
	PageNo=Request.QueryString("PageNo")
	if PageNo="" or not isnumeric(PageNo) then
	PageNo=1
	else
	PageNo=int(PageNo)
	end if

ConnectionDatabase

select case Action 
case "main"
	call main
case "R"
	call referer
case "S"
	call RefSite
case "V"
	call LastRecord
case "H"
	call HourCount
case "D"
	call DayCount
case "W"
	call Weekcount
case "M"
	call MonthCount
case "Y"
	call YearCount
case "I"
	call IpCount
case "SI1"
	call SIpCount(1)
case "SI2"
	call SIpCount(2)
case "SI3"
	call SIpCount(3)
case "Where"
	call WhereCount
case "VT"
	call Page_VT	
case "PV"
	call Page_PV
case "Q"
	call Keyword
case "SQ1"
	call SQ1
case "SQ2"
	call SQ2
case "SQ3"
	call SQ3
case "O"
	call OsCount
case "B"
	call Browser
case "Width"
	call Width
case "C"
	call GetCode
case "Online"
	call Online
case else
	call main
end select
hx.ShowFooter
set hx=nothing

Function FormatNum(num,num2)
	if num=0 then
	FormatNum=0
	else
	FormatNum=FormatNumber(num,num2)
	end if
End Function

Function FormatTime(CurrentTime)
	if SelectedTimeZone = TimeZone then
		FormatTime=CurrentTime
	else
		FormatTime=DateAdd("h",SelectedTimeZone-TimeZone,CurrentTime)			
	end if
End Function


Sub ShowForm
	%>
	<table width="500" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
	  <tr class="tablebody1"><form name="form1" method="post" action="">
		<td><%=Lang.item("s_02")%>
        <input type="password" name="viewpass" class="input1">
		<input type="submit" name="Submit" value="<%=Lang.item("b_02")%>"></td>
		</form>
	  </tr>
	</table>
	<%
	End Sub
	Sub ShowQuery
	if IsPublic=2 or Session(CacheName & "_Admin")="OK" or Session(CacheName & "_ViewPass")="OK" then
	%>
	<table border="0" align="center" cellpadding="3" cellspacing="0" class="tableBorder2"><tr class="tablebody1"><form name="form" method="post" action="?action=<%=action%>"><td align="center"><%=Lang.item("s_01")%>
<input class="input1" name="query" type="text" id="query" size="25" maxlength="25" <%if query<>"" then response.write "value="""&query&""""%>><input type="submit" name="Submit" value="<%=Lang.item("b_01")%>"></td></form></tr></table>
	<%
	end if
	End Sub

	Sub MainTitle(str)
		Response.Write("<p align=center class=""menutitle"">:::::: " & str & " ::::::</p>")
	End Sub

	Sub Showinfo(str)
		dim Outstr
		Outstr = "<table width=500 border=0 align=center cellpadding=0 cellspacing=0><tr><td height=10></td></tr>"
		Outstr = Outstr & "<tr><td align=center>"
		Outstr = Outstr & str
		Outstr = Outstr & "</td></tr></table>"
		Response.write Outstr
	End Sub

Sub main
dim rs
dim startdate
dim day01,day02,day11,day12,day21,day22,day31,day32,day41,day42,Datenum,Month11,Month12
set rs=hx.execute("select top 1 CDate from CC_D order by id")
if not rs.eof  then
	StartDate=rs(0)
	Datenum=FormatNum(Now()-StartDate,1)
else
	StartDate=Date()
	Datenum=0
end if
set rs=nothing

set rs=hx.execute("select top 1 Visitor,PageView from CC_M order by CMonth desc")
if not rs.eof  then
	Month11=FormatNum(rs(0),0)
	Month12=FormatNum(rs(1),0)
else
	Month11=0
	Month12=0
end if
set rs=nothing

'if IsSqlDataBase = 1 then
'Dim Date1
'Date1=Date()
'set rs=hx.execute("select Visitor,PageView from CC_D where CDate='"&Date1&"'")
'else
'set rs=hx.execute("select Visitor,PageView from CC_D where CDate=Date()")
'end if
set rs=hx.execute("select Visitor,PageView from CC_D where CDate="&SqlDateString)

if not rs.eof then
	Day11=FormatNum(rs(0),0)
	Day12=FormatNum(rs(1),0)
else
	Day11=0
	Day12=0
end if	
set rs=nothing

if IsSqlDataBase = 1 then
Dim Date2
Date2=DateAdd("d",-1,Date())
set rs=hx.execute("select Visitor,PageView from CC_D where CDate='"&Date2&"'")
else
set rs=hx.execute("select Visitor,PageView from CC_D where CDate=DateAdd('d',-1,Date())")
end if
if not rs.eof then
	Day21=FormatNum(rs(0),0)
	Day22=FormatNum(rs(1),0)
else
	Day21=0
	Day22=0
end if	
set rs=nothing

set rs=hx.execute("select AVG(Visitor),AVG(PageView) from CC_D")
if not rs.eof then
	day41=rs(0)
	day42=rs(1)
	if not isnumeric(day41) then
		day41=0
	else
		day41=FormatNum(day41,0)
	end if
	if not isnumeric(day42) then
		day42=0
	else
		day42=FormatNum(day42,0)
	end if	
end if

Day31=FormatNum(Day11/(Now()-Date()),0)
Day32=FormatNum(Day12/(Now()-Date()),0)

Dim sql,sql2,rs2
sql="select Count(*) from CC_I where DateDiff("&TimeDiff(0)&",vtime,"&SqlNowString&")<"&OnlineTime
sql2="select Count(*) from CC_I where DateDiff("&TimeDiff(0)&",vtime,"&SqlNowString&")<"&OnlineTime*2

Dim OnlineNum
set rs=hx.execute(sql)
set rs2=hx.execute(sql2)
OnlineNum=(rs2(0)-rs(0))/2+rs(0)
if isnull(OnlineNum) then
OnlineNum=0
else
OnlineNum=FormatNum(OnlineNum,0)
end if

Day01=hx.execute("select Sum(Visitor) from CC_D")(0)
Day02=hx.execute("select Sum(PageView) from CC_D")(0)
if isnull(Day01) then
Day01=0
else
Day01=FormatNum(Day01,0)
end if
if isnull(Day02) then
Day02=0
else
Day02=FormatNum(Day02,0)
end if

Call MainTitle(Lang.item("m_01"))

set rs=hx.execute("select * from WebInfo where ID=1")
%>
<table width="768" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><%=Lang.item("currentLang")%></td>
  </tr>
</table>
<br>
<table width="768" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr align="center" class=tablebody1 id="tabletitlelink"> 
    <th><%=Lang.item("g_001")%></th>
    <th><%=Lang.item("g_002")%></th>
    <th><%=Lang.item("g_003")%></th>
    <th><%=Lang.item("g_004")%></th>
    <th><%=Lang.item("g_005")%></th>
    <th><a href="admin.asp" target="_blank"><%=Lang.item("g_006")%></a></th>
  </tr>
  <tr align="center" class=tablebody1> 
    <td><%=rs("WebName")%></td>
    <td><%=rs("WebAdmin")%></td>
    <td><%=StartDate%> </td>
    <td><%=Datenum%></td>
    <td><%=OnlineNum%></td>
    <td><%if Application.Contents(CacheName & "_isStart")=1 then
	response.write Lang.item("g_071")
	else
	response.write Lang.item("g_072")
	end if
%></td>
  </tr>
  <tr class=tablebody2> 
    <td align="center"><%=Lang.item("g_007")%>&nbsp;</td>
    <td colspan="5"><%="<a href="&rs("WebUrl")&">"&rs("WebUrl")&"</a>"%></td>
  </tr>
  <tr class=tablebody1> 
    <td align="center"><%=Lang.item("g_008")%>&nbsp;</td>
    <td colspan="5"><%=rs("WebIntro")%></td>
  </tr>
</table>
<br>
<table width="768" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr>
    <th> 
    <th><%=Lang.item("g_011")%></th>
    <th><%=Lang.item("g_012")%></th>
    <th><%=Lang.item("g_013")%></th>
    <th><%=Lang.item("g_014")%></th>
    <th><%=Lang.item("g_015")%></th>
    <th><%=Lang.item("g_016")%></th>
    <th>     
  </tr>
  <tr align="center" class=tablebody1>
    <td><%=Lang.item("g_017")%></td>
    <td><%=Day11%></td>
    <td><%=Day21%></td>
    <td><%=Day31%></td>
    <td><%=Day41%></td>
    <td><%=Month11%></td>
    <td><%=Day01%></td>
    <td rowspan="2"><script src=<%=hx.baseurl%>mystat.asp></script>
    </td>
  </tr>
  <tr align="center" class=tablebody1>
    <td><%=Lang.item("g_018")%></td>
    <td><%=Day12%></td>
    <td><%=Day22%></td>
    <td><%=Day32%></td>
    <td><%=Day42%></td>
    <td><%=Month12%></td>
    <td><%=Day02%></td>
  </tr>
</table>
<br>
<%
set rs=nothing
%>
<table width="768" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr id="tabletitlelink" align="center" bgcolor="#CCCCFF"> 
    <th width="70"><%=Lang.item("g_029")%></th>
    <th width="100"><%=Lang.item("g_030")%></th>
    <th width="275"><%=Lang.item("g_031")%></th>
    <th width="285"><%=Lang.item("g_032")%></th>
    <th width="38"><%=Lang.item("g_033")%></th>
  </tr>
  <%
	dim i
	dim condition,linkstr
	dim vpage,referer,q
	condition="where Dateandtime is not null"
	linkstr="action=V"
	sql="select top 10 Dateandtime,Ip,Page,Referer,User_Agent from CC_V "&condition&" order by Dateandtime desc,id asc"
	set rs=conn.execute(sql)
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=5>"&Lang.item("g_045")&"</td></tr>"	
	else
		dim NewVisitor
		NewVisitor=Rs.GetRows()
		for i=0 to UBound(NewVisitor,2)
		if i > 9 then exit for
		response.write "<tr class=tablebody1>"
		response.write "<td align=center>"&formatdatetime(FormatTime(NewVisitor(0,i)),3)&"</td>"
		response.write "<td align=center>"
		response.write "<a href=http://union.itlearner.com/ip/ipinfo.asp?ip="&NewVisitor(1,i)&" target=_blank title="""&Lang.item("g_074")&""">"
		response.write NewVisitor(1,i)
		response.write "</a>"
		response.write "</td>"
		response.write "<td>"
		if NewVisitor(2,i)="0" then
			vpage=Lang.item("g_053")
			response.write vpage
		else
			vpage=NewVisitor(2,i)
			response.write "<a href="""&vpage&""" target=_blank title="""&vpage&""">"&hx.OutStr(mid(vpage,8),40)&"</a>"
		end if
		response.write "</td>"
		response.write "<td>"
		if NewVisitor(3,i)="0" then
			referer=Lang.item("g_044")
			q=""
			response.write referer
		else
			referer=NewVisitor(3,i)
			response.write "<a href="""&referer&""" target=_blank title="""&referer&""">"&hx.OutStr(mid(referer,8),40)&"</a>"	
			q=hx.GetSearchKeyword(referer)	
			if q<>"" then
				q=ReadText(q)
			end if
		end if
		response.write "</td>"
		response.write "<td align=center><a href=# title='"&Lang.item("g_029")&":"&FormatTime(NewVisitor(0,i))&vbCRLF
		response.write Lang.item("g_069")&":"&NewVisitor(4,i)&vbCRLF
		response.write Lang.item("g_031")&":"&vpage&vbCRLF
		response.write Lang.item("g_032")&":"&referer&vbCRLF
		if q<>"" then
			response.write Lang.item("g_070")&":"&q
		end if
		response.write "'><font face=Wingdings>1</font></a></td>"	
		response.write "</tr>"
		next
	end if
	%>
</table>
<%
	set rs=nothing
End Sub
Sub GetCode
Call MainTitle(Lang.item("m_71"))
%>
<table border="0" align="center" cellpadding="0" cellspacing="0" class="tableBorder1">
  <tr> 
    <td><p><font color=#0033FF><%=Lang.item("g_057")%></font></p>
      <ul>
        <li><%=Lang.item("g_059")%>&lt;script src="<%=hx.baseurl%>mystat.asp"&gt;&lt;/script&gt;</li>
        <li><%=Lang.item("g_060")%>&lt;script src="<%=hx.baseurl%>mystat.asp?style=no"&gt;&lt;/script&gt;</li>
      </ul>
      <p><%=Lang.item("g_059")%><script src="<%=hx.baseurl%>mystat.asp"></script></p>
      <p>---------------------------------------------------------------------------------</p>

        <p><font color=#0033FF><%=Lang.item("g_058")%></font></p>
        <p>$AllVisitor -&gt; <%=Lang.item("g_016")&Lang.item("g_017")%><br>
		$AllPageView -&gt; <%=Lang.item("g_016")&Lang.item("g_018")%><br>
	    $TodayVisitor -&gt; <%=Lang.item("g_011")&Lang.item("g_017")%><br>
        $TodayPageView -&gt; <%=Lang.item("g_011")&Lang.item("g_018")%><br>
        $YestodayVisitor -&gt; <%=Lang.item("g_012")&Lang.item("g_017")%><br>
        $YestodayPageView -&gt; <%=Lang.item("g_012")&Lang.item("g_018")%><br>
</p>
		<%dim showstyle,p1,p2,p3,p4
		showstyle=server.htmlencode(request.form("showstyle"))
		if showstyle="" then
		showstyle=Lang.item("g_016")&Lang.item("g_018")&":$AllPageView"
		end if

        p1 = request("p1")
        p2 = request("p2")
        p3 = request("p3")
        p4 = request("p4")
        if p4 = "" then p4 = 0
	  %>
	  <table border="0" cellpadding="3" cellspacing="0" class="tableBorder2" width="650">
	  <tr class=tablebody1>
      <form name="form2" method="post" action="">
        <td>
        <input type="submit" name="Submit2" value="<%=Lang.item("b_04")%>">
        </td><td>
		<%
		dim temp1
		temp1 = split(Lang.item("g_143"),"|")		
		%>
        <input name="showstyle" type="text" id="showstyle" value="<%=showstyle%>" size="100"><br>
        <input type="radio" name="p1" value="0" <%if p1=0 or p1="" then response.write "checked"%>><%=temp1(0)%> <input type="radio" name="p1" value="1" <%if p1=1 then response.write "checked"%>><%=temp1(1)%>
        &nbsp;&nbsp;<%=temp1(2)%><select name="p2"><%dim i
        for i=1 to 10%><option<%if int(p2)=i then response.write " selected"%>><%=i%></option><%next%></select>&nbsp;&nbsp;
        <%=temp1(3)%><select name="p3"><%for i=1 to 3%><option<%if int(p3)=i then response.write " selected"%>><%=i%></option><%next%></select><%=temp1(4)%>

          <input type="checkbox" value="1" name="p4" <%if p4=1 then response.write "checked"%>><%=Lang.item("g_146")%>
      </td></form></tr></table>
	  <table border="0" cellpadding="3" cellspacing="0" class="tableBorder2" width="650">
	  <tr class=tablebody1><td>
      <p><%=Lang.item("g_061")%>
<%		showstyle = "str=" & server.URLEncode(showstyle)%>
        <textarea name="textarea" cols="100" rows="3"><script src="<%=hx.baseurl%>showstat.asp?<%=showstyle%>&p1=<%=p1%>&p2=<%=p2%>&p3=<%=p3%>&p4=<%=p4%>"></script></textarea>
      </p>
      <p><%=Lang.item("g_062")%><script src="<%=hx.baseurl%>showstat.asp?<%=showstyle%>&p1=<%=p1%>&p2=<%=p2%>&p3=<%=p3%>&p4=<%=p4%>"></script></p>
       </td></tr></table>
      </td>
  </tr>
</table>
<%
End Sub
%>
<%
Sub referer
Call MainTitle(Lang.item("m_31"))
Call ShowQuery
dim orderby,ordernum
ordernum=request("ordernum")
if ordernum=1 then
    orderby="vtime"
else
    ordernum=0
    orderby="CR"
end if
%>              
<table width="768"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF" id="tabletitlelink"> 
    <th width="420" align="center"><%=Lang.item("g_032")%></th>
	<th width="128" align="center"><a href=?action=R&ordernum=1&query=<%=query%>><%=Lang.item("g_066")%></a></th>
	<th width="220" align="center"><a href=?action=R&query=<%=query%>><%=Lang.item("g_038")%></a></th>
 </tr>
  <%
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition
	dim linkstr
	if query="" then
		condition=""
		linkstr="action=R&ordernum="&ordernum
	else
		condition="where Referer like '%"&query&"%'"
		linkstr="query="&query&"&action=R&ordernum="&ordernum
	end if
	sql="select top "&MaxRecord&" Referer,CR,vtime from CC_R "&condition&" order by "&orderby&" desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"	
	else
		rs.pagesize=MaxPageSize
		rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					if rs(0)="0" or rs(0)="" then
						response.write Lang.item("g_044")
					else
						response.write "<a href='"&rs(0)&"'  title='"&rs(0)&"' target=_blank>"&hx.OutStr(rs("Referer"),64)&"</a>"
					end if
					
					response.write "</td><td>"
  					response.write "&nbsp;"&FormatTime(rs(2))					
					response.write "</td><td>"	
					if ordernum=0 then	
					barwidth=FormatNum(rs(1)/maxnum*150,2)			
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
 					end if
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if				
					%>
</table>
<%
	hxshow.showPageInfo "CC_R","id",condition,PageNo,MaxPageSize,linkstr
%>
              <%End Sub%>
<%
Sub RefSite
Call MainTitle(Lang.item("m_32"))
Call ShowQuery
%>
<table width="768"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="558" align="center"><%=Lang.item("g_066")%></th>
    <th width="220" align="center"><%=Lang.item("g_066")%></th>
  </tr>
  <%
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition
	dim linkstr
	if query="" then
		condition=""
		linkstr="action=S"
	else
		condition="where RefSite like '%"&query&"%'"
		linkstr="query="&query&"&action=S"
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select top "&MaxRecord&" RefSite,CSite from RefSite "&condition&" order by CSite desc"
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
	else
		rs.pagesize=MaxPageSize
		rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					if rs(0)="0" then
						response.write Lang.item("g_044")
					else
						response.write "<a href='"&rs(0)&"'  title='"&rs(0)&"' target=_blank>"&left(rs(0),60)&"</a>"
					end if
						barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "RefSite","CSite",condition,PageNo,MaxPageSize,linkstr%>

<%End Sub%>
<%Sub HourCount
Call MainTitle(Lang.item("m_11"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_034")%></th>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_035")%></th>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_036")%></th>
  </tr>
  <%
				dim maxnum1,maxnum2,i,barwidth1,barwidth2 
				dim rs,sql
				maxnum1=hx.execute("select max(CTH) from CC_H")(0)
				if maxnum1=0 then maxnum1=1
				maxnum2=hx.execute("select max(CCH) from CC_H")(0)
				if maxnum2=0 then maxnum2=1
				sql="select Hour,CTH,CCH from CC_H order by Hour"
				set rs=hx.execute(sql)
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					i=0
					do while not rs.eof
					response.write "<tr class=tablebody1><td>"
					response.write rs(0)
					barwidth1=FormatNum(rs(1)/maxnum1*150,2)
					barwidth2=FormatNum(rs(2)/maxnum2*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
  					response.write rs(1)
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
  					response.write rs(2)
					response.write "</td>"
					response.write "</tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%End Sub%>
<%Sub DayCount
Call MainTitle(Lang.item("m_12"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_037")%></th>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_038")%></th>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_039")%></th>
  </tr>
  <%
				dim maxnum1,maxnum2,i,barwidth1,barwidth2 
				dim rs,sql
				set rs=server.createobject("adodb.recordset")
				sql="select top "&MaxRecord&" CDate,Visitor,PageView from CC_D order by CDate desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					maxnum1=hx.execute("select max(Visitor) from CC_D")(0)
					maxnum2=hx.execute("select max(PageView) from CC_D")(0)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td nowrap>"
					response.write rs(0)
					barwidth1=FormatNum(rs(1)/maxnum1*150,2)
					barwidth2=FormatNum(rs(2)/maxnum2*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
  					response.write rs(1)
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
  					response.write rs(2)
					response.write "</td>"
					response.write "</tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "CC_D","id","",PageNo,MaxPageSize,"action=D"%>
<%End Sub%>
<%Sub WeekCount
dim vweek(7,2)
dim maxnum1,maxnum2,i,barwidth1,barwidth2 
dim rs,sql
for i=1 to 7
if IsSqlDataBase = 1 then
sql="select top 1 Visitor,PageView from CC_D where DATEPART(weekday,CDate)="&i &" order by id desc"
else
sql="select top 1 Visitor,PageView from CC_D where weekday(CDate)="&i &" order by id desc"
end if
set rs=hx.execute(sql)
if rs.eof then
	vweek(i,0)=0
	vweek(i,1)=0
else
	vweek(i,0)=rs(0)
	vweek(i,1)=rs(1)
	if rs(0)>maxnum1 then maxnum1=rs(0)
	if rs(1)>maxnum2 then maxnum2=rs(1)
end if
next
if maxnum1=0 then maxnum1=1
if maxnum2=0 then maxnum2=1

Call MainTitle(Lang.item("m_13"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_040")%></td>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_038")%></td>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_039")%></td>
  </tr>
<%
for i=1 to 7
response.write "<tr class=tablebody1>"
response.write "<td align=center>"&findweek(i)&"</td>"
barwidth1=FormatNum(vweek(i,0)/maxnum1*150,2)
barwidth2=FormatNum(vweek(i,1)/maxnum2*150,2)
response.write "<td>"
response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
response.write vweek(i,0)
response.write "</td>"
response.write "<td>"
response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
response.write vweek(i,1)
response.write "</td>"
response.write "</tr>"
next		
%>
</table>
<%
for i=1 to 7
if IsSqlDataBase = 1 then
sql="select sum(Visitor),sum(PageView) from CC_D where DATEPART(weekday,CDate)="&i
else
sql="select sum(Visitor),sum(PageView) from CC_D where weekday(CDate)="&i
end if
set rs=hx.execute(sql)
if not isnumeric(rs(0)) then
	vweek(i,0)=0
else
	vweek(i,0)=rs(0)
end if
if not isnumeric(rs(1)) then
	vweek(i,1)=0
else
	vweek(i,1)=rs(1)
end if
if vweek(i,0)>maxnum1 then maxnum1=vweek(i,0)
if vweek(i,1)>maxnum2 then maxnum2=vweek(i,1)
next
if maxnum1=0 then maxnum1=1
if maxnum2=0 then maxnum2=1
%>
<br>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_041")%></td>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_038")%></td>
    <th width="220" align="center" bgcolor="#CCCCFF"><%=Lang.item("g_039")%></td>
  </tr>
<%
for i=1 to 7
response.write "<tr class=tablebody1>"
response.write "<td align=center>"&findweek(i)&"</td>"
barwidth1=FormatNum(vweek(i,0)/maxnum1*150,2)
barwidth2=FormatNum(vweek(i,1)/maxnum2*150,2)
response.write "<td>"
response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
response.write vweek(i,0)
response.write "</td>"
response.write "<td>"
response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
response.write vweek(i,1)
response.write "</td>"
response.write "</tr>"
next		
%>
</table>

<%
End Sub%>
<%Sub MonthCount
Call MainTitle(Lang.item("m_14"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center"><%=Lang.item("g_042")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
    <th width="220" align="center"><%=Lang.item("g_039")%></th>
  </tr>
  <%
				dim maxnum1,maxnum2,i,barwidth1,barwidth2
				dim rs,rs2,sql
				set rs=server.createobject("adodb.recordset")
				sql="select top "&MaxRecord&" CMonth,Visitor,PageView from CC_M order by CMonth desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					maxnum1=hx.execute("select max(Visitor) from CC_M")(0)
					maxnum2=hx.execute("select max(PageView) from CC_M")(0)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td align=center>"
					response.write rs(0)
					barwidth1=FormatNum(rs(1)/maxnum1*150,2)
					barwidth2=FormatNum(rs(2)/maxnum2*150,2)
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
  					response.write rs(1)					
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
  					response.write rs(2)										
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "CC_M","CMonth","",PageNo,MaxPageSize,"action=M"	%>
<%End Sub%>
<%Sub YearCount
Call MainTitle(Lang.item("m_15"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="60" align="center"><%=Lang.item("g_043")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
    <th width="220" align="center"><%=Lang.item("g_039")%></th>
  </tr>
  <%
				dim maxnum1,maxnum2,i,barwidth1,barwidth2
				dim rs,rs2,sql
				set rs=server.createobject("adodb.recordset")
				sql="select top "&MaxRecord&" CYear,Visitor,PageView from CC_Y order by CYear desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					maxnum1=hx.execute("select max(Visitor) from CC_Y")(0)
					maxnum2=hx.execute("select max(PageView) from CC_Y")(0)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td align=center>"
					response.write rs(0)
					barwidth1=FormatNum(rs(1)/maxnum1*150,2)
					barwidth2=FormatNum(rs(2)/maxnum2*150,2)
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth1&"> "   
  					response.write rs(1)					
					response.write "</td><td>"
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth2&"> "   
  					response.write rs(2)										
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "CC_Y","CYear","",PageNo,MaxPageSize,"action=Y"	%>
<%End Sub%>
<%Sub WhereCount
Call MainTitle(Lang.item("m_33"))
Call ShowQuery%> 
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("g_068")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
  </tr>
  <%
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition
	dim linkstr
	if query="" then
		condition=""
		linkstr="action=Where"
	else
		condition="where [Where] like '%"&query&"%'"
		linkstr="query="&query&"&action=Where"
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select top "&MaxRecord&" [Where],CW from CC_W "&condition&" order by CW desc"
	rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					response.write rs(0)
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "CC_W","id",condition,PageNo,MaxPageSize,linkstr%>
<%End Sub%>
<%Sub IpCount
Call MainTitle(Lang.item("m_61"))
Call ShowQuery%> 
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="140" align="center"><%=Lang.item("g_030")%></th>
    <th width="140" align="center"><%=Lang.item("g_066")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
  </tr>
  <%
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition
	dim linkstr
	if query="" then
		condition=""
		linkstr="action=I"
	else
		query=hx.checkstr(query,20)	
		condition="where Ip like '%"&query&"%'"
		linkstr="query="&query&"&action=I"
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select top "&MaxRecord&" Ip,CIP,vtime from CC_I "&condition&" order by CIP desc"
	rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					response.write "&nbsp;<a href=http://union.itlearner.com/ip/ipinfo.asp?ip="&rs(0)&" target=_blank>"
					response.write rs(0)
					response.write "</a>"
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"
					response.write "&nbsp;"&FormatTime(rs(2))
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%	hxshow.showPageInfo "CC_I","id",condition,PageNo,MaxPageSize,linkstr%>
<%End Sub%>
<%Sub SIpCount(num)
Select Case num
case 1
	Call MainTitle(Lang.item("m_64"))
case 2
	Call MainTitle(Lang.item("m_63"))
case 3
	Call MainTitle(Lang.item("m_62"))	
End Select
Call ShowQuery%> 
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("g_030")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
  </tr>
  <%
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition
	dim linkstr
	dim datatable '视图名称
	dim outstr
	datatable="v_SI"&num
	if query="" then
		condition=""
		linkstr="action="&action
	else
		condition="where Ip like '%"&query&"%'"
		linkstr="query="&query&"&action="&action
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select top "&MaxRecord&" Ip,SCIP from "&datatable&" "&condition&" order by SCIP desc"
	rs.open sql,conn,1,1
				if rs.eof then
					outstr = outstr & "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					outstr = outstr & "<tr class=tablebody1><td>&nbsp;"
					select case num
					case 1
						outstr = outstr & rs(0)&".*.*.*"
					case 2
						outstr = outstr & rs(0)&".*.*"
					case 3												
						outstr = outstr & rs(0)&".*"
					end select					
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					outstr = outstr & "</td><td>"					
 					outstr = outstr & "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					outstr = outstr & rs(1)
					outstr = outstr & "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
				response.write outstr
					%>
</table>
<%	hxshow.showPageInfo datatable,"0",condition,PageNo,MaxPageSize,linkstr%>
<%End Sub%>
<%Sub Page_PV()
dim orderby,ordernum
ordernum=request("ordernum")
if ordernum=1 then
    Call MainTitle(Lang.item("m_03"))
    Call ShowQuery
    orderby="vtime"
else
    ordernum=0
    Call MainTitle(Lang.item("m_22"))
    Call ShowQuery
    Call Showinfo(Lang.item("g_065"))
    orderby="Visitor+PageView"    
end if
%> 
<table width="768"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="420" align="center"><%=Lang.item("g_031")%></th>
    <th width="128" align="center"><a href=?action=PV&ordernum=1&query=<%=query%>><%=Lang.item("g_066")%></a></th>
    <th width="220" align="center"><a href=?action=PV&query=<%=query%>><%=Lang.item("g_039")%></a></th>
  <%
				dim maxnum,i,barwidth
				dim rs,sql
				dim condition
				dim linkstr
				if query="" then
					condition=""
					linkstr="action=PV&ordernum="&ordernum
				else
					condition="where [Page] like '%"&query&"%'"
					linkstr="query="&query&"&action=PV&ordernum="&ordernum
				end if
				set rs=server.createobject("adodb.recordset")
				sql="select top "&MaxRecord&" Page,Visitor+PageView,vtime from CC_P "&condition&" order by "&orderby&" desc"
				'response.write sql
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					do while not rs.eof and i<MaxPageSize
					if i=0 then maxnum=rs(1)
					response.write "<tr class=tablebody1><td>"
					if rs(0)="0" then
						response.write Lang.item("g_053")
					else
						response.write "<a href='"&rs(0)&"' title='"&rs(0)&"' target=_blank>"&hx.OutStr(rs(0),65)&"</a>"
					end if
						
					response.write "</td><td>"
					response.write "&nbsp;"&FormatTime(rs(2))					
					response.write "</td><td>"
				if ordernum=0 then
					barwidth=FormatNum(rs(1)/maxnum*150,2)
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
				end if
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
				end if
				set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "CC_P","id",condition,PageNo,MaxPageSize,linkstr%>
              
<%End Sub%>
<%Sub Page_VT
Call MainTitle(Lang.item("m_21"))
Call ShowQuery
Call Showinfo(Lang.item("g_065")) 
dim orderby,ordernum
ordernum=request("ordernum")
if ordernum=1 then
    orderby="vtime"
else
    ordernum=0
    orderby="Visitor"
end if
				dim maxnum,i,barwidth
				dim rs,sql
				dim condition
				dim linkstr
				if query="" then
					condition="where Visitor>0"
					linkstr="action=VT&ordernum="&ordernum
				else
					condition="where Visitor>0 and Page like '%"&query&"%'"
					linkstr="query="&query&"&action=VT&ordernum="&ordernum
				end if
%> 
<table width="768"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="420" align="center"><%=Lang.item("g_031")%></th>
    <th width="128" align="center"><a href=?action=VT&ordernum=1&query=<%=query%>><%=Lang.item("g_066")%></a></th>
    <th width="220" align="center"><a href=?action=VT&query=<%=query%>><%=Lang.item("g_038")%></a></th>
  <%

				set rs=server.createobject("adodb.recordset")
				sql="select top "&MaxRecord&" Page,Visitor,vtime from CC_P "&condition&" order by "&orderby&" desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"
				else
					rs.PageSize = MaxPageSize
					rs.absolutepage=PageNo
					i=0
					if i=0 then maxnum=rs(1)
					do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					if rs(0)="0" then
						response.write Lang.item("g_053")
					else
						response.write "<a href='"&rs(0)&"' title='"&rs(0)&"' target=_blank>"&hx.OutStr(rs(0),65)&"</a>"
					end if						
					response.write "</td><td>"
					response.write "&nbsp;"&FormatTime(rs(2))					
					response.write "</td><td>"
					if ordernum=0 then
					barwidth=FormatNum(rs(1)/maxnum*150,2)
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
 					end if
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
				end if
				set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "CC_P","id",condition,PageNo,MaxPageSize,linkstr%>
              
<%End Sub%>

<%Sub Keyword()
Call MainTitle(Lang.item("m_500"))
Call ShowQuery
Call Showinfo(Lang.item("m_511"))
dim orderby,ordernum
ordernum=request("ordernum")
if ordernum=1 then
    orderby="vtime"
else
    ordernum=0
    orderby="CR"
end if
%>
<table width="768"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="210" align="center"><%=Lang.item("g_070")%></th>
    <th width="210" align="center"><%=Lang.item("g_073")%></th> 
    <th width="128" align="center"><a href=?action=Q&ordernum=1&query=<%=query%>><%=Lang.item("g_066")%></a></th>
	<th width="220" align="center"><a href=?action=Q&query=<%=query%>><%=Lang.item("g_038")%></a></th> </tr>
  <%
	dim q
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition,linkstr
	if query="" then
		condition="where Q is not null"
		linkstr="action=Q&ordernum="&ordernum
	else
		condition="where (Q like '%"&encodeURIComponent(query)&"%' or Q like '%"&AnsiCode(query)&"%' or RefSite like '%"&query&"%') and Q is not null"
		linkstr="query="&query&"&action=Q&ordernum="&ordernum
	end if
	sql="select top "&MaxRecord&" Q,CR,RefSite,Referer,vtime from CC_R "&condition&" order by "&orderby&" desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=4>"&Lang.item("g_045")&"</td></tr>"	
	else
		rs.PageSize = MaxPageSize
		rs.absolutepage=PageNo
			i=0
			if i=0 then maxnum=rs(1)
				do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					q=ReadText(rs(0))
					response.write "<a href='"&rs(3)&"' target='_blank'>"
					response.write q					
					response.write "</a>"					
					response.write "</td><td>"
					response.write "<a href="&rs(2)&">"&rs(2)&"</a>"					
					response.write "</td><td>"
					response.write "&nbsp;"&FormatTime(rs(4))					
					response.write "</td><td>"	
					if ordernum=0 then	
					barwidth=FormatNum(rs(1)/maxnum*150,2)							
					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
					end if
					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
				rs.movenext
				loop
			end if
	rs.close:set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "CC_R","id",condition,PageNo,MaxPageSize,linkstr %>
<%End Sub%>
<%Sub SQ1
Call MainTitle(Lang.item("m_500"))
Call ShowQuery
Call Showinfo(Lang.item("m_521"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("g_070")%></td>              
    <th width="220" align="center"><%=Lang.item("g_038")%></td>
  </tr>
                <%
	dim q
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition,linkstr
	if query="" then
		condition=""
		linkstr="action=SQ1"
	else
		condition="where Q like '%"&encodeURIComponent(query)&"%' or Q like '%"&AnsiCode(query)&"%'"
		linkstr="query="&query&"&action=SQ1"
	end if
	sql="select top "&MaxRecord&" Q,SCR from v_SQ1 "&condition&" order by SCR desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"	
	else
		rs.PageSize = MaxPageSize
		rs.absolutepage=PageNo
			i=0
			if i=0 then maxnum=rs(1)
				do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					q=ReadText(rs(0))
					response.write q					
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td>"				
					response.write "<td>"									
					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
				rs.movenext
				loop
			end if
	rs.close:set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "v_SQ1","SCR",condition,PageNo,MaxPageSize,linkstr %>
<%End Sub%>
<%Sub SQ3
Call MainTitle(Lang.item("m_500"))
Call ShowQuery
Call Showinfo(Lang.item("m_531"))
%>

<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("g_073")%></td>              
    <th width="220" align="center"><%=Lang.item("g_038")%></td>
  </tr>
                <%
	dim q
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition,linkstr
	if query="" then
		condition=""
		linkstr="action=SQ3"
	else
		condition="where RefSite like '%"&query&"%'"
		linkstr="query="&query&"&action=SQ3"
	end if
	sql="select top "&MaxRecord&" RefSite,SCR from v_SQ3 "&condition&" order by SCR desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"	
	else
		rs.PageSize = MaxPageSize
		rs.absolutepage=PageNo
			i=0
			if i=0 then maxnum=rs(1)
				do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					q=ReadText(rs(0))
					response.write q					
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td>"				
					response.write "<td>"									
					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
				rs.movenext
				loop
			end if
	rs.close:set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "v_SQ3","SCR",condition,PageNo,MaxPageSize,linkstr %>
<%End Sub%>
<%Sub SQ2
Call MainTitle(Lang.item("m_500"))
Call ShowQuery
Call Showinfo(Lang.item("m_541"))
%>
<table width="600"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="180" align="center"><%=Lang.item("g_070")%></th>
    <th width="120" align="center"><%=Lang.item("g_073")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
  </tr>
  <%
	dim q
	dim maxnum,i,barwidth 
	dim rs,sql
	dim condition,linkstr
	if query="" then
		condition=""
		linkstr="action=SQ2"
	else
		condition="where Q like '%"&encodeURIComponent(query)&"%' or Q like '%"&AnsiCode(query)&"%' or RefSite like '%"&query&"%'"
		linkstr="query="&query&"&action=SQ2"
	end if

	sql="select top "&MaxRecord&" Q,RefSite,SCR from v_SQ2 "&condition&" order by SCR desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=3>"&Lang.item("g_045")&"</td></tr>"	
	else
		rs.PageSize = MaxPageSize
		rs.absolutepage=PageNo
			i=0
			if i=0 then maxnum=rs(2)
				do while not rs.eof and i<MaxPageSize
					response.write "<tr class=tablebody1><td>"
					q=ReadText(rs(0))
					response.write q					
					barwidth=FormatNum(rs(2)/maxnum*150,2)
					response.write "</td>"
					response.write "<td>"&rs(1)&"</td>"
					response.write "<td>"									
					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
					response.write rs(2)
					response.write "</td></tr>"
					i=i+1
				rs.movenext
				loop
			end if
	rs.close:set rs=nothing
					%>
</table>
<%	hxshow.showPageInfo "v_SQ2","SCR",condition,PageNo,MaxPageSize,linkstr %>
<%End Sub%>

<%Sub OsCount
Call MainTitle(Lang.item("m_41"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("m_41")%></td>
    <th width="220" align="center"><%=Lang.item("g_038")%></td>
  </tr>
  <%
				dim maxnum,i,barwidth 
				dim rs,sql
				set rs=server.createobject("adodb.recordset")
				sql="select Client,CC from CC_C where left(id,1)=1 order by CC desc,id asc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
				else
					i=0
					if i=0 then maxnum=rs(1)
					if maxnum=0 then maxnum=1
					do while not rs.eof
					response.write "<tr class=tablebody1><td>"
					response.write rs(0)
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%End Sub%>
<%Sub Width
Call MainTitle(Lang.item("m_42"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("m_42")%></th>
    <th width="220" align="center"><%=Lang.item("g_038")%></th>
  </tr>
  <%
				dim maxnum,i,barwidth 
				dim rs,sql
				set rs=server.createobject("adodb.recordset")
				sql="select Client,CC from CC_C where left(id,1)=3 order by CC desc,id asc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
				else
					i=0
					if i=0 then maxnum=rs(1)
					if maxnum=0 then maxnum=1
					do while not rs.eof
					response.write "<tr class=tablebody1><td>"
					response.write rs(0)
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%End Sub%>
<%Sub Browser
Call MainTitle(Lang.item("m_43"))
%>
<table width="500"  align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr bgcolor="#CCCCFF"> 
    <th width="280" align="center"><%=Lang.item("m_43")%></td>
    <th width="220" align="center"><%=Lang.item("g_038")%></td>
  </tr>
  <%
				dim maxnum,i,barwidth 
				dim rs,sql
				set rs=server.createobject("adodb.recordset")
				sql="select Client,CC from CC_C where left(id,1)=2 order by CC desc,id asc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.write "<tr class=tablebody1><td colspan=2>"&Lang.item("g_045")&"</td></tr>"
				else
					i=0
					if i=0 then maxnum=rs(1)
					if maxnum=0 then maxnum=1
					do while not rs.eof
					response.write "<tr class=tablebody1><td>"
					response.write rs(0)
					barwidth=FormatNum(rs(1)/maxnum*150,2)
					response.write "</td><td>"					
 					response.write "<img height=12 align=absmiddle class=PicBar width="&barwidth&"> "   
  					response.write rs(1)
					response.write "</td></tr>"
					i=i+1
					rs.movenext
					loop
					set rs=nothing
				end if
					%>
</table>
<%End Sub%>
<%Sub LastRecord
Call MainTitle(Lang.item("m_02"))
Call ShowQuery%>	
<table width="768" align="center" cellpadding="3" cellspacing="1" class="tableBorder2">
  <tr id="tabletitlelink" align="center" bgcolor="#CCCCFF"> 
    <th width="70"><%=Lang.item("g_029")%></th>
    <th width="100"><%=Lang.item("g_030")%></th>
    <th width="275"><%=Lang.item("g_031")%></th>
    <th width="285"><%=Lang.item("g_032")%></th>
    <th width="38"><%=Lang.item("g_033")%></th>
  </tr>
  <%
	dim rs,sql
	dim i
	dim condition,linkstr
	dim vpage,referer,q
	if query="" then
		condition="where Dateandtime is not null"
		linkstr="action=V"
	else
		condition="where (Ip like '%"&query&"%' or Page like '%"&query&"%' or Referer like '%"&query&"%') and Dateandtime is not null"
		linkstr="query="&server.HTMLEncode(query)&"&action=V"
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select top "&MaxRecord&" * from CC_V "&condition&" order by Dateandtime desc,id asc"
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<tr class=tablebody1><td colspan=5>"&Lang.item("g_045")&"</td></tr>"	
	else
		i=0
		rs.PageSize = MaxPageSize
		rs.absolutepage=PageNo
		do while not rs.eof and i<MaxPageSize
		response.write "<tr class=tablebody1>"
		response.write "<td align=center>"&formatdatetime(FormatTime(rs("Dateandtime")),3)&"</td>"
		response.write "<td align=center>"
		response.write "<a href=http://union.itlearner.com/ip/ipinfo.asp?ip="&rs("Ip")&" target=_blank title="""&Lang.item("g_074")&""">"
		response.write rs("Ip")
		response.write "</a>"
		response.write "</td>"
		response.write "<td>"
		if rs("Page")="0" then
			vpage=Lang.item("g_053")
			response.write vpage
		else
			vpage=rs("Page")
			response.write "<a href="""&vpage&""" target=_blank title="""&vpage&""">"&hx.OutStr(mid(vpage,8),40)&"</a>"
		end if
		response.write "</td>"
		response.write "<td>"
		if rs("Referer")="0" then
			referer=Lang.item("g_044")
			q=""
			response.write referer
		else
			referer=rs("Referer")
			response.write "<a href="""&referer&""" target=_blank title="""&referer&""">"&hx.OutStr(mid(referer,8),40)&"</a>"	
			q=hx.GetSearchKeyword(referer)	
			if q<>"" then
				q=ReadText(q)
			end if
		end if
		response.write "</td>"
		response.write "<td align=center><a href=# title='"&Lang.item("g_029")&":"&FormatTime(rs("Dateandtime"))&vbCRLF
		response.write Lang.item("g_069")&":"&rs("User_Agent")&vbCRLF
		response.write Lang.item("g_031")&":"&vpage&vbCRLF
		response.write Lang.item("g_032")&":"&referer&vbCRLF
		if q<>"" then
			response.write Lang.item("g_070")&":"&q
		end if
		response.write "'><font face=Wingdings>1</font></a></td>"	
		response.write "</tr>"
		rs.movenext
		i=i+1
		loop
	end if
	%>
</table>
<%
	set rs=nothing
	hxshow.showPageInfo "CC_V","id",condition,PageNo,MaxPageSize,linkstr
End Sub%>
<%
	'将星期序号翻译为汉字
	
	Function findweek(theweek)
		select case theweek
		case 1
			findweek=Lang.item("g_046")
		case 2
			findweek=Lang.item("g_047")
		case 3
			findweek=Lang.item("g_048")
		case 4
			findweek=Lang.item("g_049")
		case 5
			findweek=Lang.item("g_050")
		case 6
			findweek=Lang.item("g_051")
		case 7
			findweek=Lang.item("g_052")
		end select
	End Function

	Function AnsiCode(vstrIn)
		Dim i, strReturn, innerCode, ThisChr
		Dim Hight8, Low8
		strReturn = "" 
		For i = 1 To Len(vstrIn) 
			ThisChr = Mid(vStrIn,i,1) 
			If Abs(Asc(ThisChr)) < &HFF Then 
				strReturn = strReturn & ThisChr 
			Else
				innerCode = Asc(ThisChr)
				If innerCode < 0 Then
					innerCode = innerCode + &H10000
				End If
				Hight8 = (innerCode And &HFF00) \ &HFF
				Low8 = innerCode And &HFF
				strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
			End If 
		Next 
		AnsiCode = strReturn 
	End Function
	
	Function DeCodeAnsi(s)
		Dim i, sTmp, sResult, sTmp1
		sResult = ""
		For i=1 To Len(s)
			If Mid(s,i,1)="%" Then
				sTmp = "&H" & Mid(s,i+1,2)
				If isNumeric(sTmp) Then
					If CInt(sTmp)=0 Then
						i = i + 2
					ElseIf CInt(sTmp)>0 And CInt(sTmp)<128 Then
						sResult = sResult & Chr(sTmp)
						i = i + 2
					Else
						If Mid(s,i+3,1)="%" Then
							sTmp1 = "&H" & Mid(s,i+4,2)
							If isNumeric(sTmp1) Then
								sResult = sResult & Chr(CInt(sTmp)*16*16 + CInt(sTmp1))
								i = i + 5
							End If
						Else
							sResult = sResult & Chr(sTmp)
							i = i + 2
						End If
					End If
				Else
					sResult = sResult & Mid(s,i,1)
				End If
			Else
				sResult = sResult & Mid(s,i,1)
			End If
		Next
		DeCodeAnsi = sResult
	End Function
%>
<script language="JScript" runat="server">
	function ReadText(s){
		try{
			return decodeURIComponent(s);
		}catch(e){
			try{		//某些服务器用此函数会出错
			return DeCodeAnsi(s);
			}catch(e){   
      			return "&nbsp;"; 
			}
		}

	}
</script>
</body>
</html>