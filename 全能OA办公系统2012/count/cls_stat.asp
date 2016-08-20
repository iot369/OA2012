<%
class cls_stat
	Public ip,vpage
	Public vHour,vDate
	Public rs
	Dim Referer,RefSite,User_Agent,Client_id,width
	Dim sql
	Dim q,Sip

	Private Sub Class_Initialize()
		ip=hx.CheckStr(Request.ServerVariables("Remote_Addr"),15)
		vpage=hx.checkstr(Request.ServerVariables("HTTP_REFERER"),250)	
		vHour=hour(now())
		vDate=date()
	End Sub

	Public Sub StartCount
		Checkvpage()
		Init()
		ip=GetIp(ip)		
		if IsVisited(ip) then
			CountPage "PageView",vpage
			OutPut
			response.end
		else
			CountPage "Visitor",vpage
		end if


'处理Referer信息
	GetReferer()
	User_Agent=hx.checkstr(Request.ServerVariables("HTTP_USER_AGENT"),250)
'操作系统、浏览器、屏幕宽度
	GetClient()
'处理来源地区
	GetWhere()
	vHour=hour(now())

	'将记录加入到Visitor表中	
	hx.execute("update CC_V set Ip='"&ip&"',Referer='"&Referer&"',Page='"&vpage&"',User_Agent='"&User_Agent&"',Dateandtime="&SqlNowString&" where id=(select top 1 id from CC_V order by Dateandtime asc,id asc)")
	'处理Client信息
	hx.execute("update CC_C set CC=CC+1 where id in "& Client_id)
	'处理日统计信息
		hx.Execute("update CC_D set Visitor=Visitor+1 where CDate="&SqlDateString)				
	'处理小时信息
		hx.execute "update CC_H set CTH=CTH+1,CCH=CCH+1,vtime="&SqlNowString&" where Hour=" &vHour
		OutPut

	End Sub

	Private Sub GetReferer()
Referer=hx.CheckStr(Request("referer"),250)
'if right(Referer,1)="/" then Referer=left(Referer,len(Referer)-1)
If Referer<>"" Then
	RefSite=Mid(Referer,8)
	RefSite="http://"&Mid(RefSite,1,instr(RefSite,"/"))
	RefSite=hx.CheckStr(RefSite,100)
else
	Referer=0
	RefSite=0	
end if
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select CR,Referer,Q,RefSite,vtime from CC_R where Referer='"&Referer&"'"
		rs.open sql,conn,1,2
		if rs.eof then	
			rs.addnew
			rs(1)=Referer
			q=hx.GetSearchKeyword(Referer)	
			if q<>"" then
				rs(2)=q
			end if
			rs(3)=RefSite
			rs.update
		else
			rs(0)=rs(0)+1
			rs(4)=now()
			rs.update
		end if
		rs.close
		set rs=nothing
	End Sub

	Private Sub GetClient()
if instr(User_Agent,"Win 9x 4.90") then
	Client_id=107
elseif instr(User_Agent,"Windows 98") then
	Client_id=101
elseif instr(User_Agent,"Windows NT 5.1") then
	Client_id=102
elseif instr(User_Agent,"Windows NT 5.0") then
	Client_id=104
elseif instr(User_Agent,"Windows NT 5.2") then
	Client_id=105
elseif instr(User_Agent,"Windows NT") then
	Client_id=103
elseif instr(User_Agent,"unix")  or instr(User_Agent,"Linux") or instr(User_Agent,"SunOS") or instr(User_Agent,"BSD") then
	Client_id=106
else
	Client_id=108
end if
' 浏览器
if instr(User_Agent,"MSIE 6") then
	Client_id=Client_id & "," & 201 
elseif instr(User_Agent,"MSIE 5") then
	Client_id=Client_id & "," & 202
elseif instr(User_Agent,"MSIE 4") then
	Client_id=Client_id & "," & 203
elseif instr(User_Agent,"Netscape") then
	Client_id=Client_id & "," & 204
elseif instr(User_Agent,"Opera") then
	Client_id=Client_id & "," & 206
else
	Client_id=Client_id & "," & 207
end if
'屏幕宽度
width=Request("screenwidth")

if width="640" then
	Client_id=Client_id & "," & 301 
elseif width="800" then	
	Client_id=Client_id & "," & 302
elseif width="1024" then	
	Client_id=Client_id & "," & 303
elseif width="1152" then	
	Client_id=Client_id & "," & 304
elseif width="1280" then	
	Client_id=Client_id & "," & 305
elseif width="1600" then	
	Client_id=Client_id & "," & 306
else	
	Client_id=Client_id & "," & 307
end if

Client_id="(" & Client_id & ")"
	End Sub


	Private Sub Checkvpage()
		if right(vpage,1)="/" then vpage=left(vpage,len(vpage)-1)		
		if vpage="" or ip="" then
			OutPut
			response.end
		end if	
	End Sub

	Private Sub Init()
		Dim rs,sql
		Dim AppName1
		AppName1=CacheName & "_Date"
		If Application.Contents(AppName1) <> vDate Then
			if left(SysMode,1)="1" then '自动清理模式，清理以往数据
				Dim temp1,temp2
				temp1 = int(split(SysMode,"|")(1))
				temp2 = int(split(SysMode,"|")(2))
				hx.Execute("delete from CC_I where DateDiff('d',vtime,"&SqlNowString&")>"&temp1&" and CIP<"&temp2)
				hx.Execute("delete from CC_P where DateDiff('d',vtime,"&SqlNowString&")>"&temp1&" and Visitor+PageView<"&temp2)
				hx.Execute("delete from CC_R where DateDiff('d',vtime,"&SqlNowString&")>"&temp1&" and CR<"&temp2)
			end if
			Application.Contents(AppName1) = vDate
			sql="select CDate from CC_D where CDate="&SqlDateString
			set rs=hx.GetRs(sql,1,2)
			If rs.eof then
				rs.addnew
				rs(0)=vDate
				rs.update
			End If
			rs.close
		Else
			'处理PageView信息
			hx.Execute("update CC_D set PageView=PageView+1 where CDate="&SqlDateString)			
		End If
		
		Dim AppName2
		AppName2=CacheName & "_Hour"
		If Application.Contents(AppName2) <> vHour Then
			Application.Contents(AppName2) = vHour
			hx.Execute "update CC_H set CTH=0 where DATEDIFF("&TimeDiff(1)&",vtime,"&SqlNowString&") > 1 and Hour=" & vHour 
		End If
	End Sub

	Private Function IsVisited(ip)	'判断是否要重新记数
		Dim rs,sql
		If Len(Session.Contents(CacheName)) = 0 Then					
			Session.Contents(CacheName) = 1
			Sql = "select Ip,vtime,CIP from CC_I where Ip='" & ip & "'"
			set rs=hx.Getrs(Sql,1,2)
			If rs.EOF then
				rs.AddNew
				rs(0)=ip
				rs(1)=now()
				rs.update
				isVisited=False
			Elseif DateDiff("h",rs(1),now())>ExpireTime Then
				rs(1)=now()
				rs(2)=rs(2)+1
				rs.update	
				isVisited=False	
			Else
				isVisited=True
			End If
				rs.close
			set rs=nothing
		Else
			isVisited = True
		End If
			'isVisited = false
	End Function


	
	Public Function GetPage(PageUrl)
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		PageUrl = PageUrl
		re.Pattern = "[\?#].*"
		GetPage = re.Replace(PageUrl,"")
	End Function
	
	Private Function Getip(ip)
		Dim a,i
		a = Split(ip,".")
		if ubound(a)<>3 then Getip=0:Exit Function
		For i=0 to 3
			Sip= Sip + CInt(a(i)) * (256^(3-i))
 			Getip=Getip & String(3-Len(a(i)),"0") & a(i) & "."
		Next
		Getip=left(Getip,15)
	End Function	

	Private Function GetWhere()
		Dim sql,Where,rs2
		Dim ip_db,ip_conn,ip_connstr,ip_rs
		ip_db="data/ip.mdb"
		ip_connstr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ip_db)
		Set ip_conn = Server.CreateObject("ADODB.Connection")
		ip_conn.Open ip_connstr
		set ip_rs=ip_conn.execute("select country from address where ip1 <= "&Sip&" and ip2 >= "&Sip&"")
		if not ip_rs.Eof then
			Where = ip_rs(0)
		else
			Where = "未知地址"
		end if
		set ip_rs=nothing
		ip_conn.close
		set ip_conn=nothing
		
		sql="select CW,[Where] from CC_W where [Where] = '"&Where&"'"
		set rs2=hx.Getrs(sql,1,2)
		if rs2.eof then
			rs2.addnew
			rs2(1)=Where
		else
			rs2(0)=rs2(0) + 1
		end if
		rs2.update
		set rs2=nothing
	End Function

'处理vpage信息
Public Sub CountPage(str,vpage)
	dim rs2,sql
	sql="select "&str&",Page,vtime from CC_P where Page='"&vpage&"'"
	set rs2=server.createobject("adodb.recordset")
	rs2.open sql,conn,1,2
		if rs2.eof then
			rs2.addnew
			rs2(0)=1
			rs2(1)=vpage
		else
			rs2(0)=rs2(0) + 1
			rs2(2)=now()
		end if
		rs2.update
		rs2.close
	set rs2=nothing
End Sub
End class

Private Sub OutPut	
	dim outstr		
	'根据要求输出
	dim style
	style=Request("style")
	select case style
	case "no"	'什么都不显示
		outstr= ""
	case else	'默认显示小图标
		outstr= "<a href='"&hx.baseurl&"show.asp' target='_blank'><img src='"&hx.baseurl&"cc_icon.gif' border='0' alt='"&hx.WebName&" CuteCounter'></a>"
	end select
	response.write "document.write("& chr(34) & outstr & chr(34) &")"
End Sub
%>