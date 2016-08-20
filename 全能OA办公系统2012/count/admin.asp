<!-- #include file="conn.asp"-->
<!--#include file="skin.asp"-->
<!--#include file="languages.asp"-->
<%
 Response.Expires = -1  
 Response.ExpiresAbsolute = Now() - 1  
 Response.cachecontrol = "no-cache" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=Lang.item("charset")%>">
<title><%=Lang.item("g_101")%></title>
<link href="style/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
Dim TableTop,TableEnd
TableTop="<table width=500 border=0 align=center cellpadding=3 cellspacing=0 class=main_tdbg2><tr><td>"
TableEnd="</td></tr></table>"  
Dim ObjInstalled
DIm Objfso
Objfso = "Scripting.FileSystemObject"
ObjInstalled=IsObjInstalled(Objfso)

if Session(CacheName & "_Admin")<>"OK" then
	if Request.form("adminpass")="" then
		Call ShowForm
	elseif Request.form("adminpass")<>"" then
	dim AdminPass
	AdminPass = hx.execute("select AdminPass from WebInfo where ID=1")(0)
		if Request.form("adminpass")=AdminPass then
			Session(CacheName & "_Admin")="OK"
			Response.redirect "admin.asp"
		else
			Response.redirect "admin.asp"			
		end if
	end if
hx.ShowFooter
set hx=nothing
Response.end
end if

Dim Action,ErrMsg
Action=Request.QueryString("action")
select case Action
case "ExitLogin"
	call ExitLogin
case "StartCount"
	Application.Contents(CacheName & "_isStart")=1
	call WriteSuccessMsg(Lang.item("g_071"))
case "StopCount"
	Application.Contents(CacheName & "_isStart")=0
	call WriteSuccessMsg(Lang.item("g_072"))
case "ShowConfig"
	call main
	call ShowConfig
case "ClearCache"
	Call ClearCache
case "SaveConfig"
	call SaveConfig
case "CompactDB"
	call Compact
case else
	call main
end select
Dim FoundErr
if FoundErr=True then
	call WriteErrMsg(ErrMsg)
end if
hx.ShowFooter
set hx=nothing

	Sub ExitLogin
		Session(CacheName & "_Admin")=""
		Response.redirect "admin.asp"
	End Sub
	
	Sub ShowForm
		Response.write TableTop
		Response.write "<form name=form1 method=post action=''>"
		Response.write Lang.item("s_03")&"<input type=password name=adminpass>"
		Response.write "<input type=submit name=Submit value="&Lang.item("b_01")&">"				
		Response.write TableEnd
	End Sub
	
	Sub main
%>
<table width="0%" border="0" align="center" cellpadding="0" cellspacing="0" class="main_tdbg">
  <tr>
    <td>
        <p align="center"><%=Lang.item("g_102")%></p>
        
      <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr> 
          <td><li> 
              <%
	  if Application.Contents(CacheName & "_isStart")=0 then
	  Response.Write("<a href=?action=StartCount>"&Lang.item("g_103")&"</a>")
	  else
	  Response.Write("<a href=?action=StopCount>"&Lang.item("g_104")&"</a>")
	  end if	  
	  %>
          </td>
        </tr>
        <tr> 
          <td><li><a href=?action=ShowConfig><%=Lang.item("g_105")%></a> (<a href="show.asp" target="_blank"><font color="#000000"><%=Lang.item("g_128")%></font></a>)</td>
        </tr>
        <tr> 
          <td><li><a href=aspcheck.asp target="_blank"><%=Lang.item("g_106")%></a> (<a href="http://www.itlearner.com/aspcheck/" target="_blank"><%=Lang.item("g_107")%></a>)</td>
        </tr>
        <tr>
          <td><li><a href=?action=ClearCache><%=Lang.item("g_108")%></a></td>
        </tr>
        <tr> 
          <td><li><a href=?action=ExitLogin><%=Lang.item("g_109")%></a> </td>
        </tr>
<%If IsSqlDataBase = 0 Then%>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td><li>
		<%If ObjInstalled=false Then
			Response.Write Lang.item("g_110")
		else%>
        <a href=?action=CompactDB onclick="return confirm('<%=Lang.item("g_111")%>')"><%=Lang.item("g_112")%></a> 
              <%
			Response.Write Lang.item("g_113")
            Call ShowFileInfo(DB)
            Response.Write Lang.item("g_114")
            end if%>
          </td>
        </tr>
<%end if%>
        <tr> 
          <td></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%End Sub%>
<%Sub ShowConfig
dim rs
set rs=hx.execute("select * from WebInfo where ID=1")
%>
<table width="768" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" Class="main_tdbg2">
  <form name="form1" method="post" action="?Action=SaveConfig">
    <tr> 
      <td height="22" colspan="2" class="topbg"> <div align="center"><strong><%=Lang.item("g_129")%></strong></div></td>
    </tr>
    <tr> 
      <td width="400" height="25"><strong><%=Lang.item("g_001")%></strong></td>
      <td width="368"> <input name="WebName" type="text" value="<%=rs("WebName")%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr> 
      <td height="25"><strong><%=Lang.item("g_008")%></strong></td>
      <td> <input name="WebIntro" type="text" id="WebIntro" value="<%=rs("WebIntro")%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr> 
      <td height="25"><strong><%=Lang.item("g_007")%></strong> <%=Lang.item("g_115")%></td>
      <td> <input name="WebUrl" type="text" id="WebUrl" value="<%=rs("WebUrl")%>" size="40" maxlength="100"> 
      </td>
    </tr>
    <tr> 
      <td height="25"><strong><%=Lang.item("g_002")%></strong></td>
      <td> <input name="WebAdmin" type="text" id="WebAdmin" value="<%=rs("WebAdmin")%>" size="40" maxlength="20"> 
      </td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_118")%></td>
      <td> <input name="MaxPageSize" type="text" id="MaxPageSize" value="<%=MaxPageSize%>" size="40" maxlength="3"> 
      </td>
    </tr>
    <tr> 
      <td height="12"><%=Lang.item("g_119")%></td>
      <td> <input name="ExpireTime" type="text" id="ExpireTime" value="<%=ExpireTime%>" size="40" maxlength="2"> 
      </td>
    </tr>
    <tr> 
      <td height="12"><%=Lang.item("g_120")%></td>
      <td><input name="MaxRecord" type="text" id="MaxRecord" value="<%=MaxRecord%>" size="40" maxlength="3"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_121")%></td>
      <td> <select name="IsPublic" id="IsPublic">
          <option value="0" <%if rs("IsPublic")="0" then response.write " selected"%>>0</option>
          <option value="1" <%if rs("IsPublic")="1" then response.write " selected"%>>1</option>
          <option value="2" <%if rs("IsPublic")="2" then response.write " selected"%>>2</option>
        </select> </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_122")%></td>
      <td><input name="ViewPass" type="text" id="ViewPass" value="<%=rs("ViewPass")%>" size="40" maxlength="12"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_123")%></td>
      <td> <input name="AdminPass" type="text" id="AdminPass" value="<%=rs("AdminPass")%>" size="40" maxlength="12"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_124")%></td>
      <td> <input name="OnlineTime" type="text" id="OnlineTime" value="<%=OnlineTime%>" size="40" maxlength="2"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_125")%></td>
      <td><input name="TimeZone" type="text" id="TimeZone" value="<%=TimeZone%>" size="40" maxlength="2"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_126")%></td>
      <td> <select name="Language" id="Language">
          <option value="CHS" <%if Language="CHS" then response.write " selected"%>>中文简体(CHS)</option>
          <option value="CHT" <%if Language="CHT" then response.write " selected"%>>いゅ羉蔨(CHT)</option>
          <option value="ENG" <%if Language="ENG" then response.write " selected"%>>English(ENG)</option>
        </select></td>
    </tr>
    <tr> 
      <td height="12"><%=Lang.item("g_140")%></td>
      <td> <select name="Skin" id="Skin">
          <%
		dim i
		for i=0 to 4
		response.write "<option value="&i
		if int(Skin)=i then
		response.write " selected"
		end if 
		response.write ">"
		response.write Lang.item("m_"&cstr(81+i))
		response.write "</option>"
		next
		%>
        </select></td>
    </tr>
    <tr> 
      <td height="13"><%=Lang.item("g_141")%></td>
      <td><%
	  dim arr_sysmode,strg_142,sysmode1
	  arr_sysmode = split(sysmode,"|")
	  strg_142 = Lang.item("g_142")
	  sysmode1 = "<input type=checkbox name=sysmode1 value=1"
	  if int(arr_sysmode(0)) = 1 then sysmode1=sysmode1 & " checked"
	  sysmode1 = sysmode1 & ">"
	  strg_142 = replace(strg_142,"sysmode1",sysmode1)	  
	  strg_142 = replace(strg_142,"sysmode2","<input name=sysmode2 type=text value="&arr_sysmode(1)&" size=2 maxlength=2>")	  
	  strg_142 = replace(strg_142,"sysmode3","<input name=sysmode3 type=text value="&arr_sysmode(2)&" size=2 maxlength=2>")
		response.write strg_142
	  %>
         </td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_127")%></td>
      <td><input name="RecordNum" type="text" id="RecordNum" value="<%=RecordNum%>" size="40" maxlength="3"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_144")%></td>
      <td><input name="YVisitor" type="text" id="YVisitor" value="<%=YVisitor%>" size="5" maxlength="10"></td>
    </tr>
    <tr> 
      <td height="25"><%=Lang.item("g_145")%></td>
      <td><input name="YPageView" type="text" id="YPageView" value="<%=YPageView%>" size="5" maxlength="10"></td>
    </tr>
    <tr align="center"> 
      <td height="25" colspan="2"><input name="cmdSave" type="submit" id="cmdSave" value=" <%=Lang.item("b_03")%> "></td>
    </tr>
    <%
set rs=nothing
If ObjInstalled=false Then
	Response.Write "<tr><td height='40' colspan='2'><b><font color=red>"&Lang.item("g_134")&"(" & Objfso & ")! "&Lang.item("g_135")&"</font></b></td></tr>"
end if
  %>
  </form>
</table>

<%End Sub%>
</body>
</html>


<%
Sub Compact
		Response.write TableTop
		Application.Contents(CacheName & "_isStart")=0
		Response.Write CompactDB(Server.Mappath(db),false)
		Application.Contents(CacheName & "_isStart")=1		
		Response.write "<p align=center><a href=javascript:history.go(-1)> "& Lang.item("g_130")&" </a>"
		Response.write " <a href=?action=main> "& Lang.item("g_131")&" </a></p>"		
		Response.write TableEnd
End Sub
		

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
	On Error Resume Next
	Dim fso, Engine, strDBPath,JET_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = CreateObject(Objfso)
	If Err Then
		Err.Clear
		CompactDB = Lang.item("g_110") & vbCrLf
		Exit Function
	End If
	If fso.FileExists(dbPath) Then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")

		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If

		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = Nothing
		Set Engine = Nothing
		CompactDB = Lang.item("g_136") & vbCrLf
	Else
		CompactDB = Lang.item("g_137") & vbCrLf
	End If
End Function

Sub ShowFileInfo(filespec)
    Dim fs, f, s, showsize
    Set fs = Server.CreateObject(Objfso)
    Set f = fs.GetFile(server.mappath(filespec))
    s = f.size
	if s>1024*1024 then
		showsize=formatnumber(s/1024/1024,2) & "&nbsp;MB"
	elseif s>1024 then
		showsize=formatnumber(s/1024,2) & "&nbsp;KB"
	else
		showsize=s & "&nbsp;Byte" 				
	end if
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub


'检查组件是否已经安装
Function IsObjInstalled(ClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim TestObj
	Set TestObj = Server.CreateObject(ClassString)
	If 0 = Err Then IsObjInstalled = True
	Set TestObj = Nothing
	Err = 0
End Function

sub SaveConfig()

	If ObjInstalled=false Then
		'FoundErr=True
		'ErrMsg=ErrMsg & "<br><li>"&Lang.item("g_134")&"("&Objfso&")! </li>"
		'exit sub
	else
		dim sysmode1,sysmode
		if request("sysmode1")="" then
		sysmode1 = 0	
		else
		sysmode1 = request("sysmode1")	
		end if
		sysmode= sysmode1 & "|" & request("sysmode2") & "|" &  request("sysmode3")
		
		dim fso,fs
		set fso=Server.CreateObject(Objfso)
		set fs=fso.CreateTextFile(Server.mappath("config.asp"),true)
		
		fs.write "<" & "%" & vbcrlf
		fs.write "Const MaxPageSize=" & trim(request("MaxPageSize")) & "        '查看统计记录时，每页最多显示多少条记录" & vbcrlf
		fs.write "Const ExpireTime=" & trim(request("ExpireTime")) & "        '同一IP每隔多少时间后访问才继续计数，单位小时，默认为24小时" & vbcrlf
		fs.write "Const MaxRecord=" & trim(request("MaxRecord")) & "        '后台管理时显示多少条记录，默认为100条" & vbcrlf
		fs.write "Const OnlineTime=" & trim(request("OnlineTime")) & "        '在线人数截取时间，单位分钟，默认为20分钟" & vbcrlf
		fs.write "Const TimeZone=" & trim(request("TimeZone")) & "        '服务器所在时区，中国为东8区，所以默认为8" & vbcrlf
		fs.write "Const Language=" & chr(34) & trim(request("Language")) & chr(34) & "        '默认语言，默认为简体中文CHS" & vbcrlf
		fs.write "Const Skin =  " & chr(34) & request("Skin") & chr(34) & "       '	系统默认风格，可选范围0－4" & vbcrlf 
		fs.write "Const Sysmode =  " & chr(34) & sysmode & chr(34) & "       '第一个参数默认为0，日ip小于1000的设置为0；大于1000以上的设置为1，默认自动清理10天没有访问且访问数据小于5次的内容。" & vbcrlf
		fs.write "Const RecordNum=" & trim(request("RecordNum")) & "        '最后详细来访信息记录多少猹记录，默认为100条。因涉及到对数据库的操作，请登陆后台管理后修改此值，在此修改无效。" & vbcrlf  
		fs.write "Const YVisitor=" & trim(request("YVisitor")) & "        '原网站访问量" & vbcrlf  
		fs.write "Const YPageView=" & trim(request("YPageView")) & "        '原网站浏览量" & vbcrlf  
		fs.write "%" & ">"
		fs.close
		set fs=nothing
		set fso=nothing	
	end if
	
	dim rs
	set rs=hx.getrs("select * from WebInfo where ID=1",1,3)
	rs("WebName")=hx.checkstr(request("WebName"),12)
	rs("WebUrl")=hx.checkstr(request("WebUrl"),50)
	rs("WebAdmin")=hx.checkstr(request("WebAdmin"),12)
	rs("WebIntro")=hx.checkstr(request("WebIntro"),100)
	rs("IsPublic")=Cint(request("IsPublic"))
	rs("ViewPass")=hx.checkstr(request("ViewPass"),12)
	rs("AdminPass")=hx.checkstr(request("AdminPass"),12)
	rs.update
	set rs=nothing
	
					Dim RecordNum,RecordNum1,cha,i
					Dim ars,drs
					RecordNum = Request("RecordNum")
 					if RecordNum = "" or not isnumeric(RecordNum) then RecordNum=100
						RecordNum1 = hx.execute("select count(id) from CC_V")(0)
						cha = RecordNum - RecordNum1
						if cha > 0 then 
							set ars=Server.CreateObject("ADODB.Recordset")
							ars.open "CC_V",Conn,2,3
							for i = 1 to cha
								ars.addnew
							next
							ars.UpdateBatch
							ars.close
						elseif cha < 0 then
							set drs=Server.CreateObject("ADODB.Recordset")
							drs.open "select top "&abs(cha)&" * from CC_V order by dateandtime asc,id asc",conn,2,3
							for i = 1 to abs(cha)
								drs.delete
								drs.MoveNext
							next
							drs.UpdateBatch
							drs.close							
						end if

	call WriteSuccessMsg(Lang.item("g_139"))
end sub

'显示错误提示信息
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=TableTop & Lang.item("g_133")
	strErr=strErr & ErrMsg
	strErr=strErr & "<p align=center><a href=javascript:history.go(-1)>"&Lang.item("g_130")&"</a>"
	strErr=strErr & " <a href=?action=main>"&Lang.item("g_131")&"</a></p>"		
	strErr=strErr & TableEnd
	response.write strErr
end sub


'显示成功提示信息
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=TableTop & Lang.item("g_132") & SuccessMsg
	strSuccess=strSuccess & "<p align=center><a href=javascript:history.go(-1)>"&Lang.item("g_130")&"</a>"
	strSuccess=strSuccess & " <a href=?action=main>"&Lang.item("g_131")&"</a></p>"		
	strSuccess=strSuccess & TableEnd
	response.write strSuccess
end sub

sub ClearCache
	application.Contents.RemoveAll()
	Session.Contents.RemoveAll()
	Session(CacheName & "_Admin")="OK"
	WriteSuccessMsg(Lang.item("g_138"))
end sub


%>
