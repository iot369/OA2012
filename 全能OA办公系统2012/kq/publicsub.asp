
<%
'返回新时间
function getnewtime(oldtime,addtime)
	dim hourvalue,minutevalue,newminute,newtime
	hourvalue=hour(oldtime)
	minutevalue=minute(oldtime)+addtime
	hourvalue=hourvalue+fix(minutevalue/60)
	newminute=minutevalue mod 60
	newtime=timeserial(hourvalue,newminute,0)
	getnewtime=newtime
end function
'读取用户查看所有部门的考勤资料权限
function lookkqpopedom()
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from userinf where username='"&oabusyusername&"'"
	rs.open sql,conn,1
	allow_look_all_kq_info=rs("allow_look_all_kq_info")
	conn.close
	set conn=nothing
	set rs=nothing
	lookkqpopedom=allow_look_all_kq_info
end function
'判断是迟到还是早退
'statusflag=0:表示是否迟到
'statusflag=1:表示是否早退
function verdictuserstatus(time1,time2,number,statusflag)
	dim returnstr	
	if statusflag=0 and time2>getnewtime(time1,number) then
		returnstr="迟到"
	elseif statusflag=1 and time2<getnewtime(time1,-number) then
		returnstr="早退"
	else
		returnstr=""
	end if
	verdictuserstatus=returnstr
end function
'处理上午上班时间
sub disposeamcometime()
	if rs1.eof or rs1.bof then
		getamcometime="<font color='#ee0000'>未考勤</font>"
		getamgotime="00:00:00"
		getamexplain="（上午上班<font color='#dd0000'>未到</font>！）"
		amnocomesums=amnocomesums+1
	elseif rs1("comedate")=#0:00:00# or verdictuserstatus(amcometime,rs1("comedate"),comedelaytime,0)<>"" then
		getamcometime="<font color='#dd0000'>"&rs1("comedate")&"</font>"
		getamgotime="00:00:00"
		getamexplain="（上午上班<font color='#dd0000'>迟到</font>！）"
		if rs1("explain1")<>"" then
			getamexplain=getamexplain&"<br>（迟到原因："&rs1("explain1")&"）"
		end if
		amlatesums=amlatesums+1
	else
		getamcometime=rs1("comedate")
		getamgotime="00:00:00"
		getamexplain=""
	end if
end sub
'处理上午下班时间
sub disposeamgotime()
	if rs1.eof or rs1.bof then
		getamgotime="<font color='#dd0000'>未考勤</font>"
	elseif rs1("leavedate")=#0:00:00# then
		getamgotime="<font color='#dd0000'>未考勤</font>"
		getamexplain=getamexplain&"（上午下班<font color='#dd0000'>未考勤</font>！）"
	elseif verdictuserstatus(amgotime,rs1("leavedate"),goaheadtime,1)<>"" then
		getamgotime="<font color='#dd0000'>"&rs1("leavedate")&"</font>"
		getamexplain=getamexplain&"（上午下班<font color='#dd0000'>早退</font>！）"
		if rs1("explain2")<>"" then
			getamexplain=getamexplain&"<br>（迟到原因："&rs1("explain2")&"）"
		end if
		amleaveearlysums=amleaveearlysums+1
	else
		getamgotime=rs1("leavedate")
		if rs1("explain2")<>"" then
			getamexplain=getamexplain&"（上午下班事件："&rs1("explain2")&"）"
		end if
	end if
end sub
'处理下午上班时间
sub disposepmcometime()
	if rs2.eof or rs2.bof then
		getpmcometime="<font color='#dd0000'>未考勤</font>"
		getpmgotime="00:00:00"
		getpmexplain="（下午上班<font color='#dd0000'>未到</font>！）"
		pmnocomesums=pmnocomesums+1
	elseif rs2("comedate")=#0:00:00# or verdictuserstatus(pmcometime,rs2("comedate"),comedelaytime,0)<>"" then
		getpmcometime="<font color='#dd0000'>"&rs2("comedate")&"</font>"
		getpmgotime="00:00:00"
		getpmexplain="（下午上班<font color='#dd0000'>迟到</font>！）<br>"
		if rs2("explain1")<>"" then
			getpmexplain=getpmexplain&"<br>（迟到原因："&rs2("explain1")&"）"
		end if
		pmlatesums=pmlatesums+1
	else
		getpmcometime=rs2("comedate")
		getpmgotime="00:00:00"
		getpmexplain=""
	end if
end sub
'处理下午下班时间
sub disposepmgotime()
	if rs2.eof or rs2.bof then
		getpmgotime="<font color='#dd0000'>未考勤</font>"
	elseif rs2("leavedate")=#0:00:00# then
		getpmgotime="<font color='#dd0000'>未考勤</font>"
		getpmexplain=getpmexplain&"（下午下班<font color='#dd0000'>未考勤</font>！）"
	elseif verdictuserstatus(pmgotime,rs2("leavedate"),goaheadtime,1)<>"" then
		getpmgotime="<font color='#dd0000'>"&rs2("leavedate")&"</font>"
		getpmexplain=getpmexplain&"（下午下班<font color='#dd0000'>早退</font>！）<br>"
		if rs2("explain2")<>"" then
			getpmexplain=getpmexplain&"<br>（早退原因："&rs2("explain2")&"）"
		end if
		pmleaveearlysums=pmleaveearlysums+1
	else
		getpmgotime=rs2("leavedate")
		if rs2("explain2")<>"" then
			getpmexplain=getpmexplain&"（下午下班事件："&rs2("explain2")&"）"
		end if
	end if
end sub
%>