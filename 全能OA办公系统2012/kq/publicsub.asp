
<%
'������ʱ��
function getnewtime(oldtime,addtime)
	dim hourvalue,minutevalue,newminute,newtime
	hourvalue=hour(oldtime)
	minutevalue=minute(oldtime)+addtime
	hourvalue=hourvalue+fix(minutevalue/60)
	newminute=minutevalue mod 60
	newtime=timeserial(hourvalue,newminute,0)
	getnewtime=newtime
end function
'��ȡ�û��鿴���в��ŵĿ�������Ȩ��
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
'�ж��ǳٵ���������
'statusflag=0:��ʾ�Ƿ�ٵ�
'statusflag=1:��ʾ�Ƿ�����
function verdictuserstatus(time1,time2,number,statusflag)
	dim returnstr	
	if statusflag=0 and time2>getnewtime(time1,number) then
		returnstr="�ٵ�"
	elseif statusflag=1 and time2<getnewtime(time1,-number) then
		returnstr="����"
	else
		returnstr=""
	end if
	verdictuserstatus=returnstr
end function
'���������ϰ�ʱ��
sub disposeamcometime()
	if rs1.eof or rs1.bof then
		getamcometime="<font color='#ee0000'>δ����</font>"
		getamgotime="00:00:00"
		getamexplain="�������ϰ�<font color='#dd0000'>δ��</font>����"
		amnocomesums=amnocomesums+1
	elseif rs1("comedate")=#0:00:00# or verdictuserstatus(amcometime,rs1("comedate"),comedelaytime,0)<>"" then
		getamcometime="<font color='#dd0000'>"&rs1("comedate")&"</font>"
		getamgotime="00:00:00"
		getamexplain="�������ϰ�<font color='#dd0000'>�ٵ�</font>����"
		if rs1("explain1")<>"" then
			getamexplain=getamexplain&"<br>���ٵ�ԭ��"&rs1("explain1")&"��"
		end if
		amlatesums=amlatesums+1
	else
		getamcometime=rs1("comedate")
		getamgotime="00:00:00"
		getamexplain=""
	end if
end sub
'���������°�ʱ��
sub disposeamgotime()
	if rs1.eof or rs1.bof then
		getamgotime="<font color='#dd0000'>δ����</font>"
	elseif rs1("leavedate")=#0:00:00# then
		getamgotime="<font color='#dd0000'>δ����</font>"
		getamexplain=getamexplain&"�������°�<font color='#dd0000'>δ����</font>����"
	elseif verdictuserstatus(amgotime,rs1("leavedate"),goaheadtime,1)<>"" then
		getamgotime="<font color='#dd0000'>"&rs1("leavedate")&"</font>"
		getamexplain=getamexplain&"�������°�<font color='#dd0000'>����</font>����"
		if rs1("explain2")<>"" then
			getamexplain=getamexplain&"<br>���ٵ�ԭ��"&rs1("explain2")&"��"
		end if
		amleaveearlysums=amleaveearlysums+1
	else
		getamgotime=rs1("leavedate")
		if rs1("explain2")<>"" then
			getamexplain=getamexplain&"�������°��¼���"&rs1("explain2")&"��"
		end if
	end if
end sub
'���������ϰ�ʱ��
sub disposepmcometime()
	if rs2.eof or rs2.bof then
		getpmcometime="<font color='#dd0000'>δ����</font>"
		getpmgotime="00:00:00"
		getpmexplain="�������ϰ�<font color='#dd0000'>δ��</font>����"
		pmnocomesums=pmnocomesums+1
	elseif rs2("comedate")=#0:00:00# or verdictuserstatus(pmcometime,rs2("comedate"),comedelaytime,0)<>"" then
		getpmcometime="<font color='#dd0000'>"&rs2("comedate")&"</font>"
		getpmgotime="00:00:00"
		getpmexplain="�������ϰ�<font color='#dd0000'>�ٵ�</font>����<br>"
		if rs2("explain1")<>"" then
			getpmexplain=getpmexplain&"<br>���ٵ�ԭ��"&rs2("explain1")&"��"
		end if
		pmlatesums=pmlatesums+1
	else
		getpmcometime=rs2("comedate")
		getpmgotime="00:00:00"
		getpmexplain=""
	end if
end sub
'���������°�ʱ��
sub disposepmgotime()
	if rs2.eof or rs2.bof then
		getpmgotime="<font color='#dd0000'>δ����</font>"
	elseif rs2("leavedate")=#0:00:00# then
		getpmgotime="<font color='#dd0000'>δ����</font>"
		getpmexplain=getpmexplain&"�������°�<font color='#dd0000'>δ����</font>����"
	elseif verdictuserstatus(pmgotime,rs2("leavedate"),goaheadtime,1)<>"" then
		getpmgotime="<font color='#dd0000'>"&rs2("leavedate")&"</font>"
		getpmexplain=getpmexplain&"�������°�<font color='#dd0000'>����</font>����<br>"
		if rs2("explain2")<>"" then
			getpmexplain=getpmexplain&"<br>������ԭ��"&rs2("explain2")&"��"
		end if
		pmleaveearlysums=pmleaveearlysums+1
	else
		getpmgotime=rs2("leavedate")
		if rs2("explain2")<>"" then
			getpmexplain=getpmexplain&"�������°��¼���"&rs2("explain2")&"��"
		end if
	end if
end sub
%>