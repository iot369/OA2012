<%

'�½�����վ�����
sub create_online_site()
	dim onlinesite()
	redim onlinesite(0)
	Application.Lock 
	application("onlinesite")=onlinesite
	Application.UnLock 
end sub

'�½���ǰվ�������û�����
sub create_online_user(site_id)
	dim onlineuser()
	redim onlineuser(0)
	Application.Lock 
	application("onlineuser"&site_id)=onlineuser
	Application.UnLock 
end sub

'��������վ��������Ƿ����и�վ��
'����0--������û�и�վ��
'����1--�������и�վ��
function find_online_site(site_id)
	dim i,dimsums,findok,sitestr
	findok=-1
	Application.Lock
	onlinesite=application("onlinesite")
	dimsums=ubound(onlinesite)
	for i=0 to dimsums
		siteinfo=onlinesite(i)
		sitestr=left(siteinfo,instr(siteinfo,"$"))
		if sitestr=cstr(site_id)&"$" then
			findok=i
			exit for
		end if
	next
	Application.UnLock
	find_online_site=findok
end function

'д������վ�����
function write_online_site(id,mc,lx,url,jj)
	dim siteinfo,dimsums,filename,fs,fpoint
	Application.Lock
	siteinfo=id&"$"&url&"$"&mc
	onlinesite=application("onlinesite")
	dimsums=ubound(onlinesite)
	redim preserve onlinesite(dimsums+1)
	onlinesite(dimsums+1)=siteinfo
	application("onlinesite")=onlinesite

	Application.UnLock
	write_online_site=dimsums+1
end function

'���ҵ�ǰ�û��Ƿ�����
function find_online_user(site_id)
	dim i,dimsums,siteinfo,findok
	findok=0
	Application.Lock
	onlineuser=application("onlineuser"&site_id)
	dimsums=ubound(onlineuser)
	for i=0 to dimsums
		siteinfo=onlineuser(i)
		if siteinfo<>"" then
			if instr(siteinfo,session.SessionID)>0 then
				findok=1
				exit for
			end if
		end if
	next
	Application.UnLock
	find_online_user=findok
end function

'д�������û�����
sub write_online_user(site_id,faceid)
	dim userinfo
	Application.Lock
	oabusyuserdept=request.cookies("oabusyuserdept")
	oabusyuserlevel=request.cookies("oabusyuserlevel")
	if session("manager")="1" then
	userinfo=session.SessionID&"$"&session("username")&"$"&"1$("&oabusyuserdept&"��"&oabusyuserlevel&")"&now()&"$"&now()&"$"&faceid
	elseif session("manager")="2" then	userinfo=session.SessionID&"$"&session("username")&"$"&"2$("&oabusyuserdept&"��"&oabusyuserlevel&")"&now()&"$"&faceid
	else
	userinfo=session.SessionID&"$"&session("username")&"$"&"0$("&oabusyuserdept&"��"&oabusyuserlevel&")"&now()&"$"&now()&"$"&faceid
	end if
	onlineuser=application("onlineuser"&site_id)
	dimsums=ubound(onlineuser)
	redim preserve onlineuser(dimsums+1)
	onlineuser(dimsums+1)=userinfo
	application("onlineuser"&site_id)=onlineuser
	Application.UnLock
end sub

function getsitename(site_id)
	dim infostr,filename,fs,fpoint
	on error resume next
	infostr=""
	filename=server.mappath("/")&"\qq\siteinfo\"&site_id&".txt"
	set fs=createobject("scripting.filesystemobject")
	if fs.fileexists(filename) then	
		set fpoint=fs.opentextfile(filename,1,true)
		fpoint.skipline
		infostr=fpoint.readline
		fpoint.close
	end if
	set fs=nothing
	Application.UnLock
	getsitename=infostr
end function

'ɾ�������û�
sub lost_user()
	dim delflag,dimsums,onlinesums,num
	Application.Lock
	onlineuser=application("onlineuser"&session("siteid"))
	dimsums=ubound(onlineuser)
	onlinesums=dimsums
	num=0
	for i=0 to dimsums
		delflag=0
		if onlineuser(i)="" then
			delflag=1
		end if
		if delflag=0 then
			if num<i then
				onlineuser(num)=onlineuser(i)
			end if
			num=num+1
		else
			if num<i then
				onlineuser(num)=onlineuser(i)
			end if
			onlinesums=onlinesums-1
		end if
	next
	redim preserve onlineuser(onlinesums)
	application("onlineuser"&session("siteid"))=onlineuser
	Application.UnLock
end sub

'�õ�7�����ֳ��Ȼ�14����ĸ���ȵ��ַ���
Function GetNewStr(InputStr)
	dim i,number,newstr,substr
	number=0
	newstr=""
	for i=1 to len(InputStr)
		substr=mid(InputStr,i,1)
		if asc(substr)<0 then
			number=number+2
		else
			number=number+1
		end if
		if number<=12 then
			newstr=newstr&substr
		else
			newstr=newstr&"..."
			exit for
		end if
	next
	GetNewStr=newstr
End Function
'������վ����
Function YesSite(SiteName)
	Dim NoName(6),i,Yes
	NoName(0)="����"
	NoName(1)="ɫ"
	NoName(2)="��"
	NoName(3)="ͬ־"
	NoName(4)="��Ů"
	NoName(5)="��ü"
	NoName(6)="����"
	Yes=0
	If SiteName<>"" then
		For i=0 To 6
			If Instr(SiteName,NoName(i))>0 then
				Yes=1
				Exit For
			End If
		Next
	Else
		Yes=1
	End If
	YesSite=Yes
End Function
%>
