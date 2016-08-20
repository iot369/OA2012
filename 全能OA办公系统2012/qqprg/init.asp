<%Response.Expires=0%>
var openwindow=0;
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername<>"" then 
	session("username")=oabusyusername
	session("siteid")=1
%>
//if (window.name!="main8315")
//{
	//openwinflag=window.open('','hrz8315','fullscreen=1,toolbar=no,scrollbars=no,resizable=0,menubar=no',name='main8315');
	//openwinflag.resizeTo(135,405);
	//openwinflag.moveTo(200,163);
	//openwinflag.focus();
	//openwinflag.location.href="/qqprg/main.asp";
//}
	openwinflag=window.open('','hrz8315','fullscreen=1,toolbar=no,scrollbars=no,resizable=0,menubar=no',name='main8315');
	openwinflag.resizeTo(112,360);
	openwinflag.moveTo(700,363);
	openwinflag.focus();
	openwinflag.location.href="qqprg/kuaig.asp";
<%
end if
%>