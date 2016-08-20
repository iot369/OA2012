<%@ LANGUAGE = VBScript %>
<%Response.Expires=0%>
<script language="JScript" runat="server">
function bm(inputstr)
{
return unescape(inputstr);
}
</script>
<%
idinfo=Request.QueryString("id")
infostr=bm(Request.QueryString("infostr"))
yhm=Request.QueryString("yhm")
sitemc=Request.QueryString("sitemc")
url=Request.QueryString("url")
%>
 <html><head> 
 <title>即时通信工具</title> 
 <meta http-equiv="Content-Type" content="text/html;charset=gb2312"> 
 <meta http-equiv="Content-Language" content="zh-CN"> 
 <link rel="stylesheet" type="text/css" href="style.css">
 <SCRIPT LANGUAGE=javascript FOR=document EVENT=onkeyup> 
 <!-- 
 document_onkeyup(); 
 //--> 
 </SCRIPT> 
 <script language=javascript> 
 var sendfunction; 
 sendfunction=0; 
 window.focus();
 function document_onkeyup() { 
 if ((event.altKey==true && event.keyCode==83) || (event.ctrlKey==true && event.keyCode==13)) 
 	{	 
 		change_function(); 
 	} 
 else if (event.altKey==true && event.keyCode==67) 
 	window.close(); 
 else if (event.altKey==true && event.keyCode==68) 
 	listhistory();
 } 
  
 function change_function() 
 { 
 	if (sendfunction==0) 
 		{ 
 			form.send.value="发 送[S]"; 
 			form.info.readOnly=false; 
 			form.info.style.background="#ffffff"; 
 			form.info.value="";	 
 			form.info.focus();
 			sendfunction=1;
 		} 
 	else 
 		{ 
 			if (form.info.value=="") 
 				alert("不能发送空信息！"); 
 			else if (form.info.value.length>100) 
 				alert("发送信息太长，不能超过100个汉字！"); 
 			else if (form.info.value.indexOf("$")!=-1 || form.info.value.indexOf("|")!=-1 ) 
 				alert("信息中不能含有$,|字符！"); 
 			else 
 				{ 
					opener.parent.history.value=form.backinfo.value+escape(form.info.value)+"|"+opener.parent.history.value;
 					form.send.disabled=true; 
 					window.moveTo(1100,700); 
 					document.form.submit(); 
 				} 
 		} 
 	 
 } 
 
 var winflag;
winflag=0;
function listhistory()
{
if (winflag==0)
{
	window.resizeTo(365,473);
	winflag=1;
}
else
{
	window.resizeTo(365,273);
	winflag=0;
}
disp_history();
}

function disp_history()
{
var i,infodim,infstr,number,infostr,re;
infostr=opener.parent.history.value;
number=<%=idinfo%>;
window("historywin").document.open();
window("historywin").document.write("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\"><style type=\"text/css\">");
window("historywin").document.write("<!--div{  font-family: \"宋体\"; font-size: 9pt;}--></style></head><body>");
if (infostr.length>0)
{
	infodim=infostr.split("|");
	for (i in infodim)
	{
		if (infodim[i]!="")
		{
			infstr=infodim[i].split("$");
			if (infstr[0]==number || infstr[1]==number)
			{
				info=unescape(infstr[5]);
				re=/</g;
				info=info.replace(re,"&lt;");
				re=/>/g;
				info=info.replace(re,"&gt;");
				re=/\r\n/g;
				info=info.replace(re,"<br>");
				window("historywin").document.write("<div align=\"left\"><font color=\"#0000ee\">"+infstr[2]+"---"+infstr[4]+"</font><br>"+info+"</div>");
			}
		}
	}
}
window("historywin").document.write("</body></html>");
window("historywin").document.close();
}
 </script> 
 </head> 
 <body bgcolor="#ffffff" leftmargin="0" topmargin="0"> 
 <bgsound src="ri.wav" loop="1">
 	 <input id="winnumber" name="winnumber" type="hidden" value="">
 <form method="post" name="form" action="writesendinfo.asp"> 
 	 <input id="backinfo" name="backinfo" type="hidden" value="">
 <table border="0" cellpadding="0" cellspacing="0" width="354" bgcolor="#E7E7E7"> 
   <tr> 
    <td><img src="../qqpic/info/spacer.gif" width="2" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="77" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="85" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="154" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="17" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="17" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="2" height="1" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="1" border="0"></td> 
   </tr> 
  
   <tr> 
    <td colspan="7"><img name="porm_r1_c1" src="../qqpic/info/porm_r1_c1.gif" width="354" height="2" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="2" border="0"></td> 
   </tr> 
   <tr> 
    <td rowspan="5"><img name="porm_r2_c1" src="../qqpic/info/porm_r2_c1.gif" width="2" height="240" border="0"></td> 
    <td bgcolor="#BDBEBD"><img name="porm_r2_c2" src="../qqpic/info/title.gif" width="77" height="16" border="0"></td> 
    <td colspan="2" width="239" height="16" bgcolor="#BDBEBD"> 
    </td> 
    <td width="17" height="16" bgcolor="#BDBEBD"> 
    </td> 
    <td width="17" height="16" bgcolor="#BDBEBD"> 
     <p align="center"><img border="0" src="../qqpic/Close.gif" width="11" height="11"></p> 
    </td> 
    <td rowspan="5"><img name="porm_r2_c7" src="../qqpic/info/porm_r2_c7.gif" width="2" height="240" border="0"></td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="16" border="0"></td> 
   </tr> 
   <tr> 
    <td colspan="5" width="350" height="26"> 
     <table border="1" cellpadding="0" cellspacing="0" width="100%" height="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
       <tr> 
         <td width="100%" id="listinfo" name="listinfo"> 
           <p align="center"> 
           <%=""&server.HTMLEncode(yhm)&"在<a href='#' onclick=""window.open('"&url&"','','toolbar=yes,scrollbars=yes,status=yes,resizable=1,menubar=yes,location=yes,width=600,height=400','name=main8315')"&chr(34)&" title="&url&">"&server.htmlencode(sitemc)&"</a>网站给您发来信息"%>
 		   </p> 
         </td> 
       </tr> 
     </table> 
    </td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="26" border="0"></td> 
   </tr> 
   <tr> 
    <td colspan="2" width="162" height="40"> 
     <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" height="100%" bordercolordark="#FFFFFF"> 
       <tr> 
         <td width="100%" valign="middle" height="40"> 
           <p align="left">请输入信息<br> 
           请按Alt+S键发送信息</p> 
         </td> 
       </tr> 
     </table> 
    </td> 
    <td colspan="3" width="188" height="40" valign="middle"> 
	<img src="../qqpic/logo.gif" border="0" width="180" height="40">
	</td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="40" border="0"></td> 
   </tr> 
   <tr> 
    <td colspan="5" width="350" height="131" valign="top">
    <p align="left"><textarea rows="5" name="info" cols="42" readonly style="height:131;width=350;font-family: 宋体; font-size: 9pt;background:#eeeeee"><%=infostr%></textarea></p> 
    </td> 
    <td><img src="../qqpic/info/spacer.gif" width="1" height="131" border="0"></td> 
   </tr> 
   <tr> 
    <td colspan="5" width="350" height="27">&nbsp;<input type="button" onclick="listhistory()" value="对话记录[D]" style="FONT-FAMILY: 宋体; FONT-SIZE: 9pt; HEIGHT: 20px" id=button1 name=button1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <input type="button" value="回 复[S]" name="send" onclick="change_function()" style="FONT-FAMILY: 宋体; FONT-SIZE: 9pt; HEIGHT: 20">&nbsp;&nbsp; <input type="button" value="关 闭[C]" name="close" onclick="window.close();" style="FONT-FAMILY: 宋体; FONT-SIZE: 9pt; HEIGHT: 20"></td>  
    <td><img src="../qqpic/info/spacer.gif" width="1" height="27" border="0"></td>  
   </tr>  
 </table> 
<IFRAME name="historywin" frameBorder=1 scrolling="yes" marginwidth=1 marginheight=1 height="200" src="" width="355" bgcolor="#8482C6"></IFRAME>   
 </form>  
<script language="javascript">
form.backinfo.value=<%=idinfo%>+"$"+opener.parent.usersessionid.value+"$"+opener.parent.sitename.value+"$"+opener.parent.siteurl.value+"$"+opener.parent.username.value+"$"; 
</script>
 </body>  
 </html>  
