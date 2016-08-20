<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="sqlstr.asp"-->
<!--#include file="conn.asp"-->
<%
set rsMsg=server.createobject("adodb.recordset")
strsql="select * from Msg where Id=" & Request("MsgId") & ""
rsMsg.open strsql,conn,1,1
%>
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>

<html>
<head>
<%
nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)
session("receiveuser")=request("receiveuser")
session("receive")=request("id")

%>

<title>在线短消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<script language=javascript>
function checkform()
{   
	if (document.form1.Content.value=="")
	{
		alert("不能发送空内容");
  	    return  false;
	}
	else  
    return true;
}
</script>
</head>
<body leftmargin="0" topmargin="5" marginwidth="0" onLoad="MM_preloadImages('images/history_on.gif','images/cancel_on.gif','images/reset_on.gif','images/submit_on.gif','images/more_on.gif','images/close_on.gif')">

<form method="post" action="sendmessage.asp" name="form1" >
<input type=hidden name=FromUserId value=<%=request.cookies("oabusyusername")%>>
  <table id=table1 width="100%" border="0" cellspacing="0" class="borderon">
    <tr> 
      <td colspan="2" align="center">
        <p align="left"><font color="#666666">&nbsp;姓名         
        <input type="text" name="ToUserId" size="11" value="<%=session("receiveuser")%>">&nbsp;  
        时间：<%=Now()%></font> </p>      
 </td> 
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
 <font color="#666666"> 
 <textarea name="Content" rows="7" cols="40"></textarea>
        
 </font>
        
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="center">
        <p align="right"><font color="#666666"><a href="#" onMouseOut="MM_swapImgRestore()" >
      <img name="Image2" border="0" src="images/history_off.gif" hspace="1" onClick="javascript:window.resizeTo(320,350);"></a> 
        <a href="#" onMouseOut="MM_swapImgRestore()"><img name="Image41" border="0" src="images/reset1_off.gif" hspace="1" onClick="MM_callJS('window.close();')"></a>     
        
        <a href="Javascript:document.form1.submit();" onMouseOut="MM_swapImgRestore()">       
        <img name="Image5" border="0" src="images/submit1_off.gif" hspace="1" onclick="return checkform();"></a>  
        
        </font>  
        
        </p>
        
      </td>   
      
    </tr>   
    <%
	set rs=server.createobject("ADODB.recordset")
session("Uid")=request.cookies("oabusyusername")
    rs.open "select * from msg where send='"&session("Uid")&"' or receive='"&session("Uid")&"' order by id desc",conn,1,1
	%>
    <tr>    
      <td colspan="2" align="center">    
        <font color="#666666">    
        <textarea name="textarea"  rows="8" cols="40">
<%
if not rs.eof then
do while not (rs.eof or rs.bof)
%>
发送者:<%=rs("send")%> 接受者:<%=rs("receive")%>
内容:<%=rs("content")%>
时间:<%=rs("dateandtime")%>
<%
rs.movenext 
loop 
end if%>

</textarea>   
        </font>   
      </td>   
    </tr>   
    <tr>    
         
    <td colspan="2" align="center"> 
    <p align="right"> 
    <font color="#666666"> 
    <a href="clearMsg.asp?id=<%=request("id")%>" onMouseOut="MM_swapImgRestore()"><img name="Image8" border="0" src="images/cancel_off.gif"></a> 
    <a href="#" onMouseOut="MM_swapImgRestore()"><img name="Image7" border="0" src="images/close0_off.gif" onClick="javascript:window.resizeTo(320,193);" hspace="10"></a></font></p> 
    </td>                   
    </tr>                  
  </table>                  
</form>                  
</body>                  
</html>
