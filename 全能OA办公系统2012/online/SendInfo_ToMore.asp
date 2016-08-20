<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Conn.asp"-->
<!--#include file="Function/Function.asp"-->
<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=60,width=600,height=500");
 return false;
 
}
</script>


<html>

<head>

<title>在线用户</title>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="styles.css" -->

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
//-->
</script>

</head>

<script language="JavaScript">
<!--
 function TB_animateWindow(windowWidth,windowHeight,targetWidth,targetHeight,widthMod,heightMod,fullScreen)
 {
   // www.timbooker.com
   // www.saltstonemedia.co.uk

	if (fullScreen)
	 {
		targetWidth = screen.availWidth;
		targetHeight = screen.availHeight;
	 }

	if (windowWidth < targetWidth) windowWidth += widthMod;
	if (windowHeight < targetHeight) windowHeight += heightMod;

	windowLeft = (screen.availWidth / 2) - (windowWidth / 2);
	windowTop = (screen.availHeight / 2) - (windowHeight / 2);

	window.resizeTo(windowWidth,windowHeight);
	window.moveTo(windowLeft,windowTop);

	if (windowWidth < targetWidth || windowHeight < targetHeight)
		setTimeout('TB_animateWindow(' + windowWidth + ', ' + windowHeight + ', ' + targetWidth + ', ' + targetHeight + ', ' + widthMod + ', ' + heightMod + ', ' + fullScreen + ');',10);
 }
//-->
</script>
<script language="JavaScript">
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;  
    }
  }
</script>


<body oncontextmenu=self.event.returnValue=false bgcolor="#276DB2" leftmargin="0" topmargin="5" marginwidth="0" marginheight="0"  onLoad="MM_preloadImages('images/confirm_on.gif')">
<form method="post" >
<div align="center"><font color="#FFFFFF">
<input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)'>
全选　在线用户</font>
<br>
</div>
<table width="90%" border="1" cellspacing="1" cellpadding="2" align="center" bgcolor="#FFFFEC" bordercolorlight="#cccccc" height="50%">
  <tr valign="top"> 
    
      <td colspan="2"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
<%
call OnlineUser()
    '*****************************
    user=split(application("OfficeOnlineUser"),",") '获得在线用户列表
    for i=1 to ubound(user) '获得非空的在线用户列表
%>         
          <tr> 
            <td width="4%"> 
              <input type="checkbox" name="chk<%=i-1%>"  value="<%=user(i)%>">    
            </td>
            <td width="96%"><%=GetUserName(user(i))%></td>
          </tr>
<%next%>
        </table>
      </td> 
  </tr>
</table>
<div align="center"><a href="Javascript:adduser(<%=i-1%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/confirm_on.gif',1)"><img name="Image1" border="0" src="images/confirm_off.gif"></a> 
</div></form>
</body>
</html>
<script language="javascript">
function adduser(usercount)
{
	var userlist;
	userlist=";";
	for(i=0;i<usercount;i++)
	{
		chk = "chk" + i;   
		chk = document.all(chk);   	   
		if(chk.checked)
		{
			userlist=userlist+chk.value+";";
		}	
	}
	if(userlist!="")
		window.opener.form1.receiveuser.value=userlist;
		window.opener.form1.ToUserId.value=userlist;
	window.close();
}
</script>
