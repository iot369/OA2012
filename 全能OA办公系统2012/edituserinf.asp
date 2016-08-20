<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%#@~^xAEAAA==@#@&WC(EkXxm:xD;;+kY ^KW3bn/vJWm8;/Hxm:nJb@#@&WC4!dX!/n.	l:nxM+5EdDR^KW0k+k`rGl(EdX!/n.	l:E#@#@&Wm8;/HEk+.Nn2D'.+$;+kY ^KW3rnk`EWm8!/z!/DNwDE#@#@&Gl(Edz!/+Ms+7+V{.n;!+kY mGG0kn/vEWm4;dHE/n.^+-+^E*@#@&b0,Wl(EkzEk+.xm:nxrJPD4+	P@#@&7.+kwKxd+ AMkO+vE@!km.raYPsC	o;lTn{JENl7l/1Db2YrJ@*J*@#@&7M+/aGxk+RS.rY`rhrx[GSROWa VKmCObWx 4M+W'E[0C!VDRlkwEIJ*@#@&dM+d2Kx/ hMkYcE@!J/1DrwO@*r#@#@&i.+kwGUk+RnU9@#@&+	[,kW@#@&@#@&Ek+MUls+xD;;nkY`r;/DxmhnJ*@#@&gpUAAA==^#~@%>
<%#@~^BgIAAA==@#@&/;(P!/+Mk	W`4Dn0*@#@&r6PD5E/YvEdE(:bYE#xE更改rPOtU@#@&wCdkhW.[{Dn;!nkYcrwm//SWM[J*@#@&xm:nxM+;!n/D`J	Ch+r#@#@&[+2O{Dn;!n/D`E[wYEb@#@&;/.^+-V{D+$EdYvJ;/Dsn7+Vrb@#@&/+D~^W	x{W2+U[(`EWm8EkXESrmWUUr~El1^/d9/	J#@#@&knY,Dd'k+.-DR1.+mY+K8%+1YvJCNG[(R.+1GD9/nOr#@#@&d$V~',E!w[mYPEk+Mrx6Pd+DPE@#@&/;^~',/;^~'Prwm/dhG.9'EPL~?$VjOM`wCdkhGD9b,[~r~,J@#@&/$sP{Pd;^P'~rxlsn'rP[,j5VUYM`Ulhn*P'PrSEk+.[wYxE,[~/$skY.vNwY*PL~JBEd+MVn-V'r~[,/;^dODvEk+.Vn-VbPL~J,h4nM+P;dDUlsn{J~LPk;VkYMcEk+.xm:nb@#@&mKUxc26^;YPk;s@#@&ZqEAAA==^#~@%>
<br><br>
<font color=red>用户资料维护成功！</font>
<br><br>
<form mothed=post action="usercontrol.asp">
<input type="submit" name="submit" value="返回">
</form>
<%#@~^EAAAAA==@#@&+sk+@#@&@#@&7gEAAA==^#~@%>

<script Language="JavaScript">
 function maxlength(str,minl,maxl) {
    if(str.length <= maxl && str.length >= minl){return true;}else{return false;}
                                    }

 function form_check(){

   var l2=maxlength(document.form2.password.value,1,20);
   if(!l2){window.alert("密码的长度大于1位小于20位");document.form2.password.focus();return (false);}

   var a1=document.form2.password.value;
   var a2=document.form2.repassword.value;
   if(a1!=a2){window.alert("两次输入的密码应相同");document.form2.repassword.focus();return (false);}

   var l3=maxlength(document.form2.name.value,1,20);
   if(!l3){window.alert("姓名的长度大于1位小于20位");document.form2.name.focus();return (false);}

                    }

</script>


<%#@~^vQAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~CPWDKhP!/n.bx0~A4+.+,;k+.	ls+'rPL~/$VdYM`;dDxmh+*@#@&Md Wa+	Pd;sS1WUxBq@#@&JD0AAA==^#~@%>

<form action="<%=#@~^BAAAAA==4M+WpQEAAA==^#~@%>" method=post name="form2" onsubmit="return form_check();">
  <table border=1 cellspacing="0" cellpadding="5">
    <tr>
<td>
用&nbsp;户&nbsp;名：<%=#@~^DgAAAA==.k`E!/Dxm:E#2gQAAA==^#~@%><input type="hidden" name="username" size=20 value="<%=#@~^DgAAAA==.k`E!/Dxm:E#2gQAAA==^#~@%>">
</td>
</tr>
<tr>
<td>
密&nbsp;&nbsp;&nbsp;&nbsp;码：<input type="password" name="password" size=20 value="<%=#@~^DgAAAA==.k`Ealk/hKD9E#7QQAAA==^#~@%>">
</td>
</tr>
<tr>
<td>
密码确认：<input type="password" name="repassword" size=20 value="<%=#@~^DgAAAA==.k`Ealk/hKD9E#7QQAAA==^#~@%>">
</td>
</tr>
<tr>
<td>
姓&nbsp;&nbsp;&nbsp;&nbsp;名：<input type="text" name="name" size=20 value="<%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%>">
</td>
</tr>
<tr>
<td>
部&nbsp;&nbsp;&nbsp;&nbsp;门：
<select name="dept" size=1>
<%#@~^vgAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d8'/n.7+.R1.lOW(L+1YvEl9W[4cDn^KDNknYr#@#@&d5V{Jk+s+^O,e~0MG:,Nn2DJ@#@&.kF Wan	Pd$VBmW	xBq@#@&h4k^+~UKYPMdFc+W6~Cx9P	WOP.d8R8W6@#@&HTsAAA==^#~@%>
<option value="<%=#@~^CwAAAA==.kFcrNwYr#WAMAAA==^#~@%>"<%=#@~^JAAAAA==dVn1YN`M/vEEk+.NwOE*~Dkq`rN+aOE#*BAwAAA==^#~@%>><%=#@~^CwAAAA==.kFcrNwYr#WAMAAA==^#~@%></option>
<%#@~^HAAAAA==@#@&Dd8RsW\xaY@#@&A+	N@#@&rQYAAA==^#~@%>
</select>
</td>
</tr>
<tr>
<td>
职&nbsp;&nbsp;&nbsp;&nbsp;位：
<select name="userlevel" size=1>
<%#@~^wwAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.dy'/n.7+.R1.lOW(L+1YvEl9W[4cDn^KDNknYr#@#@&d5V{Jk+s+^O,e~0MG:,EdnMV+-n^J@#@&MdyRGa+	P/$VB^W	xSF@#@&A4bV+,UWDPDk+ +K0,lUN~UKY~Dk+R(WW@#@&Sz0AAA==^#~@%>
<option value="<%=#@~^EAAAAA==.k crEk+D^+7nVr#gwUAAA==^#~@%>"<%=#@~^KgAAAA==dVn1YN`M/vEEk+.V\nsr#~Md vJEkn.V\VE#bmg4AAA==^#~@%>><%=#@~^EAAAAA==.k crEk+D^+7nVr#gwUAAA==^#~@%></option>
<%#@~^HAAAAA==@#@&DdyRsW\xaY@#@&A+	N@#@&rgYAAA==^#~@%>
</select>
</td>
</tr>
<tr>
<td align=center>
<input type="submit" name="submit" value="更改">&nbsp;&nbsp;<input type="button" value="返回" onclick="window.location.href='usercontrol.asp'">
</td>
</tr>
</table>
</form>
<%#@~^GQAAAA==@#@&+U9Pb0@#@&+	[PkE8@#@&DAUAAA==^#~@%>



<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>伴江行办公系统</title>
</head>
<body  topmargin="5" leftmargin="5">

<center>
<table>
<tr>
<td>编辑用户资料</td>
</tr>
</table>
</center>

<center>
<br>
<%#@~^IQAAAA==~1ls^P!/+Mk	W`r+[kDEdnMkx6 lkwJ*~oQsAAA==^#~@%>
</center>
<%#@~^CAAAAA==@#@&@#@&LgAAAA==^#~@%>

</body>
</html>