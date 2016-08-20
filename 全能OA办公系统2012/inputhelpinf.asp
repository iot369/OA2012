<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%#@~^pAEAAA==@#@&WC(EkXxm:xD;;+kY ^KW3bn/vJWm8;/Hxm:nJb@#@&WC4!dX!/n.	l:nxM+5EdDR^KW0k+k`rGl(EdX!/n.	l:E#@#@&Wm8;/HEk+.Nn2D'.+$;+kY ^KW3rnk`EWm8!/z!/DNwDE#@#@&Gl(Edz!/+Ms+7+V{.n;!+kY mGG0kn/vEWm4;dHE/n.^+-+^E*@#@&b0,Wl(EkzEk+.xm:nxrJPD4+	P@#@&7.+kwKxd+ AMkO+vE@!km.raYPsC	o;lTn{JENl7l/1Db2YrJ@*J*@#@&7M+/aGxk+RS.rY`rhrx[GSROWa VKmCObWx 4M+W'E[0C!VDRlkwEIJ*@#@&dM+d2Kx/ hMkYcE@!J/1DrwO@*r#@#@&i.+kwGUk+RnU9@#@&+	[,kW@#@&@#@&0IoAAA==^#~@%>
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
<td>输入帮助信息</td>
</tr>
</table>
</center>

<%#@~^TAEAAA==@#@&kW,D;E/DcJkE8:bYEb{J输入J,Otx@#@&dnY,mKxU'G2x[4vEWm4;dHJ~E^KxUJBEmm^/kN/	J*@#@&k;sP{PE(	/+MOP&xYK~4+^wbxWPc4V2/K.YBtnsaYkOs~4+^21WUD+	Y#,.msE/cPr@#@&d$VP{~/$VPL~j;^?DDcDn5!+dYvEtV2dKDYEb*P'PrS,J@#@&/$VP{Pk5V,[~?$VjOM`D5E/YvE4+^wDkOVnE*#~[,E~,J@#@&k;V~x,/5V,',?5^?DD`M+$;+kYcJ4+s21WxDnxDJ#*~'Pr#r@#@&mGU	R36^ED+~d$V@#@&e2QAAA==^#~@%>
<center><br><font color=red>成功输入帮助信息！</font>
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>
<script Language="JavaScript">
 function form_check(){
   var l1=document.form1.helptitle.value.length;
   if(l1<1){window.alert("标题必须填写");document.form1.helptitle.focus();return (false);}

   var l2=document.form1.helpcontent.value.length;
   if(l2<1){window.alert("内容必须填写");document.form1.helpcontent.focus();return (false);}
                    }
</script>
<center>
<br>
<form method="post" action="inputhelpinf.asp" name="form1" onsubmit="return form_check();">
<table border="1"  cellspacing="0" cellpadding="0">
<tr>
<td>帮助标题</td><td><input type="text" name="helptitle" size=50><font color=red>*</font></td>
</tr>
<tr>
<td>帮助类别</td><td>
<select name="helpsort" size=1>
<%#@~^vgAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~CPWDKhP4+s2kWDOE@#@&./cGa+U,/$V~1W	U~8@#@&h4ksn,xWD~DkR+KW~l	N,xGY~.kR8W6@#@&HTwAAA==^#~@%>
<option value="<%=#@~^DgAAAA==.k`E4+^w/KDDE#6wQAAA==^#~@%>"><%=#@~^DgAAAA==.k`E4+^w/KDDE#6wQAAA==^#~@%></option>
<%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
</select><font color=red>*</font>&nbsp;&nbsp;&nbsp;&nbsp;(如果要增加或修改帮助类别，请<a href="edithelpsort.asp">由此进入</a>)
</td>
</tr>
<tr>
<td>帮助内容</td><td><textarea rows="9" name="helpcontent" cols="49"></textarea><font color=red>*</font></td>
</tr>
</table>
<font color=red>*</font>项必须填写&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="输入">
</form>
</center>

</body>
</html>










