<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/checked.asp"-->
<%#@~^pAEAAA==@#@&WC(EkXxm:xD;;+kY ^KW3bn/vJWm8;/Hxm:nJb@#@&WC4!dX!/n.	l:nxM+5EdDR^KW0k+k`rGl(EdX!/n.	l:E#@#@&Wm8;/HEk+.Nn2D'.+$;+kY ^KW3rnk`EWm8!/z!/DNwDE#@#@&Gl(Edz!/+Ms+7+V{.n;!+kY mGG0kn/vEWm4;dHE/n.^+-+^E*@#@&b0,Wl(EkzEk+.xm:nxrJPD4+	P@#@&7.+kwKxd+ AMkO+vE@!km.raYPsC	o;lTn{JENl7l/1Db2YrJ@*J*@#@&7M+/aGxk+RS.rY`rhrx[GSROWa VKmCObWx 4M+W'E[0C!VDRlkwEIJ*@#@&dM+d2Kx/ hMkYcE@!J/1DrwO@*r#@#@&i.+kwGUk+RnU9@#@&+	[,kW@#@&@#@&0IoAAA==^#~@%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统.边缘特别版</title>
</head>
<body  topmargin="5" leftmargin="5">

<center>
<%#@~^tAAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~CPWDKhP4+s2bx0~A4+.+,r9'E,[,D+$EdYvJrNr#@#@&M/RK2+	P/$sSmKx	~q@#@&jDgAAA==^#~@%>
<table>
<tr>
<td>编辑<%=#@~^DwAAAA==.k`E4+^wYbY^nJ*RQUAAA==^#~@%>帮助信息&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%#@~^JgAAAA==@#@&kW,D;E/DcJkE8:bYEb{J删除J,Otx@#@&nwkAAA==^#~@%>
<form method="post" action="helpinf.asp">
<%#@~^DAAAAA==@#@&+sk+@#@&1wEAAA==^#~@%>
<form method="post" action="showhelpinf.asp">
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>
<td><input type="submit" value="返回"><input type="hidden" name="id" value=<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>><input type="hidden" name="helpsort" value="<%=#@~^DgAAAA==.k`E4+^w/KDDE#6wQAAA==^#~@%>"></td></form></td>
</tr>
</table>
</center>


<center>
<%#@~^MQEAAA==@#@&kW,D;E/DcJkE8:bYEb{J修改J,Otx@#@&dnY,mKxU'G2x[4vEWm4;dHJ~E^KxUJBEmm^/kN/	J*@#@&k;sP{PE;aNlDnP4+VarU0,/Y~tnsa/GDDxJ,[~d$V/O.vDn;!nkYcrtVwkWMOJ*#~[,JS4VwDrY^+'r~'Pk;^/ODc.;;+kO`rtnsaYkOsJb#,',JS4+^wmKxDnxD'EPLPd5^/YMcD;EdO`rtV2mGUD+UYrb#,[~E,htn.PrN{E,[~M+$E+kYvEk9Jb@#@&mGU	R2Xnm!Y+,d5V@#@&YmAAAA==^#~@%>
<br><br><font color=red >成功修改帮助信息！</font>
<%#@~^qwAAAA==@#@&+sk+@#@&k6PMn;!+dYvJd;(:kDE#{J删除J,O4+	@#@&/nY~^KxU'K2+	N8crWl8;kXE~r^KxUr~rlm1+kdNkxE#@#@&d5^P',ENV+Dn~0MWsP4+s2bxWPS4+M+~r9'J~',Dn;!nkYcrk9J#@#@&1Gx	R36m;OP/$s@#@&kzIAAA==^#~@%>
<br><br><font color=red >成功删除帮助信息！</font>
<%#@~^DAAAAA==@#@&+sk+@#@&1wEAAA==^#~@%>
<br>
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
<form method="post" action="edithelpinf.asp" name="form1" onsubmit="return form_check();">
<table border="1"  cellspacing="0" cellpadding="0">
<tr>
<td>帮助标题</td><td><input type="text" name="helptitle" size=50 value="<%=#@~^DwAAAA==.k`E4+^wYbY^nJ*RQUAAA==^#~@%>"><font color=red>*</font></td>
</tr>
<tr>
<td>帮助类别</td><td>
<select name="helpsort" size=1>
<%#@~^wgAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d8'/n.7+.R1.lOW(L+1YvEl9W[4cDn^KDNknYr#@#@&d5V{Jk+s+^O,e~0MG:,tnsa/W.Or@#@&DkqcW2x,/;^~1Gx	~q@#@&h4r^+P	GY,D/8 nW6Pmx[PUGDP./8 4K0@#@&4TwAAA==^#~@%>
<option value="<%=#@~^DwAAAA==.kFcrtVwkWMOJ*HAUAAA==^#~@%>"<%=#@~^KAAAAA==dVn1YN`M/8cJ4+swkW.Or#~Md`rt+^2dWMYr#bzQ0AAA==^#~@%>><%=#@~^DwAAAA==.kFcrtVwkWMOJ*HAUAAA==^#~@%></option>
<%#@~^HAAAAA==@#@&Dd8RsW\xaY@#@&A+	N@#@&rQYAAA==^#~@%>
</select><font color=red>*</font>&nbsp;&nbsp;&nbsp;&nbsp;(如果要增加或修改帮助类别，请<a href="edithelpsort.asp">由此进入</a>)
</td>
</tr>
<tr>
<td>帮助内容</td><td><textarea rows="9" name="helpcontent" cols="49"><%=#@~^EQAAAA==.k`E4+^wmKxDnxDJbHgYAAA==^#~@%></textarea><font color=red>*</font></td>
</tr>
</table>
<font color=red>*</font>项必须填写&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="修改">&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="删除" onclick="return window.confirm('你确定要删除此条帮助信息吗？')";>
<input type="hidden" name="id" value=<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>>
</form>
<%#@~^GAAAAA==@#@&+U9Pb0@#@&+	[Pb0@#@&kQQAAA==^#~@%>
</center>

</body>
</html>










