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
<title>OA�칫ϵͳ.��Ե�ر��</title>
</head>
<body  topmargin="5" leftmargin="5">

<center>
<%#@~^tAAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~CPWDKhP4+s2bx0~A4+.+,r9'E,[,D+$EdYvJrNr#@#@&M/RK2+	P/$sSmKx	~q@#@&jDgAAA==^#~@%>
<table>
<tr>
<td>�༭<%=#@~^DwAAAA==.k`E4+^wYbY^nJ*RQUAAA==^#~@%>������Ϣ&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%#@~^JgAAAA==@#@&kW,D;E/DcJkE8:bYEb{Jɾ��J,Otx@#@&nwkAAA==^#~@%>
<form method="post" action="helpinf.asp">
<%#@~^DAAAAA==@#@&+sk+@#@&1wEAAA==^#~@%>
<form method="post" action="showhelpinf.asp">
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>
<td><input type="submit" value="����"><input type="hidden" name="id" value=<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>><input type="hidden" name="helpsort" value="<%=#@~^DgAAAA==.k`E4+^w/KDDE#6wQAAA==^#~@%>"></td></form></td>
</tr>
</table>
</center>


<center>
<%#@~^MQEAAA==@#@&kW,D;E/DcJkE8:bYEb{J�޸�J,Otx@#@&dnY,mKxU'G2x[4vEWm4;dHJ~E^KxUJBEmm^/kN/	J*@#@&k;sP{PE;aNlDnP4+VarU0,/Y~tnsa/GDDxJ,[~d$V/O.vDn;!nkYcrtVwkWMOJ*#~[,JS4VwDrY^+'r~'Pk;^/ODc.;;+kO`rtnsaYkOsJb#,',JS4+^wmKxDnxD'EPLPd5^/YMcD;EdO`rtV2mGUD+UYrb#,[~E,htn.PrN{E,[~M+$E+kYvEk9Jb@#@&mGU	R2Xnm!Y+,d5V@#@&YmAAAA==^#~@%>
<br><br><font color=red >�ɹ��޸İ�����Ϣ��</font>
<%#@~^qwAAAA==@#@&+sk+@#@&k6PMn;!+dYvJd;(:kDE#{Jɾ��J,O4+	@#@&/nY~^KxU'K2+	N8crWl8;kXE~r^KxUr~rlm1+kdNkxE#@#@&d5^P',ENV+Dn~0MWsP4+s2bxWPS4+M+~r9'J~',Dn;!nkYcrk9J#@#@&1Gx	R36m;OP/$s@#@&kzIAAA==^#~@%>
<br><br><font color=red >�ɹ�ɾ��������Ϣ��</font>
<%#@~^DAAAAA==@#@&+sk+@#@&1wEAAA==^#~@%>
<br>
<script Language="JavaScript">
 function form_check(){
   var l1=document.form1.helptitle.value.length;
   if(l1<1){window.alert("���������д");document.form1.helptitle.focus();return (false);}

   var l2=document.form1.helpcontent.value.length;
   if(l2<1){window.alert("���ݱ�����д");document.form1.helpcontent.focus();return (false);}
                    }
</script>
<center>
<br>
<form method="post" action="edithelpinf.asp" name="form1" onsubmit="return form_check();">
<table border="1"  cellspacing="0" cellpadding="0">
<tr>
<td>��������</td><td><input type="text" name="helptitle" size=50 value="<%=#@~^DwAAAA==.k`E4+^wYbY^nJ*RQUAAA==^#~@%>"><font color=red>*</font></td>
</tr>
<tr>
<td>�������</td><td>
<select name="helpsort" size=1>
<%#@~^wgAAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d8'/n.7+.R1.lOW(L+1YvEl9W[4cDn^KDNknYr#@#@&d5V{Jk+s+^O,e~0MG:,tnsa/W.Or@#@&DkqcW2x,/;^~1Gx	~q@#@&h4r^+P	GY,D/8 nW6Pmx[PUGDP./8 4K0@#@&4TwAAA==^#~@%>
<option value="<%=#@~^DwAAAA==.kFcrtVwkWMOJ*HAUAAA==^#~@%>"<%=#@~^KAAAAA==dVn1YN`M/8cJ4+swkW.Or#~Md`rt+^2dWMYr#bzQ0AAA==^#~@%>><%=#@~^DwAAAA==.kFcrtVwkWMOJ*HAUAAA==^#~@%></option>
<%#@~^HAAAAA==@#@&Dd8RsW\xaY@#@&A+	N@#@&rQYAAA==^#~@%>
</select><font color=red>*</font>&nbsp;&nbsp;&nbsp;&nbsp;(���Ҫ���ӻ��޸İ��������<a href="edithelpsort.asp">�ɴ˽���</a>)
</td>
</tr>
<tr>
<td>��������</td><td><textarea rows="9" name="helpcontent" cols="49"><%=#@~^EQAAAA==.k`E4+^wmKxDnxDJbHgYAAA==^#~@%></textarea><font color=red>*</font></td>
</tr>
</table>
<font color=red>*</font>�������д&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="�޸�">&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷ��Ҫɾ������������Ϣ��')";>
<input type="hidden" name="id" value=<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>>
</form>
<%#@~^GAAAAA==@#@&+U9Pb0@#@&+	[Pb0@#@&kQQAAA==^#~@%>
</center>

</body>
</html>










