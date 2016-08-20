<%@ LANGUAGE = VBScript.Encode %>
<!--#INCLUDE FILE="asp/fupload.inc"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/checked.asp"-->

<%#@~^vQEAAA==@#@&B6	PADDKD,]+kEh+,1naD@#@&KC4!/X	Ch+{D;;+dOcmGW0r+k`EGm4Edz	lh+rb@#@&Gm4!/X!/.xm:n'M+5;/Yc^WK3kdcJKl(EdX;dDUlsnJ*@#@&Gm4Edz!/nD9naYxM+$E+kYc^WK3r+k`EGm4EkzEk+D9n2Yr#@#@&Gl8;kX;/.V\ns{D+5;/OR1GK3r/vJWm4!dX!/nD^+-n^J#@#@&k6PWm8;/HEk+.xCh'EJ,Otx~@#@&dDndaWU/ SDrD+vJ@!kmMrwDPsl	o;CT+'rELm\lk^.kaYrJ@*Jb@#@&d.+k2W	/n SDkOnvJAk	[Kh DWaRVKmmOkKx tM+WxEN+6CE^YRmd2BpJ*@#@&d.nkwGxknRSDrO`J@!&km.kaO@*Jb@#@&iD+kwKU/Rnx9@#@&n	NPbW@#@&@#@&RZIAAA==^#~@%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<title>伴江行管理系统</title>
</head>
<body  topmargin="5" leftmargin="5">

<center>
<table>
<tr>
<td>
公文发送&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
<form action="senddate.asp" method="post" name="form1">
<td>
<input type="submit" value="返回">
</td>
</form>
</tr>
</table>
</center>


<%#@~^tAgAAA==@#@&kW,I;E/D ?D-+M.C.bl4^n/vJIA}i2UK|H3Ku6GJbP{~JhrjPrPK4n	@#@&B R OR O OO O RO ORO ORR OO @#@&ED+k2Gxk+ch.kOn,J开始发送@!4M@*J@#@&vR OORR ORO R OR O OO O RO O@#@&Gkh~wk+^[/@#@&jasGl9?bynSrhbYx*ZT!Z!T@#@&?+O~wknV9d,'~V+Djw^Wm[`*@#@&Nb:~ob+V9@#@&wWD,3Cm4PwknV[~&x~sbnV9/ (D+:d@#@&kWPwrV[cxm:+{JDrY^+EPDtnU,YkDs+{Ak	C.X:WUY.kULvsr+^[R7ls;#@#@&r6Poks9RUm:'J^4r~Y4+UP^4x$bxlMzKK?YMrUov0b+sN -mV;+*@#@&b0~ob+V[ 	lh+{E1WUD+	YJ,Y4nx,mGxD+UO{Ak	CDHKWUO.k	ovsr+s[c\CV!n#@#@&rW,skns9RUlsn{Jdx9YWrPD4+	Pd+	NOG{Ak	CDHKWUO.k	ovsr+s[c\CV!n#@#@&rW,skns9RUlsn{JWbVFJ,Y4nx@#@&Wk^+UCs+'6r+^NRwrs+gls+@#@&Wr^+/W	O+	YPza+'WrV[R;G	Yn	Y:Xw@#@&Wk^+-l^Enx6k+^[R7lV!n@#@&x9Pr0@#@&	+aY@#@&B ORR OORR ORO R OR O OO O RO ORO ORR O@#@&vD/wKUd+chMkO+~EDkOVxJ,[~ObYVn~LPE@!(.@*J@#@&BM+/aW	d+ch.kD+~E1WxDnxD'J,'~mKxD+UY~',J@!4M@*J@#@&v./wGUk+ hMrD+~r/xNDW{EPLPd+	NOG,[Pr@!4M@*J@#@&vD/aWU/n SDrY~J6ksn9xlhn{J~[,WbVn	ls+PLPr@!4M@*E@#@&B.nkwW	d+chDbOnPr0bVnZGUD+UY:zw'E~LP0rsZGxDn	YPHwP[,J@!8D@*J@#@&EORR OO RO OO RRO O ORORR ORO RO ORR OORR@#@&[ks~sXdx9YW@#@&sz/x[YK'd2^kYvd+	NYKSEur~ FSFb@#@&0GD,nl1t~dxNOGbxWPbU,:zk+	NYK@#@&;/D[+aY2GbxY{(xUYDvdnx9YKkU0SE=Jb@#@&r0,EdnMN+2OaWrxD@*ZPO4+	@#@&k+	[YKkU0^+Ux^+xvd+	NYKrU0*@#@&Dnmr2b+UY!d+MxCh'DrL4Yc/U9YGbx6~/x9OWbxWVxR;k+D9nwDwWbUO#@#@&b0~Dn^bwr+	OEk+.Um:+xE所有人rPOtU@#@&.mbwkxD;/DUls+xE所有人r@#@&s/@#@&!dnD	ls+2WrUD'(xkODvDn^bwknUDEd+MUm:nBJvJ#@#@&!d+MxC:VnU{V+	cDmkarnxDEk+.xCh#@#@&Mnmbwrn	YEdnMxC:x^+WD`M+mbwbnxDEd+MxCh~EknD	l:snx F*@#@&Dn^bwr+	OEk+.Um:+x.bo4Yv.mrakxY!/.xm:n~!/n.	l:s+	OF ;d+Mxm:nwGr	Yb@#@&nx9PrW@#@&Dn^bwr+	O!/nMNwY{VWYv/nx9YGr	0~!d+MN+aO2WbxDOq#@#@&k+OP1Gx	'G2xN8crWC4!dHJSrmKxxr~rCm1+d/9/UE*@#@&knY,D/{dnD7+MR^DnCD+G4NnmD`E)Grf$ M+^WM[k+Or#,@#@&k;^~',Jd+^+^O,eP6.WsP/U[NmYJ@#@&.dcr2+	~/$VS^KxxSqB&@#@&Mdcl[9xhP@#@&Md`rYrY^+Eb{YkDs+@#@&DkcENKm!:nxOOHwnJ*xV(@#@&.k`J^G	YnxDE*'^KxD+xD@#@&./vJd+	Nn.r#'KC4!/X!dnD	ls+@#@&.dvJ.+1rwb+UO!/+.Um:nJ*xM+^bwb+xDEknD	lh+@#@&.dvJD^kak+	O;/D9+2YEb{Dnmb2kxO;k+D[naY@#@&bW,0r^+	l:@!@*EJ,Y4+	@#@&.k`J6rVxlsnE#{0bVnxCh@#@&DkcJ6ksn;WxOn	YPXanr#x6k^+ZKxDnxDKzw@#@&.k`J6rV\l^;nJ*Rmw2+U[1t;x0~0bVn-mVEn@#@&+UN,r6@#@&M/cEw9lDnP@#@&./cmsGk+P@#@&/YPMdxxKY4kUo~@#@&/nY,^W	xxUKYtrUTP@#@&@#@&x[,k6@#@&@#@&	n6D@#@&@#@&9KsCAA==^#~@%>
<br><br>
<center>发送完成</center>
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>


</body>  
  
</html>  


