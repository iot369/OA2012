<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/checked.asp"-->

<%#@~^wAEAAA==@#@&WC(EkXxm:xD;;+kY ^KW3bn/vJWm8;/Hxm:nJb@#@&WC4!dX!/n.	l:nxM+5EdDR^KW0k+k`rGl(EdX!/n.	l:E#@#@&Wm8;/HEk+.Nn2D'.+$;+kY ^KW3rnk`EWm8!/z!/DNwDE#@#@&Gl(Edz!/+Ms+7+V{.n;!+kY mGG0kn/vEWm4;dHE/n.^+-+^E*@#@&b0,Wl(EkzEk+.xm:nxrJPD4+	P@#@&7.+kwKxd+ AMkO+vE@!km.raYPsC	o;lTn{JENl7l/1Db2YrJ@*J*@#@&7M+/aGxk+RS.rY`rhrx[GSROWa VKmCObWx 4M+W'E[0C!VDRlkwEIJ*@#@&dM+d2Kx/ hMkYcE@!J/1DrwO@*r#@#@&i.+kwGUk+RnU9@#@&+	[,kW@#@&@#@&/x9OW{Dn;!+dOvJ/UNDWJ*@#@&3JMAAA==^#~@%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<title>伴江行办公系统</title>
</head>
<body  topmargin="0" leftmargin="5">

<center>
<table>
<tr>
<td>公文回复和转发&nbsp;&nbsp;</td>
<form method="post" action="showdate.asp">
<td><input type="submit" value="返回"><input type="hidden" name="id" value=<%=#@~^DQAAAA==.;;/D`JbNrbawQAAA==^#~@%>></td></form>
</tr>
</table>
</center>

<%#@~^6AkAAA==@#@&kW,D;E/DcJkE8:bYEb{J转发J,Otx@#@&v打开要转发的记录@#@&k+DP^WUU{W2+	[4vJGC(E/zEBJ^W	Ur~Emm1+/kNkUJ*@#@&?Y~.k'?.\DR;.nlD+}4%+^OvJ)f}9AcDn^KDNdnDJb@#@&d$Vxr/V+1Y,MP6DG:,/nU9NlDnPSt+Mn~k9'rP'P.n$En/DcJbNEb@#@&Dd Kwnx,d$VS1W	x~8@#@&OkDVnW^Nx.k`JDrY^+J*@#@&mKxD+UYGs9'./vEmKxOn	YJb@#@&/nx9nMWs9'M/`r/UNDE#@#@&Wr^+xmh+KVN{.d`r0bVnxChJb@#@&Wk^+/G	Y+UO:X2+Ks9'.k`r0k^+;GxD+UY:X2nr#@#@&Wk^+\ms;+KV9'./cE6ks+7CV!+Eb@#@&B读出转发人@#@&k+OP1G	xxKwxN(`rGl(EdXr~E^KxxrSJmmmddNkxr#@#@&jnDP./{j+M\n.cZDnCD+64Nn1YcrbGrf~RMnmKD[/YEb@#@&/$s'r/+^n^Y,xm:nPW.K:~EknDbxW~St+.n,Ed+MUm:n{J,[Pk;^dYM`Gl(Edz!/+MUls+#@#@&./cWa+UPd5^~^W	U~8@#@&^4lxLn	lh+{.k`E	ls+J*@#@&@#@&DkOV'OrDV+KsN@#@&mKUO+	Y{mGxOn	YGV9~[,J@!8M@*ORR ORO R OR O OO O R@!(D@*$转发人：rP'~1tl	L+	l:~'PrTLx8/2I]转发时间：J~[,UWS`b~LPJD@!(D@*J,',Dn$E/YvJ1GxD+UYr#@#@&k+x9OW{D+$;n/D`r/nx[OKJb@#@&[ksPhzk+x[OK@#@&:Hdx[DW{/w^kDc/x[YK~Ekr~O8SF*@#@&6G.Pl1t~/nU9YGk	WPbx~hH/+U[DW@#@&!dD[wDwWbxDxq	?ODv/nU9YWbU0BJ)rb@#@&b0,Ed+.[wOwKrxD@*T~Dt+U@#@&/nx9OKkU6Vx'^+	c/x[YKkUW*@#@&MnmbwkUOEk+MxC:nxMkLtDc/x[OKkxWSk+UNDGbxW^+	OEk+M[+aY2WbxOb@#@&k6~DmkarnxDEk+.xCh'E所有人J,Otx@#@&M+mr2b+UY!dDUm:'J所有人r@#@&nVk+@#@&!/n.	l:2WbxY{(U/DDvDnmr2b+UY!d+MxCh~JcE*@#@&EknMxCs+^+x{VU`M+^kaknUDE/.xm:+*@#@&Dmbwr+UO!/nD	C:'sn6Y`.n1k2kUDEdD	l:~!d+MxC:VnU F#@#@&DmkarnxDEk+.xCh'.kT4YvDn^bwknUDEd+MUm:nBEk+D	lsnVxRF EdnMxlsnwKkxDb@#@&x9Pr0@#@&M+^kar+	Y;dDNn2D's+6Ov/n	NDWk	0B;/D[+aY2GbxY q#@#@&@#@&dnY,mKxU'G2x[4vEWm4;dHJ~E^KxUJBEmm^/kN/	J*@#@&k+OPM/xdD\.R1D+mOnW(LmO`E)Gr9Ac.+1W.[k+YEb,@#@&/$s,'~r/V+1Y,MP6DG:,/nU9NlDnJ@#@&Dk 6wx,/5VS^KxU~8S&@#@&.dclN[Uh~@#@&.k`EDkDV+r#{OkDVn@#@&DdcrmW	O+	YJ*x^W	YxO@#@&.k`E/UNDEb{/+U[DGV9@#@&DdvJM+mbwbnxDEd+MxChJ#{.+1kwbnUY!/DUlhn@#@&./vEDmr2b+xO;k+.N2DJb{DmkakUY!/nD9+2O@#@&k6~0bV+	Ch+KV9@!@*JE~Dtnx@#@&Dk`EWbV+UCs+E#{WbVn	ls+W^N@#@&Dk`E0bVn/KxYUY:XwEb'6k^+/WUOxOKH2+KV[@#@&D/cE6ks+7C^Enr#clwa+	[m4EU3,0rs\l^;+KVN@#@&nx9Pb0@#@&.dcE2NmO+,@#@&.kRmsGk+~@#@&dY~M/{xWDtbUo,@#@&/Y~^Kxx{UWDtk	L~@#@&+	N~kW@#@&xn6D@#@&E把转发信息回复给原发送人@#@&dxNn.{Dn;!nkYcr/xNDrb@#@&Dnmbwrn	YEknD	l:x.+$E/O`E.mrwbnxDEdnMxlhnr#@#@&MnbNxM+$E+kYvEDk[J*@#@&ObYVxD;EdO`rYbYs+Eb@#@&^W	O+	YxE此公文已经转发给：rP[~dx[YK~LPE@!4M@*J,[,.+$En/D`E^KxYUYr#@#@&dnY,mKxU'G2x[4vEWm4;dHJ~E^KxUJBEmm^/kN/	J*@#@&k;s'rqUdDY,rxDWPknUN9lD+~`OrDVn~1GxD+UOB/+U[DSD^bwrxDE/D	C:~.+bNb~7lV!n/,`J@#@&d;^'k;sP'~k;s/D.`DkOs#P'~r~E@#@&d$Vxk;^P[,/$s/DDcmKxOn	Y#,'Pr~J@#@&d;^'k;sP'~k;s/D.`k+U[D#~',JSJ@#@&k;s{/$VPLPk5VkY.`M+^rak+	OEk+D	Ch+*PLPE~E@#@&/5V{d;^P'~M+k[~LPE#r@#@&mG	xc26m!O+,/5V@#@&ehYDAA==^#~@%>
<br><br><br>
<center><font color=red >转发完成</font></center>
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>





<%#@~^CAIAAA==@#@&kW,D;E/DcJkE8:bYEb{J回复J,Otx@#@&dnx9+M'.+5;/O`rd+	Nn.r#@#@&.mrwbn	Y;k+Mxls+{.+$En/D`E.mkar+	YEkn.xm:Jb@#@&.k['Mn;!+dOvJDnr9Jb@#@&ObYs'M+;!+kO`rYrY^+Eb@#@&mKUYxY{.n;!+kYcJ^G	YnxDE#@#@&dnDPmGU	'GwU94crWm4EkXrSJ1WUxr~EC1m+kdNkxJ*@#@&/$V{J(xdnMY~k	OW,/nU9NlOn,`OkDs~^KxD+xD~knx9+.~M+^rak+	OEk+D	Ch+BDk[#~-mV;+k~`r@#@&d$V'd5^P'Pk5^/OM`DkY^+*~[,JSJ@#@&d5^'/$sPLP/$sdYM`1WUYnUD#~[,E~r@#@&d$V'd5^P'Pk5^/OM`k+x9+MbPLPE~r@#@&d$V'k5V,[Pk5s/DDvDnmr2b+UY!d+MxCh#P'~r~E@#@&d$Vxk;^P[,DrN,[~J*J@#@&1Wx	 2X+m!OnPk;^@#@&A6MAAA==^#~@%><br><br><br>
<center><font color=red >回复完成</font></center>
<%#@~^DgAAAA==@#@&+U9Pb0@#@&VAIAAA==^#~@%>

<%#@~^CAAAAA==@#@&@#@&LgAAAA==^#~@%>

</body>
</html>