<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%#@~^SAMAAA==@#@&BR O OO O RO ORO ORR OO RO OO RRO O ORORR @#@&Wm8EkXUCs+'.n$En/D 1WG0k/`rWm8EkXUls+Eb@#@&Wm8EkXEkn.xm:'.+5;/OR1GW0kndvJWC8!/zEknMxCs+r#@#@&Wm8EkX;/D[naY'Mn;!+/D ^WK3b+d`EGm4;/H;/D[naYJb@#@&WC4!dHEdD^+\V{.+$En/DR^GK3kd`rWl(;dX!/Ds+-n^Jb@#@&r0,WC8!/X;dDUlsn{JE,Y4+x,@#@&7D/2W	/n SDkDn`r@!/1.rwDP^lUo;CT+xJr%l7ld^MkwOEr@*E#@#@&iDnkwKx/RS.kD+cJSkU[KhRDGwcVW1COkKxct.+WxENn0m;VDRCdaBiEb@#@&7DdaWUk+chDbYcJ@!zdmMk2O@*J#@#@&dM+/aGU/Rx[@#@&n	N~k6@#@&@#@&vR OORR ORO R OR O OO O RO ORO ORR OO R@#@&k0,.n;!+kYcJd;(:rYrb'r应用J~O4+x@#@&mVsWSmNrD{4+Va'Mn;!+dYvJCs^Wh|nNbY{4nswr#@#@&r0~C^VGh|nNbYm4VwxErPOtU,ls^WS{+9kDmtV2'rxGE@#@&/OP1Wx	xGwx94cJGC(EdXrSJ1WUUr~JC^1+d/9d	Jb@#@&k;V{J!2NmYnP!/n.bx0,d+DPJ@#@&d;^'k;sP'~rlsVKA{NrO|t+s2{J~[,d$VdDDvlV^WSm+9kO{4+s2*P[,EPSt+Mn~k9'rP'P.n$En/DcJbNEb@#@&mGU	R36^!Yn,/$V@#@&@#@&nx9Pr0@#@&kAEBAA==^#~@%>
<%#@~^swEAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~CPWDKhP!/n.bx0~A4+.+,;k+.	ls+'EJ,'PKl8EkX;dDxmh+LJBr@#@&DkRKwnx~d$VSmKUxBF@#@&mVVGA|+[kDm4+sa'M/`rl^sWS{nNbYm4VwrbP,PP,@#@&mKx	R^VGd@#@&/OP1WUU{xWO4bxL@#@&dY~M/{xWDtbUo@#@&r0,mGG0{l^sWS{mKUODKV|lsVm;k+.'rUWrPO4x@#@&./2W	dRAMkD+`r@!6GxDP^W^W.xM+N,dk.+'rEQFrJ@*对不起，您没有这个权限！@!&0GUD@*E#@#@&dM+d2Kx/n x[@#@&7x[,k6@#@&qoQAAA==^#~@%>
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
      <td>编辑员工帮助文件管理权限设置&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%#@~^uQAAAA==@#@&B打开数据库，读出部门@#@&/YP1W	U'Kwnx94cEKl4!dXr~J1GUxr~rl^mndkNdxrb@#@&/nO,D/xdD-+M 1DnmYW4N+1O`rl[W94 .mWM[/YJ*@#@&/$V{Jd+sn1Y~f&jK&1/P,E/n.9+2Y,WMWh,Ek+Dbx6E@#@&DdRKwnU,/;^SmKxxBq@#@&bjYAAA==^#~@%>
<form method="post" action="userkqmanager.asp">
<td>
<select size=1 name="userdept">
<%#@~^pgAAAA==@#@&kW,xKYPM/cnW6PCx9PUGDPDk 4K0PD4nx,Ek+.Nn2D'./vEEk+.[wYEb@#@&r0,.;;/D`J!/.NwOJ*@!@*ErPY4nx,E/.[+aY{Dn;;nkYcJ!d+MNn2DJ#@#@&StrV~	WO,DkR+K0,Cx9PUWDP.dc4W6@#@&MzUAAA==^#~@%>
<option value="<%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%>"<%=#@~^IQAAAA==dVn1YN`!/.NwO~M/cE!/+M[+aYJ*bGAwAAA==^#~@%>><%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%></option>
<%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
</select><input type="submit" value="查看">
</td>
</form>
</tr>
</table>
</center>

<br>
<center>
  <table border="1"  cellspacing="0" cellpadding="0" width="95%" bgcolor="#FFFFFF" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr bgcolor="#EFEFEF">
      <td align=center height=25>姓名</td>
      <td align=center>部门</td>
      <td align=center>职位</td>
      <td align=center>可设置员工帮助文件管理</td>
      <td>&nbsp;</td>
    </tr>
    <%#@~^6gAAAA==@#@&B显示用户表@#@&/YP1W	U'Kwnx94cEKl4!dXr~J1GUxr~rl^mndkNdxrb@#@&/nO,D/xdD-+M 1DnmYW4N+1O`rl[W94 .mWM[/YJ*@#@&/$V{Jd+sn1Y~e,WDK:~;k+DrU6PAt.P;k+MN+aY{EPLPd;^/O.vE/.NwY*@#@&DkRKwnx~d$VSmKUxBF@#@&Stksn,xGY,.kRnK0,lx9P	GY,DdR(WW@#@&tkgAAA==^#~@%>
    <form method="post" action="edit_helpmanage.asp">
      <tr> 
        <td align=center><%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%></td>
        <td align=center><%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%></td>
        <td align=center><%=#@~^DwAAAA==.k`E!/DV\sJ*UQUAAA==^#~@%></td>
        <td align=center>
          <input type="checkbox" name="allow_edit_help" value="yes"<%=#@~^JAAAAA==^4+^0+9`Dk`rCV^WA{NrO|t+^2J*~JHndJ*fwwAAA==^#~@%>>
        </td>
        <td align=center>
          <input type="submit" name="submit" value="应用">
        </td>
      </tr>
      <input type="hidden" name="id" value=<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>>
      <input type="hidden" name="userdept" value=<%=#@~^CAAAAA==;k+.9+aYbAMAAA==^#~@%>>
    </form>
    <%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
  </table>
</center>
<br>
<%#@~^CAAAAA==@#@&@#@&LgAAAA==^#~@%>

</body>
</html>