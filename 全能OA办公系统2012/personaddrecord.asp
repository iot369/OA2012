<%@ LANGUAGE = VBScript.Encode %>
<%#@~^EgAAAA==./2Kxk+R6arD/x!CgcAAA==^#~@%>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%#@~^HgMAAA==@#@&WC(EkXxm:xD;;+kY ^KW3bn/vJWm8;/Hxm:nJb@#@&WC4!dX!/n.	l:nxM+5EdDR^KW0k+k`rGl(EdX!/n.	l:E#@#@&Wm8;/HEk+.Nn2D'.+$;+kY ^KW3rnk`EWm8!/z!/DNwDE#@#@&Gl(Edz!/+Ms+7+V{.n;!+kY mGG0kn/vEWm4;dHE/n.^+-+^E*@#@&b0,Wl(EkzEk+.xm:nxrJPD4+	P@#@&7.+kwKxd+ AMkO+vE@!km.raYPsC	o;lTn{JENl7l/1Db2YrJ@*J*@#@&7M+/aGxk+RS.rY`rhrx[GSROWa VKmCObWx 4M+W'E[0C!VDRlkwEIJ*@#@&dM+d2Kx/ hMkYcE@!J/1DrwO@*r#@#@&i.+kwGUk+RnU9@#@&+	[,kW@#@&@#@&6s'D.ks`.+$EndD`JXhJ*#@#@&a8'M+$En/Ocr68J*@#@&.hxOMk:c.;;+kOvJ"SJ*#@#@&V(xYMkh`M+5;/YvEV(J#*@#@&NS'DDr:c.;;+kO`rNAE*#@#@&".N4'D.b:cM+$E+kYvEy.N4J*#@#@&Hy4sxYMk:v.n;!+kYcJz"(:E#*@#@&9tW%{YDrhvDn;!nkYcrN40Lr#*@#@&kLxYMkhcM+;!n/D`Jk%E#*@#@&N"X%xDDr:v.+$EndD`J["HLE#*@#@&t%{YMk:vD5E/O`rt%E*#@#@&"y9y'D.r:vD;;+dOvJ"y9"J*#@#@&1y'O.b:cD5!+dD`rmyr#*@#@&(yxYMkhcM+;!n/D`J("E#*@#@&mf0AAA==^#~@%>
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>伴江行办公系统</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
              <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">个人通讯录</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
    <tr> 
      <td>增加个人通讯录资料&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <form method="post" action="personlist.asp">
        <td> 
          <input type="submit" value="返回">
        </td>
      </form>
    </tr>
  </table>
</center>

<%#@~^hgUAAA==@#@&kW,D;E/DcJkE8:bYEb{J输入J,Otx@#@&rWPX:{JEPG.,NA'rEPDtnU@#@&d.nkwGxknch.bY`J@!/1.kaY~VmxL;mo+{EJNl\md^DbwDJE@*Eb@#@&7DdwKxdnchDrO`El^nMYcrJ姓名和单位不能为空！rJ#pJ*@#@&iDn/aWUdRhMrY`J4rdYKDHRLWcR8#IJ*@#@&iDndaWxdnch.kDnvJ@!J/1DkaY@*E#@#@&7D/2G	/+cnx9@#@&U[Pb0@#@&d+O~1WUx{Gwx[8vJWC8!/zJBE1WU	JBJl1md/9/UJ*@#@&dYPMd'k+D7n.R1DlO+G8N+^YvEl9W[8cD+^GMNd+DE*@#@&k;^'Jk+^nmDPMP6DGh,w+MdW	D+1G.N,h4+.+~as'vJLa:LJvE@#@&Dd Kwnx,d$VS1W	x~8@#@&r0,xGY,Dd W0,GD,xWD~./c4K0~Y4n	@#@&dMn/aWUdRh.rD+cJ@!d1DraY,Vl	o!Co'EJNl-CkmDb2YrJ@*rb@#@&iD/2WUdRADbO+vJCsDYcEr已有该用户的资料，请重新输入姓名！JE#pE*@#@&iD/wKxknRSDrY`E4b/YK.XcoWvRq#pJ*@#@&d.nkwGxknRSDrO`J@!&km.kaO@*Jb@#@&iD+kwKU/Rnx9@#@&n	NPbW@#@&/;^~xPrq	/nDO~&xOW,2+M/GUM+mG.9PcY4rkkU6W!/+Mxmh+BDnmKD[OHw+Ba:BmWs2CxH~!/nD"ABmG:aCxHYnsB0laS4Wh+Dn^~nslbV~4Wsnl9N.+k/S2K/Y1CD9~/aStmx9/nYS^mVs/O~M+hCM3#~#mV;+kc,J@#@&/$VP{Pk5V,[~?$VjOM`Wm8EkXEkn.xm:#~[~EBPE@#@&d;^Px~k;V~',mdYMc^4b,[,J~,J@#@&/$V~',/5s,[PU5VUYDvah#,[,JSPE@#@&/5V,xPk;s~LP?5sUY.`9A*P',JBPJ@#@&k5V,'~/$V~',?;^jYM`ySb~[,JBPE@#@&d$V~',d;^P'~U;VjOM`[t6%*P',JBPJ@#@&k5V,'~/$V~',?;^jYM`m.b~[,JBPE@#@&d$V~',d;^P'~U;VjOM`"y94*P',JBPJ@#@&k5V,'~/$V~',?;^jYM`N.z%#,[,JSPE@#@&/5V,xPk;s~LP?5sUY.`."9yb,[,J~,J@#@&/$V~',/5s,[PU5VUYDvz"4s#,[~JS~r@#@&/$sP{Pd5^P[~j$VjYMcX4b,[,J~,J@#@&/$V~',/5s,[PU5VUYDvd%#,[,JSPE@#@&/5V,xPk;s~LP?5sUY.`4%*P',JBJ@#@&/$s'k;sPLPd5^/YMc4.#[,EbJ@#@&1WUx 3X+^EDnPk;s@#@&mWUUcmsWkn@#@&dY,D/{xKOtbxL@#@&x5QBAA==^#~@%>
<br><br>
<center>
<font color=red >
成功输入个人通讯录资料！<input type="button" value="继续增加" onclick="location.href='personaddrecord.asp'">
</font>    
</center>    
<%#@~^FAAAAA==~,P~@#@&V/P,~P@#@&1wIAAA==^#~@%>    
<script Language="JavaScript">    
    
 function form_check(){    
   var l1=document.form1.xm.value.length;    
   if(l1==0){window.alert("姓名为必填项！");document.form1.xm.focus();return (false);}    
    
   var l2=document.form1.dw.value.length;    
   if(l2==0){window.alert("单位为必填项！");document.form1.dw.focus();return (false);}    
                    }    
    
</script>    
    
<center>    
<br>    
<form method="post" action="personaddrecord.asp" name="form1" onsubmit="return form_check();">    
    <table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table border="0" cellpadding="0" cellspacing="0" width="540">
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">姓名</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="192"> 
          <input type=text name="xm" size=23>
          <font color=red>*</font></td>
        <td width="82" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">性别</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="164"> 
          <select name="xb" size="1">
            <option value="男">男</option>
            <option value="女">女</option>
          </select>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">职务</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="192"> 
          <input type=text name="zw" size=23>
        </td>
        <td width="82" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">通讯录类别</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="164"> 
          <select size="1" name="lb">
            <%#@~^8QEAAA==~@#@&7i/YP1W	U'Kwnx94cEKl4!dXr~J1GUxr~rl^mndkNdxrbP@#@&77k+Y~.k'd+M-D 1DlYW(%+1YcJmNG[(RD^WMN/OE#,@#@&d7/5s{Jd+^nmDPM~6DWh~a+./KUDX2PSt+M+,;/DUls+xvr[Wm8EkXEkn.xm:[EBE~@#@&7dMdRKwnU,/;sS1WUxBq,@#@&idb0PM/cnW6PGD,Dd (W0,OtxP@#@&7dimKxUR^sK/nP@#@&diddnDPDdx	WOtbUTP@#@&didD/aGxk+ hMkOnvJ@!k^DbwY,sCxTEmon'EENl-lk^DbwOEr@*Jb~@#@&7di./2Kxk+RSDbO+vJsW1lOrKxR4.+6'Jr2nDkW	l[NOza+ lk2JriEb,@#@&77iDn/aG	/nchMkY`r@!zkm.kaY@*E*P@#@&7diD+k2Gxk+c+UN~@#@&d7+	[Pb0~@#@&dd[G,h4k^n,xGDPM/RW6~@#@&d7CJEAAA==^#~@%>
            <option value="<%=#@~^CAAAAA==.k`EbNr#RwIAAA==^#~@%>"><%=#@~^DgAAAA==.k`EDXa+xm:E#3QQAAA==^#~@%></option>
            <%#@~^SgAAAA==~@#@&7iDkR:K\U+XY~@#@&d7sKWw,@#@&idmKUUR1VK/nP@#@&idd+D~Dk'UGDtkUL,@#@&diJxEAAA==^#~@%>
          </select>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">单位</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"> 
          <input type=text name="dw" size=60>
          <font color=red>*</font></td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">住宅电话</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="192"> 
          <input type=text name="zzdh" size=23>
        </td>
        <td width="82" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">邮政编码</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="164"> 
          <input type=text name="yzbm" size=23  maxlength="6">
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">电话或分机</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="192"> 
          <input type=text name="dhfj" size=23>
        </td>
        <td width="82" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
          <p align="center">手&nbsp;&nbsp;&nbsp;&nbsp;机</p>        </td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="164"> 
          <input type=text name="sj" size=23>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">M S N</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="192"> 
          <input type=text name="hj" size=23>
        </td>
        <td width="82" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
          <p align="center">Email
        </td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="164"> 
          <input type=text name="dzyj" size=23>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">住宅地址</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"> 
          <input type=text name="zzdz" size=50>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">传真</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"> 
          <input type=text name="cz" size=50>
        </td>
      </tr>
      <tr> 
        <td width="102" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">备注说明</td>
        <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" colspan="3"> 
          <textarea rows="4" cols="59" name="bz"></textarea>
        </td>
      </tr>
    </table>   
   
  <font color=red>*</font>必须填写&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="输入">&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="返回" onclick="window.location.href='personlist.asp';">   
</form>   
</center>   
<%#@~^GwAAAA==~,P@#@&+	NPb0,~P@#@&~P,@#@&iwMAAA==^#~@%>   
</body>   
</html>