<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/checked.asp"-->
<html>
<head>
<title>���͹���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css" type="text/css">
<script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div id="Layer1" style="position:absolute; left:171px; top:445px; width:268px; height:103px; z-index:1; visibility: hidden"> 
  <table width="95%" border="2" cellspacing="0" cellpadding="5">
    <tr>
      <td bgcolor="#FF0000">
        <div align="center"><font color="#FFFFFF">���ݴ�����....���Ժ�</font></div>
      </td>
    </tr>
    <tr>
      <td bgcolor="f0f0f0">�����ϴ�����...</td>
    </tr>
  </table>
</div>

  <table width="95%">
    <tr> 
      
    <td width="53%"> ���ķ���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </td>
      
    <td width="10%"> Ҫ���͸� </td>
      <form action="sendfile.asp" method="post" name="form1">
        
      <td width="12%"> 
        <select name="userdept" size=1 onChange="document.form1.submit();">
            <%#@~^XAEAAA==@#@&B�����ݿ��������@#@&/YP1W	U'Kwnx94cEKl4!dXr~J1GUxr~rl^mndkNdxrb@#@&/nO,D/xdD-+M 1DnmYW4N+1O`rl[W94 .mWM[/YJ*@#@&/$V{Jd+sn1Y~f&jK&1/P,E/n.9+2Y,WMWh,Ek+Dbx6E@#@&DdRKwnU,/;^SmKxxBq@#@&b0,xGY~.kRnW6~l	N~UKYP.dc4G0,O4+U,0bD/DN2Y{Dd`rEdnMN+aOJ*@#@&bW~D;!+dYcE!/nD9nwDJb@!@*JJ~O4+UP6rM/O9+aY'M+$;+kYcJ!/n.9+wDE#@#@&h4rs+,xKY~Dd WWPmUN,xGO,D/ 8K0@#@&XGwAAA==^#~@%>
            <option value="<%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%>"<%=#@~^IgAAAA==dVn1YN`6kMdY9+2YBDdcrE/.NwYrbbgQwAAA==^#~@%>><%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%></option>
            <%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
          </select>
          <input type="hidden" name="sendto" value="<%=#@~^BgAAAA==dx[DWjQIAAA==^#~@%>">
        </td>
      </form>
      
    <td width="3%"> �� </td>
      <form name="form2">
        
      <td width="13%"> 
        <select name="recipient" size=1>
            <option value="������">������</option>
            <%#@~^GAEAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~	lh+B;/DUCs+PW.K:~EknMkU6PSt+M+,;/D[+aYxE,[Pk5VkYDvWrDkY9+2Yb~LPEPmUN,EdnMV+-n^@!@*B�ܹ�E~mx[,0KD4bN{vxKBE@#@&Dd Kw+	~/$V~1GUxBF@#@&AtrsPUWD~DkRnG6PlU[,xGY,.kR8K0@#@&8VkAAA==^#~@%>
            <option value="<%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%>(<%=#@~^DgAAAA==.k`E!/Dxm:E#2gQAAA==^#~@%>)"><%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%></option>
            <%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
          </select>
        </td>
      </form>
      <form name="form4">
        
      <td width="9%"> 
        <input type="button" value="����" onClick="document.form1.sendto.value=document.form1.sendto.value+'|'+document.form1.userdept.value+':'+document.form2.recipient.value;document.form3.sendto.value=document.form1.sendto.value;">
        </td>
      </form>
    </tr>
  </table>

<script language="JavaScript">
 function form_check(){

   if(document.form3.sendto.value.length<1){window.alert("��ѡ����Ŀ��");document.form2.recipient.focus();return (false);}


   if(document.form3.title.value.length<1){window.alert("���ⲻ�ܿ�");document.form3.title.focus();return (false);}

                    }
   function form2_check(){

   if(document.form3.aType.value.length<1)
   {
   window.alert("����������Ϊ�գ�");
   document.form3.aType.focus();
   return (false);}
   }					
</script>
<br>
<form method="post" enctype="multipart/form-data" name="form3" action="sendfileok.asp"  >
  <table border="0" width="550" cellpadding="3" cellspacing="1" bgcolor="#999999" align="center" height="258">
    <tr> 
      <input type="hidden" name="userdept" value="<%=#@~^CQAAAA==WbDdDNwY1QMAAA==^#~@%>">
      <input type="hidden" name="username" value="������">
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">��&nbsp;&nbsp;��</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> 
        <input type="text" name="sendto" size=56 value="<%=#@~^BgAAAA==dx[DWjQIAAA==^#~@%>" onFocus="document.form3.title.focus();">
        <font color=red>*</font></font></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">���ı��⣺</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> 
        <input type=text name="title" size=56>
        <font color=red>*</font></font></td>
    </tr>
    <tr> 
    </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="148"><font color="#FFFFFF">��&nbsp;&nbsp;��:</font></td>
      <td colspan="2" height="148" bgcolor="f0f0f0"> <font size="2"> 
        <textarea name="content" rows="7" cols="58"></textarea>
        </font></td>
    </tr>
    <tr> 
     </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">����:</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> </font>
        <table width="95%" border="0" cellspacing="1" cellpadding="2" align="center">
          <tr> 
            <td><font color="#FF0000">*</font>�ϴ����ļ����ͣ���������</td>
          </tr>
          <tr> 
            <td><font color="#FF0000">*</font>�ϴ����ļ���С���ܳ�����2,500,000�ֽ� (2.5M)</td>
          </tr>
          <tr> 
            <td> <font color="#FF0000">*</font>ÿ�ο����������ͬʱ�ϴ�50���ļ���<br>
              <script language="JavaScript">
	  function setid()
	  {
	  str='<br>';
	  if(!window.form3.upcount.value)
	   window.form3.upcount.value=1;
	  if(window.form3.upcount.value>50){
	  alert("�����ֻ��ͬʱ�ϴ�50��������");
	  window.form3.upcount.value = 5;
	  setid();
	  }
	  else{
 	  for(i=1;i<=window.form3.upcount.value;i++)
	     str+='<div align="center">����'+i+':<input type="file" name="file'+i+'" style="width:350"></div><br><br>';
	  window.upid.innerHTML=str+'<br>';}
	  }
	    
</script>
              �����ϴ��ĸ��� 
              <input type="text" class="tx" value="1" name="upcount">
              <input type="button" name="Button" class="bt" onClick="setid();" value="�� �趨 ��">
            </td>
          </tr>
        </table>
        <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td align="left" id="upid" height="2"> 
              <div align="center"><br>
                ����1: 
                <input type="file" name="file1" style="width:350" class="tx1" value="" size="50">
              </div>
      </td>
          </tr>
        </table>
        
      </td>
    </tr>
  </table>
  <div align="center">
    <input type="submit" name="Submit" value="����" onclick=MM_showHideLayers('Layer1','','show')>
  </div>
</form>
</body>
</html>
<