<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/checked.asp"-->
<html>
<head>
<title>发送公文</title>
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
        <div align="center"><font color="#FFFFFF">数据传输中....请稍候</font></div>
      </td>
    </tr>
    <tr>
      <td bgcolor="f0f0f0">正在上传附件...</td>
    </tr>
  </table>
</div>

  <table width="95%">
    <tr> 
      
    <td width="53%"> 公文发送&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </td>
      
    <td width="10%"> 要发送给 </td>
      <form action="sendfile.asp" method="post" name="form1">
        
      <td width="12%"> 
        <select name="userdept" size=1 onChange="document.form1.submit();">
            <%#@~^XAEAAA==@#@&B打开数据库读出部门@#@&/YP1W	U'Kwnx94cEKl4!dXr~J1GUxr~rl^mndkNdxrb@#@&/nO,D/xdD-+M 1DnmYW4N+1O`rl[W94 .mWM[/YJ*@#@&/$V{Jd+sn1Y~f&jK&1/P,E/n.9+2Y,WMWh,Ek+Dbx6E@#@&DdRKwnU,/;^SmKxxBq@#@&b0,xGY~.kRnW6~l	N~UKYP.dc4G0,O4+U,0bD/DN2Y{Dd`rEdnMN+aOJ*@#@&bW~D;!+dYcE!/nD9nwDJb@!@*JJ~O4+UP6rM/O9+aY'M+$;+kYcJ!/n.9+wDE#@#@&h4rs+,xKY~Dd WWPmUN,xGO,D/ 8K0@#@&XGwAAA==^#~@%>
            <option value="<%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%>"<%=#@~^IgAAAA==dVn1YN`6kMdY9+2YBDdcrE/.NwYrbbgQwAAA==^#~@%>><%=#@~^DgAAAA==.k`E!/DNwDE#5gQAAA==^#~@%></option>
            <%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
          </select>
          <input type="hidden" name="sendto" value="<%=#@~^BgAAAA==dx[DWjQIAAA==^#~@%>">
        </td>
      </form>
      
    <td width="3%"> 的 </td>
      <form name="form2">
        
      <td width="13%"> 
        <select name="recipient" size=1>
            <option value="所有人">所有人</option>
            <%#@~^GAEAAA==@#@&/nDP1Wx	'K2+	N8`rWC8!/XrSJ1Wx	ESJmm1+d/[d	Jb@#@&d+DP.d{/+.-D mMnmYnK4N+mD`rCNKN8RM+^GMN/OJ*@#@&k5s'r/VnmO~	lh+B;/DUCs+PW.K:~EknMkU6PSt+M+,;/D[+aYxE,[Pk5VkYDvWrDkY9+2Yb~LPEPmUN,EdnMV+-n^@!@*B总管E~mx[,0KD4bN{vxKBE@#@&Dd Kw+	~/$V~1GUxBF@#@&AtrsPUWD~DkRnG6PlU[,xGY,.kR8K0@#@&8VkAAA==^#~@%>
            <option value="<%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%>(<%=#@~^DgAAAA==.k`E!/Dxm:E#2gQAAA==^#~@%>)"><%=#@~^CgAAAA==.k`E	ls+J*GwMAAA==^#~@%></option>
            <%#@~^GwAAAA==@#@&Ddc:K\+	+XO@#@&hnx9@#@&fAYAAA==^#~@%>
          </select>
        </td>
      </form>
      <form name="form4">
        
      <td width="9%"> 
        <input type="button" value="增加" onClick="document.form1.sendto.value=document.form1.sendto.value+'|'+document.form1.userdept.value+':'+document.form2.recipient.value;document.form3.sendto.value=document.form1.sendto.value;">
        </td>
      </form>
    </tr>
  </table>

<script language="JavaScript">
 function form_check(){

   if(document.form3.sendto.value.length<1){window.alert("请选择发送目标");document.form2.recipient.focus();return (false);}


   if(document.form3.title.value.length<1){window.alert("标题不能空");document.form3.title.focus();return (false);}

                    }
   function form2_check(){

   if(document.form3.aType.value.length<1)
   {
   window.alert("分类名不能为空！");
   document.form3.aType.focus();
   return (false);}
   }					
</script>
<br>
<form method="post" enctype="multipart/form-data" name="form3" action="sendfileok.asp"  >
  <table border="0" width="550" cellpadding="3" cellspacing="1" bgcolor="#999999" align="center" height="258">
    <tr> 
      <input type="hidden" name="userdept" value="<%=#@~^CQAAAA==WbDdDNwY1QMAAA==^#~@%>">
      <input type="hidden" name="username" value="所有人">
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">发&nbsp;&nbsp;给</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> 
        <input type="text" name="sendto" size=56 value="<%=#@~^BgAAAA==dx[DWjQIAAA==^#~@%>" onFocus="document.form3.title.focus();">
        <font color=red>*</font></font></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">公文标题：</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> 
        <input type=text name="title" size=56>
        <font color=red>*</font></font></td>
    </tr>
    <tr> 
    </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="148"><font color="#FFFFFF">内&nbsp;&nbsp;容:</font></td>
      <td colspan="2" height="148" bgcolor="f0f0f0"> <font size="2"> 
        <textarea name="content" rows="7" cols="58"></textarea>
        </font></td>
    </tr>
    <tr> 
     </tr>
    <tr> 
      <td align="center" bgcolor="#336699" width="142" height="23"><font color="#FFFFFF">附件:</font></td>
      <td colspan="2" height="23" bgcolor="f0f0f0"> <font size="2"> </font>
        <table width="95%" border="0" cellspacing="1" cellpadding="2" align="center">
          <tr> 
            <td><font color="#FF0000">*</font>上传的文件类型：不限类型</td>
          </tr>
          <tr> 
            <td><font color="#FF0000">*</font>上传的文件大小不能超过个2,500,000字节 (2.5M)</td>
          </tr>
          <tr> 
            <td> <font color="#FF0000">*</font>每次可以最多设置同时上传50个文件。<br>
              <script language="JavaScript">
	  function setid()
	  {
	  str='<br>';
	  if(!window.form3.upcount.value)
	   window.form3.upcount.value=1;
	  if(window.form3.upcount.value>50){
	  alert("您最多只能同时上传50个附件！");
	  window.form3.upcount.value = 5;
	  setid();
	  }
	  else{
 	  for(i=1;i<=window.form3.upcount.value;i++)
	     str+='<div align="center">附件'+i+':<input type="file" name="file'+i+'" style="width:350"></div><br><br>';
	  window.upid.innerHTML=str+'<br>';}
	  }
	    
</script>
              设置上传的个数 
              <input type="text" class="tx" value="1" name="upcount">
              <input type="button" name="Button" class="bt" onClick="setid();" value="· 设定 ·">
            </td>
          </tr>
        </table>
        <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td align="left" id="upid" height="2"> 
              <div align="center"><br>
                附件1: 
                <input type="file" name="file1" style="width:350" class="tx1" value="" size="50">
              </div>
      </td>
          </tr>
        </table>
        
      </td>
    </tr>
  </table>
  <div align="center">
    <input type="submit" name="Submit" value="发送" onclick=MM_showHideLayers('Layer1','','show')>
  </div>
</form>
</body>
</html>
<