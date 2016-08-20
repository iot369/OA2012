<%Response.Expires=0%>

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>发言区</title>
<style type='text/css'>
<!--
body {font-size : 9pt;}
input {font-size : 9pt;}
a  {font-size : 9pt; color : blue;text-decoration : none;}
a:hover  {color : blue;text-decoration : underline;}
-->
</style>
<script Language='JavaScript'>
function checksays() {
document.forms[0].says.value='';
if (document.forms[0].saystemp.value!=''){
if ((document.forms[0].oldsays.value==document.forms[0].saystemp.value) && (document.forms[0].oldtowho.value==document.forms[0].towho.options[document.forms[0].towho.selectedIndex].value)){
alert('内容不可重复！');
document.forms[0].saystemp.focus();
document.forms[0].saystemp.select();
return false;
}
document.forms[0].oldtowho.value=document.forms[0].towho.options[document.forms[0].towho.selectedIndex].value;
document.forms[0].oldtowho.value=document.forms[0].towho.value;
document.forms[0].says.value=document.forms[0].saystemp.value;
document.forms[0].oldsays.value=document.forms[0].saystemp.value;
document.forms[0].saystemp.focus();
document.forms[0].saystemp.value='';
return true;
}
if (document.forms[0].towho.value==''){
alert('请选择动作对象！');
return false;
}
document.forms[0].oldacttowho.value=document.forms[0].towho.options[document.forms[0].towho.selectedIndex].value;
document.forms[0].saystemp.focus();
document.forms[0].saystemp.value='';
return true;
}
</script>
</head>
<body bgcolor=#FFFFFF background="chat2.gif" text="304060" onload="javascript:self.focus()">
<div align=left>
<form method=POST action='f1.asp#bottom' target='f1' onsubmit='return(checksays());'>
<input type=hidden name='says' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldsign' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<input type=hidden name='oldacttowho' value>
说话颜色:
<select name='sayscolor' onChange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option style="background-color: #660099; color:#660099" value="0">默认 </option>
<option style="background-color: #000000; color:#000000" value="1">黑色 </option>
<option style="background-color: #0088FF; color:#0088FF" value="2">海蓝 </option>
<option style="background-color: #0000FF; color:#0000FF" value="3">亮蓝 </option>
<option style="background-color: #000088; color:#000088" value="4">深蓝 </option>
<option style="background-color: #888800; color:#888800" value="5">黄绿 </option>
<option style="background-color: #008888; color:#008888" value="6">蓝绿 </option>
<option style="background-color: #008800; color:#008800" value="7">橄榄 </option>
<option style="background-color: #8888FF; color:#8888FF" value="8">淡紫 </option>
<option style="background-color: #AA00CC; color:#AA00CC" value="9">紫色 </option>
<option style="background-color: #8800FF; color:#8800FF" value="10">蓝紫 </option>
<option style="background-color: #888888; color:#888888" value="11">灰色 </option>
<option style="background-color: #CCAA00; color:#CCAA00" value="12">土黄 </option>
<option style="background-color: #FF8800; color:#FF8800" value="13">金黄 </option>
<option style="background-color: #FF0088; color:#FF0088" value="14">玫瑰 </option>
<option style="background-color: #FF00FF; color:#FF00FF" value="15">紫红 </option>
<option style="background-color: #FF0000; color:#FF0000" value="16">大红 </option>
</select>
<%=Session("UserName")%>说:
<input type=text name='saystemp' style='font-size:12px' size=50 maxlength=100>
<input type=submit value='发言' style='font-size:12px'>

<br>
姓名颜色:
<select name='addwordcolor' onChange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option style="background-color: #008888; color:#008888" value="0">默认 </option>
<option style="background-color: #000000; color:#000000" value="1">黑色 </option>
<option style="background-color: #0088FF; color:#0088FF" value="2">海蓝 </option>
<option style="background-color: #0000FF; color:#0000FF" value="3">亮蓝 </option>
<option style="background-color: #000088; color:#000088" value="4">深蓝 </option>
<option style="background-color: #888800; color:#888800" value="5">黄绿 </option>
<option style="background-color: #008888; color:#008888" value="6">蓝绿 </option>
<option style="background-color: #008800; color:#008800" value="7">橄榄 </option>
<option style="background-color: #8888FF; color:#8888FF" value="8">淡紫 </option>
<option style="background-color: #AA00CC; color:#AA00CC" value="9">紫色 </option>
<option style="background-color: #8800FF; color:#8800FF" value="10">蓝紫 </option>
<option style="background-color: #888888; color:#888888" value="11">灰色 </option>
<option style="background-color: #CCAA00; color:#CCAA00" value="12">土黄 </option>
<option style="background-color: #FF8800; color:#FF8800" value="13">金黄 </option>
<option style="background-color: #FF0088; color:#FF0088" value="14">玫瑰 </option>
<option style="background-color: #FF00FF; color:#FF00FF" value="15">紫红 </option>
<option style="background-color: #FF0000; color:#FF0000" value="16">大红 </option>
</select>


<input type="checkbox" name="toone" value="ON">悄悄话

对<select name='towho' onchange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option value='大家' selected>大家
</select>
表情:
<select name='addsays' onchange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option value="0" selected>无 
<option value="1">微笑 
<option value="2">温柔 
<option value="3">脸红 
<option value="4">得意 
<option value="5">大笑 
<option value="6">神秘 
<option value="7">战兢 
<option value="8">毛手 
<option value="9">嘟嘴 
<option value="10">慢条 
<option value="11">同情 
<option value="12">乐祸 
<option value="13">快哭 
<option value="14">哭 
<option value="15">拳打 
<option value="16">坏意 
<option value="17">遗憾 
<option value="18">诧异 
<option value="19">幸福 
<option value="20">翻箱 
<option value="21">悲痛 
<option value="22">正义 
<option value="23">严肃 
<option value="24">生气 
<option value="25">大声 
<option value="26">傻 
<option value="27">满足 
<option value="28">无措 
<option value="29">无辜 
<option value="30">自语 
<option value="31">瞪眼 
<option value="32">想吐 
<option value="33">无采 
<option value="34">不舍 
<option value="35">白沫 
</select>

</form></div>
<tr align="center"><td>
</td></tr>
</body>
</html>
