<%Response.Expires=0%>

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>������</title>
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
alert('���ݲ����ظ���');
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
alert('��ѡ��������');
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
˵����ɫ:
<select name='sayscolor' onChange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option style="background-color: #660099; color:#660099" value="0">Ĭ�� </option>
<option style="background-color: #000000; color:#000000" value="1">��ɫ </option>
<option style="background-color: #0088FF; color:#0088FF" value="2">���� </option>
<option style="background-color: #0000FF; color:#0000FF" value="3">���� </option>
<option style="background-color: #000088; color:#000088" value="4">���� </option>
<option style="background-color: #888800; color:#888800" value="5">���� </option>
<option style="background-color: #008888; color:#008888" value="6">���� </option>
<option style="background-color: #008800; color:#008800" value="7">��� </option>
<option style="background-color: #8888FF; color:#8888FF" value="8">���� </option>
<option style="background-color: #AA00CC; color:#AA00CC" value="9">��ɫ </option>
<option style="background-color: #8800FF; color:#8800FF" value="10">���� </option>
<option style="background-color: #888888; color:#888888" value="11">��ɫ </option>
<option style="background-color: #CCAA00; color:#CCAA00" value="12">���� </option>
<option style="background-color: #FF8800; color:#FF8800" value="13">��� </option>
<option style="background-color: #FF0088; color:#FF0088" value="14">õ�� </option>
<option style="background-color: #FF00FF; color:#FF00FF" value="15">�Ϻ� </option>
<option style="background-color: #FF0000; color:#FF0000" value="16">��� </option>
</select>
<%=Session("UserName")%>˵:
<input type=text name='saystemp' style='font-size:12px' size=50 maxlength=100>
<input type=submit value='����' style='font-size:12px'>

<br>
������ɫ:
<select name='addwordcolor' onChange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option style="background-color: #008888; color:#008888" value="0">Ĭ�� </option>
<option style="background-color: #000000; color:#000000" value="1">��ɫ </option>
<option style="background-color: #0088FF; color:#0088FF" value="2">���� </option>
<option style="background-color: #0000FF; color:#0000FF" value="3">���� </option>
<option style="background-color: #000088; color:#000088" value="4">���� </option>
<option style="background-color: #888800; color:#888800" value="5">���� </option>
<option style="background-color: #008888; color:#008888" value="6">���� </option>
<option style="background-color: #008800; color:#008800" value="7">��� </option>
<option style="background-color: #8888FF; color:#8888FF" value="8">���� </option>
<option style="background-color: #AA00CC; color:#AA00CC" value="9">��ɫ </option>
<option style="background-color: #8800FF; color:#8800FF" value="10">���� </option>
<option style="background-color: #888888; color:#888888" value="11">��ɫ </option>
<option style="background-color: #CCAA00; color:#CCAA00" value="12">���� </option>
<option style="background-color: #FF8800; color:#FF8800" value="13">��� </option>
<option style="background-color: #FF0088; color:#FF0088" value="14">õ�� </option>
<option style="background-color: #FF00FF; color:#FF00FF" value="15">�Ϻ� </option>
<option style="background-color: #FF0000; color:#FF0000" value="16">��� </option>
</select>


<input type="checkbox" name="toone" value="ON">���Ļ�

��<select name='towho' onchange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option value='���' selected>���
</select>
����:
<select name='addsays' onchange="document.forms[0].saystemp.focus();" style='font-size:12px'>
<option value="0" selected>�� 
<option value="1">΢Ц 
<option value="2">���� 
<option value="3">���� 
<option value="4">���� 
<option value="5">��Ц 
<option value="6">���� 
<option value="7">ս�� 
<option value="8">ë�� 
<option value="9">��� 
<option value="10">���� 
<option value="11">ͬ�� 
<option value="12">�ֻ� 
<option value="13">��� 
<option value="14">�� 
<option value="15">ȭ�� 
<option value="16">���� 
<option value="17">�ź� 
<option value="18">���� 
<option value="19">�Ҹ� 
<option value="20">���� 
<option value="21">��ʹ 
<option value="22">���� 
<option value="23">���� 
<option value="24">���� 
<option value="25">���� 
<option value="26">ɵ 
<option value="27">���� 
<option value="28">�޴� 
<option value="29">�޹� 
<option value="30">���� 
<option value="31">���� 
<option value="32">���� 
<option value="33">�޲� 
<option value="34">���� 
<option value="35">��ĭ 
</select>

</form></div>
<tr align="center"><td>
</td></tr>
</body>
</html>
