<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css.css">
</head>

<body text="#000000" leftmargin="0" topmargin="0">
<form name="form" method="post" action="saveup.asp" enctype="multipart/form-data" >
<input type="hidden" name="filepath" value="">
<input type="hidden" name="act" value="upload">
&nbsp;文件：
<input type="file" name="file1" size=10>
<input type="submit" name="Submit" value="上传" onclick="parent.document.kbbs.B1.disabled=true,
parent.document.kbbs.B2.disabled=true;"> 类型：<select size="1" name="D1" style="font-size: 9pt">
<option>可上传类型</option>
<option>doc</option>
<option>xls</option>
<option>gif</option>
<option>jpg</option>
<option>bmp</option>
<option>png</option>
<option>swf</option>
<option>zip</option>
<option>rar</option>
</select>，限制：<%=session("upnum")%>个，每个 <%=session("upsize")%> K
</form>
</body>
</html>