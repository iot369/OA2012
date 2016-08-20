<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>考勤系统</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body>
<form method="POST" action="finduser.asp" name="form" onsubmit="return check_form()">
<input type="hidden" name="xlh">
<input type="hidden" name="randnumber">
</form>
<OBJECT classid=clsid:4cb949a0-0976-11d5-90cb-0000b4c4c48f height=0 id="ePass" name="ePass" style="LEFT: 0px; TOP: 0px" width=0></OBJECT>
<OBJECT classid=clsid:5d9f2780-0976-11d5-90cb-0000b4c4c48f height=0 id="ctx" name="ctx" style="LEFT: 0px; TOP: 0px" width=0></OBJECT>
<script language="javascript">
//得到快狗设备序列号与随机数
function get_serialnumber()
{
	document.form.xlh.value=ctx.SerialNumber(1).toString(16)+ctx.SerialNumber(0).toString(16);
	ePass.GetChallenge();
	for(i=0;i<=7;i++)
	document.form.randnumber.value=document.form.randnumber.value+ePass.ChallengeBuf(i).toString(16);
}
//检查快狗设备是否存在
function check_device()
{
	var ErrCode,kqflagvalue;
	ErrCode=ePass.OpenDevice(1);
	if (ErrCode==0)
	{
		get_serialnumber();
		document.form.submit();
		ePass.CloseDevice();
	}
	else
	{
		ePass.CloseDevice();
	}
setTimeout("check_device()",1000);
}
check_device();
</script>
</body>
</html>
