<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0053)http://bbs.wj8.net/admincp.php?action=menu&sid=LXPP7l -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>A:link {
	COLOR: #003366; TEXT-DECORATION: none
}
A:visited {
	COLOR: #003366; TEXT-DECORATION: none
}
A:hover {
	TEXT-DECORATION: underline
}
BODY {
	FONT-SIZE: 12px; SCROLLBAR-ARROW-COLOR: #dde3ec; SCROLLBAR-BASE-COLOR: #f8f9fc; BACKGROUND-COLOR: #e9edf7
}
TABLE {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
TEXTAREA {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
INPUT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
OBJECT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
SELECT {
	FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #000000; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #f8f9fc
}
.nav {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-FAMILY: Tahoma, Verdana
}
.header {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/headerbg.gif); COLOR: #ffffff; FONT-FAMILY: Tahoma, Verdana
}
.category {
	FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/catbg.gif); COLOR: #000000; FONT-FAMILY: Tahoma
}
.multi {
	FONT-SIZE: 11px; COLOR: #003366; FONT-FAMILY: Tahoma
}
.smalltxt {
	FONT-SIZE: 11px; FONT-FAMILY: Tahoma
}
.mediumtxt {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
.bold {
	FONT-WEIGHT: bold
}
BLOCKQUOTE {
	BORDER-RIGHT: #dde3ec 1px dashed; PADDING-RIGHT: 5px; BORDER-TOP: #dde3ec 1px dashed; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; BORDER-LEFT: #dde3ec 1px dashed; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BORDER-BOTTOM: #dde3ec 1px dashed; BACKGROUND-COLOR: #ffffff
}
.code {
	PADDING-RIGHT: 5px; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BACKGROUND-COLOR: #ffffff
}
</STYLE>

<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY leftMargin=3 topMargin=3><BR>
<script language="javascript">
function checkform()
{
	var time1,time2,value1,value2;
	time1=document.form1.amgohour.value;
	time2=document.form1.pmcomehour.value;
	value1=document.form1.amgominute.value;
	value2=document.form1.pmcomeminute.value;
	if (time1==time2)
		{
			if ((value1==value2) || (value1=="30" && value2=="0"))
				{
					alert("上午下班时间与下午上班时间冲突！");
					return (false);
				}
		}
	return (true);
}
</script>
<title>红孩儿Office办公系统</title>
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5"><center>

月工作计划

<form method="POST" action="time1.asp" onsubmit="return checkform();" name="form1">
  <div align="center">
    <center>
      <table border="1" width="500" cellspacing="0" cellpadding="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  
          
          <td width="246" height="30" bgcolor="#FFFFFF">上午标准上班时间：
              <select size="1" name="amcomehour">
                <option value="6">6点</option>
                <option value="7">7点</option>
                <option value="8">8点</option>
                <option selected value="9">9点</option>
              </select>
              <select size="1" name="amcomeminute">
                <option selected value="0">00分</option>
                <option value="30">30分</option>
              </select>
              <script language="javascript">
document.form1.amcomehour.value=7;
document.form1.amcomeminute.value=0;
  </script>
          </td>
          
          <td width="248" height="30" bgcolor="#FFFFFF">上午标准下班时间：
              <select size="1" name="amgohour">
                <option value="11">11点</option>
                <option value="12" selected>12点</option>
                <option value="13">13点</option>
              </select>
              <select size="1" name="amgominute">
                <option selected value="0">00分</option>
                <option value="30">30分</option>
              </select>
              <script language="javascript">
document.form1.amgohour.value=12;
document.form1.amgominute.value=0;
  </script>
          </td>
        </tr>
        <tr>
          
          <td width="246" height="30" bgcolor="#FFFFFF">下午标准上班时间：
              <select size="1" name="pmcomehour">
                <option value="13">13点</option>
                <option value="14" selected>14点</option>
                <option value="15">15点</option>
              </select>
              <select size="1" name="pmcomeminute">
                <option selected value="0">00分</option>
                <option value="30">30分</option>
              </select>
              <script language="javascript">
document.form1.pmcomehour.value=14;
document.form1.pmcomeminute.value=0;
  </script>
          </td>
          
          <td width="248" height="30" bgcolor="#FFFFFF">下午标准下班时间：
              <select size="1" name="pmgohour">
                <option value="16">16点</option>
                <option value="17" selected>17点</option>
                <option value="18">18点</option>
                <option value="19">19点</option>
              </select>
              <select size="1" name="pmgominute">
                <option selected value="0">00分</option>
                <option value="30">30分</option>
              </select>
              <script language="javascript">
document.form1.pmgohour.value=17;
document.form1.pmgominute.value=0;
  </script>
          </td>
        </tr>
        <tr>
          <td width="246" height="30" bgcolor="#FFFFFF">上班考勤延迟时间：
              <select size="1" name="comedelaytime">
                <option value="0">0分钟</option>
                <option value="10">10分钟</option>
                <option value="15">15分钟</option>
                <option value="20">20分钟</option>
                <option value="25">25分钟</option>
                <option value="30">30分钟</option>
                <option value="35">35分钟</option>
                <option value="40">40分钟</option>
                <option value="45">45分钟</option>
                <option value="50">50分钟</option>
                <option value="55">55分钟</option>
              </select>
              <script language="javascript">
document.form1.comedelaytime.value=20;
  </script>
          </td>
          <td width="248" height="30" bgcolor="#FFFFFF">下班考勤提前时间：
              <select size="1" name="goaheadtime">
                <option value="0">0分钟</option>
                <option value="10">10分钟</option>
                <option value="15">15分钟</option>
                <option value="20">20分钟</option>
                <option value="25">25分钟</option>
                <option value="30">30分钟</option>
                <option value="35">35分钟</option>
                <option value="40">40分钟</option>
                <option value="45">45分钟</option>
                <option value="50">50分钟</option>
                <option value="55">55分钟</option>
              </select>
              <script language="javascript">
document.form1.goaheadtime.value=20;
  </script>
          </td>
        </tr>
        <tr>
          <td width="100%" height="30" bgcolor="#FFFFFF" colspan="2">考勤时间段：
              <select size="1" name="kqtimephase">
                <option value="10">10分钟</option>
                <option value="15">15分钟</option>
                <option value="20" selected>20分钟</option>
                <option value="25">25分钟</option>
                <option value="30">30分钟</option>
                <option value="35">35分钟</option>
                <option value="40">40分钟</option>
                <option value="45">45分钟</option>
                <option value="50">50分钟</option>
                <option value="55">55分钟</option>
              </select>
              <script language="javascript">
document.form1.kqtimephase.value=20;
  </script>
          </td>
        </tr>
        <tr>
          <td width="100%" height="30" bgcolor="#FFFFFF" colspan="2">
            <input type="checkbox" name="amgonokq" value=1>
            上午下班不考勤
            <input type="checkbox" name="pmcomenokq" value=1>
            下午上班不考勤
            <input type="checkbox" name="pmgonokq" value=1>
            下午下班不考勤
            <script language="javascript">

</script>
          </td>
        </tr>
      </table>
    </center>
  </div>
  <p align="center">
    <input type="submit" value="确定" name="submit">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <input type="reset" value="重填" name="resetbutton">
  </p>
</form>
</center>
        </td>
            
     
          </tr>
        </table>
    </tr>
  
  </table>
</body>
</html>