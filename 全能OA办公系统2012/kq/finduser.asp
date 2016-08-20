<%@ LANGUAGE = VBScript %>
<%Response.Expires=0%>
<!--#include file="../asp/EpassFunction.asp"-->
<!--#include file="../asp/EpassConst.asp"-->
<!--#include file="../asp/epassconn.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户登录</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body>
<%
'on error resume next
Function DispErrInfo(ErrInfo)
	Response.Write("<script language=""javascript"">")
	Response.Write("alert("&chr(34)&ErrInfo&chr(34)&");")
	Response.Write("parent(""banner2"").location.href=""kqcheck.asp"";")
	Response.Write("</script>")
End Function
Dim yhm,pincode,randnumber,xlh,Rs,sql,ePass
randnumber=Trim(Request.Form("randnumber"))
xlh=Trim(Request.Form("xlh"))
If randnumber="" Or xlh="" Then
	Call DispErrInfo("出现错误，请返回重填！")
	conn.close
	Response.End
End IF
Set Rs=Server.CreateObject("Adodb.Recordset")
sql="SELECT * FROM EPASS_USER_INFO WHERE SERIALNUMBER='"&xlh&"'"
Rs.Open sql,conn,1,1
If Rs.EOF Or Rs.BOF Then
	Call DispErrInfo("对不起，没有这个用户，请返回重填！")
	Rs.Close
	conn.close
	Response.End
Else
	Set ePass = CreateObject("EpsModu.ePass")
	Session("key")=Get_Soft_HmacMd5(randnumber,Rs("mainkey"))
	Response.Write("<input type=""hidden"" name=""dirid1"" value="&Rs("key1dir")&">")
	Response.Write("<input type=""hidden"" name=""fileid1"" value="&Rs("key1id")&">")
	Response.Write("<input type=""hidden"" name=""dirid2"" value="&Rs("key2dir")&">")
	Response.Write("<input type=""hidden"" name=""fileid2"" value="&Rs("key2id")&">")
	Response.Write("<input type=""hidden"" name=""randnumber"" value="&randnumber&">")
	Response.Write("<OBJECT classid=clsid:4cb949a0-0976-11d5-90cb-0000b4c4c48f height=0 id=""ePass"" name=""ePass"" style=""LEFT: 0px; TOP: 0px"" width=0></OBJECT>")
%>
<script language="javascript">
function get_key(KeyDir1,KeyId1,KeyDir2,KeyId2,Text_Str)
{
	var i,ErrCode,key,Text_Len;
	key="";
	Text_Len=Text_Str.length;
	for (i=0;i<Text_Len;i++)
	{		
		ePass.TextBuf(i)=Text_Str.charCodeAt(i);
	}
	ErrCode=ePass.HmacMd5(KeyDir1,KeyId1,KeyDir2,KeyId2,Text_Len);
	if(ErrCode==0)
	{
		for(i=0;i<=15;i++)
			key=key+ePass.DigestBuf(i).toString(16);
	}
	else
		key="";
	return key;
}
var keyvalue;
ErrCode=ePass.OpenDevice(1);
if (ErrCode==0)
{
	keyvalue=get_key(dirid1.value,fileid1.value,dirid2.value,fileid2.value,randnumber.value);
	if (keyvalue!="")
	{
		location.href="check.asp?randnumber=<%=xlh%>&key="+keyvalue;
	}
	else
	{
		alert("计算出错，请确定重试！");
		parent("banner1").location.href="kqcheck.asp";
	}
}
else
{
	alert("没发现eKey设备，请插入！");
}
ePass.CloseDevice();
</script>
<%
Set ePass=Nothing
Rs.Close
conn.close
End If
%>
</body>
</html>