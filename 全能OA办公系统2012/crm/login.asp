<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../asp/opendb.asp"-->
<!--#include file="../asp/sqlstr.asp"-->
<!--#include file="Connections/conn.asp" -->

    <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
%>
<html>
<style type="text/css">
<!--
.box {
	color: #FFFFFF;
	background-color: #FFFFFF;
	height: 1px;
	width: 1px;
	border: #FFFFFF;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���۹���ϵͳ</title>
<script language="JavaScript">
<!--
if (this.location.href == top.location.href){
    top.location.href = "";
}

function checkInput()
{
    var strItem1 = document.all.item1.value;
	var strItem2 = document.all.item2.value;
	if (strItem1 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">��û��Ȩ�޽���</font>"
		document.all.item1.focus();
		return false;
	}
	if (strItem2 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">��û��Ȩ�޽���</font>"
		document.all.item2.focus();
		return false;
	}
}

function checkChr(chr)
{
    var str0 = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
	var intLen = chr.length;
	for (var i=0;i<intLen;i++){
	    
	}
}
-->
</script>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<style type="text/css">

.style1 {color: #0d79b3}
.style2 {color: #FF0000}
.style8 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
</style>
</head>

<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style8"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style8"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">����ϵͳ</td>
                </tr>
            </table></td>
            <td width="1"><span class="style8"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br>
<td align="center">
</td>
<%
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And Session("CRM_level") > 0 Then Response.Redirect("listAll.asp")
Dim action
action = Trim(Request("action"))
Select Case action
Case "login"
    Call login()
Case Else
    Call loginForm()
End Select

Sub login()
    Dim account,password
	account = Trim(Request("item1"))
	password = Trim(Request("item2"))
	If account = "" Or password = "" Then
	    Response.Redirect("login.asp?errMsg=2")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uAccount = '" & account & "'",conn,3,1
	If rs.RecordCount <> 1 Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If password <> rs("uPassword") Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If rs("uBlock") = True Then
	    Response.Redirect("login.asp?errMsg=3")
		Response.End()
	End If
	Session("CRM_account") = account
	Session("CRM_name") = rs("uName")
	Session("CRM_level") = rs("uLevel")
	Session("CRM_group") = rs("uGroup")
	rs.Close
	Set rs = Nothing
	Response.Redirect("listAll.asp")
End Sub

Sub loginForm()
    Dim errMsg
	errMsg = CInt(ABS(Request("errMsg")))
	Select Case errMsg
	Case 2
	    errMsg = "<font color=""#FF0000"">�Բ���,������ҵ��Ա��,��û��ʹ������ϵͳ��Ȩ��</font>"
	Case 1
	    errMsg = "<font color=""#FF0000"">�Բ���,������ҵ��Ա��,��û��ʹ������ϵͳ��Ȩ��</font>"
	Case 3
	    errMsg = "<font color=""#FF0000"">�����˺���ʱ������</font>"
	Case Else
	    errMsg = ""
	End Select
%>
<table width="570" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top"><table width="92%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#88ADDF">&nbsp;</td>
          </tr>
        </table>          
          <br>          <span class="style1">����ϵͳ������:</span><br>
          <span class="style1">�ӿͻ���Դ��Ѱ�Ҳ�ѡ��Ŀ��ͻ��������ť��������ϵͳ�������������ݹ���,�����ѡ��Ŀͻ����������ĸ��˿ͻ�����Ӱݷüƻ������ռƻ����пͻ��ݷá�����ݷý������ٴν���˸��˿ͻ���д�ݷü�¼(�ܽ�)����ɹ�ǩ�����벿�ž���������Ա�ں�ͬ��������Ӽ�¼<br>
          </span><br>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td bgcolor="#88ADDF">&nbsp;</td>
            </tr>
          </table>
          <br>          <span class="style2">ע������:</span><br>
          1.����ͻ�����Ϊ�������˿ͻ���Ϊ���ĸ��˿ͻ�,����ҵ����Ա��������ӻ�ʹ�ô˿ͻ���Ϣ��<br> 
          2.�������Ѵ˿ͻ�ת������ҵ����Ա�ݷ�,��ʹ��ת�����ݹ��ܡ�<br>
          3.�����Բ���д�ݷüƻ�,��������д�ݷü�¼(�ܽ�)��<br>
          4.�ͻ��϶�ʱ,��ʹ�ò�ѯ���ܡ�<br>
          5.����ҵ����Ա�޷��������İݷü�¼�����˿ͻ���Դ��<br>
          6.��ҵ��Ա���޷�ʹ�ô�ϵͳ��</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="364" border="0" cellspacing="0">
            <form name="loginForm" action="?action=login" method="post" onSubmit="return checkInput();">
              <tr> </tr>
              <tr>
                <td height="24" id="errMsg"><div align="center">
                      <% = errMsg %>
                </div></td>
              </tr>
              <tr>
                <td height="5" align="center"><span style="font-size:11px"></span>
                    <input name="item1" type="text" class="box" id="item1" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" value="<%=oabusyusername%>" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="5" align="center"><input name="item2" type="password" class="box" id="item2" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" value="123" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="24" align="center">
                  <input type="submit" name="Submit" value="��������ϵͳ">
&nbsp;&nbsp; </td>
              </tr>
              <tr>
                <td height="24" align="center">&nbsp;</td>
              </tr>
            </form>
          </table>
        </div></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
End Sub
%>
</body>
</html>
