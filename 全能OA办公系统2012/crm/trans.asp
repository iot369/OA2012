<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
Response.Buffer = True
Response.Expires = 0
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache" 
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"
%>
<!--#include file="Connections/conn.asp" -->

<!--��¼Ȩ���жϣ�Session��MD5�����ж�-->'
<%
Rem Session("CRM_account") �û��ʺ�
Rem Session("CRM_name") �û���
Rem Session("CRM_level") �û��ȼ�

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Dim errMsg,flag
errMsg = CInt(Abs(Request("errMsg")))
flag = 0

Select Case errMsg
Case 1
    errMsg = "<center><br><br><font color=""#FF0000"">�ύ�����ݲ�������</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""location.replace('transData.asp');""><br><br>"
	''Response.Write(errMsg)
	flag = 1
Case 2
    errMsg = "<center><br><br><font color=""#FF0000"">��ת���û���Ŀ���û���ͬ��</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""location.replace('transData.asp');""><br><br>"
	''Response.Write(errMsg)
	flag = 1
Case 3
    errMsg = "<center><br><br><font color=""#FF0000"">��ת���û���Ŀ���û�<br>������һ�������ڡ�</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""location.replace('transData.asp');""><br><br>"
	''Response.Write(errMsg)
	flag = 1
Case 4
    errMsg = "<center><br><br><font color=""#FF0000"">����ת�����</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""location.replace('transData.asp');""><br><br>"
	''Response.Write(errMsg)
	flag = 1
Case Else
    errMsg = ""
End Select

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="http://www.web87.9126.com">
<meta name="Date" content="2003-08">
<title>���۹���ϵͳ</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function checkAll()
{
    var num = parseInt(document.all.num.value);
	if(num == 0){
	    return;
	}
	if(num == 1){
	    document.all.checkOne.checked = document.all.checkAll.checked;
	}
	else{
	    for(var i=0;i<num;i++){
	        document.all.checkOne[i].checked = document.all.checkAll.checked;
		}
	}
}

function transData()
{
    if(confirm("�˲�����ת��ѡ�еļ�¼��\r����������") == false){
	    return;
	}
	
	var num = parseInt(document.all.num.value);
	if(num == 0){//��¼Ϊ�գ����ز���
	    return;
    }
	var flag = 0
	var arrayId = "";
	if(num == 1){
	    if(document.all.checkOne.checked == true){
		    var arrayId = document.all.checkOne.value;
			document.all.arrayId.value = arrayId;
			flag = 1;
		}
	}
	else{
	    for(var i=0;i<num;i++){
		    if(document.all.checkOne[i].checked == true){
		        if(arrayId == ""){
			        arrayId = document.all.checkOne[i].value;
			    }
			    else{
				    arrayId = arrayId + ",," + document.all.checkOne[i].value;
			    }
			}
		}
		if(arrayId != ""){
			document.all.arrayId.value = arrayId
			flag = 1
		}
	}
	if(flag == 0){
	    noSelect();
	}	
}
//û��ѡ�������ʾ
function noSelect()
{
    alert("����ǰ��ѡ������һ�����ݡ�");
}

function showHideHead(strSrc)
{
	var strFile = strSrc.substring(strSrc.lastIndexOf("/"),strSrc.length);
    if (strFile == "/arrow_up.gif"){
	    oHead.style.display = "none";
		oHeadCtrl.src = "images/arrow_down.gif";
		oHeadCtrl.alt = "��ʾͷ��";
		oHeadBar.title = "��ʾͷ��";		
	}
	else {
	    oHead.style.display = "block";
		oHeadCtrl.src = "images/arrow_up.gif";
		oHeadCtrl.alt = "����ͷ��";
		oHeadBar.title = "����ͷ��";
	}
}

function ftnTransTo()
{
    if (document.transToForm.transTo.value == ""){
	    alert("������Ŀ���û���")
		document.transToForm.transTo.focus();
		return false;
	}
	transData()
}

if (this.location.href == top.location.href){
    top.location.href = "";
}
-->
</script>
</head>

<body >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="778" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="menu">
        <tr> 
          <td align="left" background='images/tab_top_background_runner.gif'> <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
            <tr>
              <td><img src="images/null.gif" width="1" height="1"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('listAll.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr > 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">�鿴��������</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('addData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr>   
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">�������</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>	  
          <table onclick="window.location.replace('advanceSearch.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">�߼�����</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('dataForm.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">���ݱ���</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('exportData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">���ݵ���</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
<% If Session("CRM_level") = 9 Then %>
          <table onclick="window.location.replace('transData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
                <td><img src="images/tab_on_left.gif" width="16" height="24"></td>
                <td background="images/tab_on_middle.gif">����ת��</td>
                <td><img src="images/tab_on_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('manageUser.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
              <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">�û�����</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
              </tr>
            </table>
			<table onclick="window.location.replace('system_level.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
              <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">ϵͳ����</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
              </tr>
            </table>
<% End If %>
            <table onclick="window.location.replace('logout.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="right">
              <tr>    
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">ע��</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
            <table onclick="window.location.reload();" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="right">
              <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">ˢ��</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
              </tr>
            </table></td>
      </tr>
    </table>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td height="5"><img src="images/null.gif" width="1" height="1"></td>
      </tr>
    </table>  
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#999999" id="oHeadBar" style="cursor: hand;" title="����ͷ��" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="����ͷ��" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
      <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
        <tr> 
          <td height="20" align="center"> 
            <% = errMsg %></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#999999"><a href="#top"><img src="images/arrow_up.gif" alt="���ض���" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
<%
If flag = 1 Then Response.End()
Dim selNum,arrayId,transFrom,transTo
selNum = Trim(Request("selNum"))
arrayId = Trim(Request("arrayId"))
transFrom = Trim(Request("transFrom"))
transTo = Trim(Request("transTo"))
arrayId = Replace(arrayId,",,",",")

If selNum = "" Then
    Response.Redirect("?errMsg=1")
Else
    If selNum = "all" Then
	    If transTo = "" Or transFrom = "" Then Response.Redirect("?errMsg=1")
    Else
	    If arrayId = "" Or transTo = "" Or transFrom = "" Then Response.Redirect("?errMsg=1")
	End If
End If


If transFrom = transTo Then Response.Redirect("?errMsg=2")
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_user Where uName In ('" & transFrom & "','" & transTo & "')",conn,3,1
If rs.RecordCount <> 2 Then Response.Redirect("?errMsg=3")
rs.Close

If selNum = "seled" Then
    rs.Open "Select * From baidu_client Where cUser = '" & transFrom & "' And cId In (" & arrayId & ")",conn,3,2
Else
    rs.Open "Select * From baidu_client Where cUser = '" & transFrom & "'",conn,3,2
End If
Do While Not rs.BOF And Not rs.EOF
	rs("cUser") = transTo
	rs.Update
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Response.Redirect("?errMsg=4")
%>