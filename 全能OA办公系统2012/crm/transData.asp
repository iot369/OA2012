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

<!--登录权限判断，Session和MD5加密判断-->
<%
Rem Session("CRM_account") 用户帐号
Rem Session("CRM_name") 用户名
Rem Session("CRM_level") 用户等级

Function getGroupName(gId)
    If gId = "" Then
	    getGroupName = ""
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gId = " & gId,conn,3,1
		If rs.RecordCount <> 1 Then
		    getGroupName = ""
		Else
		    getGroupName = rs("gName")
		End If
		rs.Close
		Set rs = Nothing
	End If
End Function

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Dim strCounter,strToPrint,i
Session("CRM_transFrom") = Trim(Request("transFrom"))

If Session("CRM_transFrom") <> "" Then  Call listData()

Sub listData()
    Dim rs,intTotalRecords,intTotalPages,intCurrentPage,intPageSize
    intCurrentPage = CInt(ABS(Request("pageNum")))
    If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
    intPageSize = 10
	
	Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From baidu_client Where cUser = '" & Session("CRM_transFrom") & "' Order By cId Desc",conn,3,1
    intTotalRecords = rs.RecordCount
    rs.PageSize = intPageSize
    intTotalPages = rs.PageCount
    If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
    If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
	
    strCounter = strCounter & "共 " & intTotalRecords & " 条记录 "
    strCounter = strCounter & "共 " & intTotalPages & " 页 "
    strCounter = strCounter & "当前第 " & intCurrentPage & " 页 "
    If intCurrentPage <> 1 And intTotalRecords <> 0 Then
        strCounter = strCounter & "<a href=""?pageNum=1""><<首页</a> "
    Else
        strCounter = strCounter & "<<首页 "
    End If
    If intCurrentPage > 1 Then
        strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage - 1 & """><上一页</a> "
    Else
        strCounter = strCounter & "<上一页 "
    End If
    If intCurrentPage < intTotalPages Then
        strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage + 1 & """>下一页></a> "
    Else
        strCounter = strCounter & "下一页> "
    End If
    If intCurrentPage <> intTotalPages Then
        strCounter = strCounter & "<a href=""?pageNum=" & intTotalPages & """>尾页>></a>"
    Else
        strCounter = strCounter & "尾页>>"
    End If
	
    'Dim i
    i = 0
    Do While Not rs.BOF And Not rs.EOF
        i = i + 1
    	strToPrint = strToPrint & "        <tr>" & VBCrlf
    	strToPrint = strToPrint & "          <td align=""center"">" & rs("cId") & "</td>" & VBCrlf
    	strToPrint = strToPrint & "          <td><input type=""checkbox"" name=""checkOne"" id=""checkOne"" value=""" & rs("cId") & """></td>" & VBCrlf
	    strToPrint = strToPrint & "        <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
	    strToPrint = strToPrint & "        <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" &  VBCrlf
	    strToPrint = strToPrint & "        <td>" & rs("cType") & "</td>" & VBCrlf
	    strToPrint = strToPrint & "        <td>" & getGroupName(rs("cGroup")) & "</td>" & VBCrlf
    	strToPrint = strToPrint & "        </tr>" & VBCrlf
        If i >= intPageSize Then Exit Do
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
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

function transData(sel)
{
    if(confirm("此操作将转移选中的记录。\r继续操作？") == false){
	    return;
	}
	
	var num = parseInt(document.all.num.value);
	if(num == 0){//记录为空，返回操作
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
	if(flag == 0 && sel != "all"){
	    //noSelect();
		return false;
	}	
}
//没有选择错误提示
function noSelect()
{
    alert("操作前请选择至少一条数据。");
}

function showHideHead(strSrc)
{
	var strFile = strSrc.substring(strSrc.lastIndexOf("/"),strSrc.length);
    if (strFile == "/arrow_up.gif"){
	    oHead.style.display = "none";
		oHeadCtrl.src = "images/arrow_down.gif";
		oHeadCtrl.alt = "显示头部";
		oHeadBar.title = "显示头部";		
	}
	else {
	    oHead.style.display = "block";
		oHeadCtrl.src = "images/arrow_up.gif";
		oHeadCtrl.alt = "隐藏头部";
		oHeadBar.title = "隐藏头部";
	}
}

function ftnTransTo()
{
    var sel;
	for (var i=0;i<document.transToForm.selNum.length;i++){
	    if (document.transToForm.selNum[i].checked == true){
		    sel = document.transToForm.selNum[i].value;
		}
	}
    if (document.transToForm.transTo.value == ""){
	    alert("请输入目标用户。")
		document.transToForm.transTo.focus();
		return false;
	}
	transData(sel)
	if (document.transToForm.arrayId.value == "" && sel != "all"){
	    alert("请选择要转移的数据记录。");
		return false;
	}
}

if (this.location.href == top.location.href){
    top.location.href = "";
}

function checkFrom()
{
    if (document.transFromForm.transFrom.value == ""){
	    alert("请输入要被转移的用户。");
		document.transFromForm.transFrom.focus();
		return false;
	}
}
-->
</script>
<style type="text/css">
.style7 {color: #2d4865}
.style8 {color: #0d79b3;
	font-weight: bold;
}
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
                  <td class="style7">销售系统</td>
                </tr>
            </table></td>
            <td width="1"><span class="style8"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br>
<table width="550" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
     
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td height="5"><img src="images/null.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td bgcolor="#88ADDF">&nbsp;</td>
      </tr>
    </table>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td width="20" align="right">&nbsp;</td>
          <td><table border="0" align="left" cellpadding="0" cellspacing="0">
              <form name="transFromForm" method="post" action="?" onSubmit="return checkFrom();">
                <tr> 
                  <td> 要转移的用户： 
                    <input name="transFrom" type="text" id="transFrom" value="<% If Session("CRM_transFrom") <> "" Then Response.Write(Session("CRM_transFrom")) %>" size="12" maxlength="16" onFocus="this.value='';"> 
                    <input name="Search" type="submit" id="Search" value="列表"></td>
                </tr>
              </form>
            </table>
            <br>
            <br>
            <table border="0" align="left" cellpadding="0" cellspacing="0">
              <form name="transToForm" action="trans.asp" method="post" onSubmit="return ftnTransTo();">
			  <tr> 
                <td align="right"> <label>转移给： 
                  <input name="transTo" type="text" id="transTo" size="12" maxlength="16" onFocus="this.value='';">
                  <input name="selNum" type="radio" value="seled" checked>
                  仅选择的记录</label> <label> 
                  <input type="radio" name="selNum" value="all">
                  该业务员全部记录</label>
                  <input name="transFrom" type="hidden" id="transFrom" value="<% If Session("CRM_transFrom") <> "" Then Response.Write(Session("CRM_transFrom")) %>">
                  <input name="arrayId" type="hidden" id="arrayId" value=""> 
                  <input type="submit" name="Submit" value="转移"> 
                </td>
              </tr>
			  </form>
            </table></td>
          <td width="20">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
      <% = strCounter %>
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr> 
          <td width="60" height="20" align="center" bgcolor="menu">编号</td>
          <td width="20" height="20" align="center" bgcolor="menu">
            <input name="checkAll" type="checkbox" id="checkAll" value="checkbox" onClick="checkAll()"></td>
          <td height="20" align="center" bgcolor="menu">公司名称</td>
          <td height="20" align="center" bgcolor="menu">公司网址</td>
          <td height="20" align="center" bgcolor="menu">客户等级</td>
          <td height="20" align="center" bgcolor="menu">用户组</td>
        </tr>
        <% = strToPrint %>
		<input name="num" type="hidden" id="num" value="<% = i %>">
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
