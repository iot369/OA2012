<!--#include file="conn.asp"-->
<!--#include file="skin.asp"-->
<!--#include file="languages.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style/style.css" rel="stylesheet" type="text/css">
<link href="style/style<%=skinid%>.css" rel="stylesheet" type="text/css">
<title>ѡ��ʱ�������� Language and TimeZone</title>
</head>
<body>
<table width="500" border="0" align="center" cellpadding="3" cellspacing="0" class="tableBorder2">
  <tr> 
    <th height="25" align="center"><b>ѡ��ʱ�������� Language and TimeZone</b></th>
  </tr>
  <tr class=tablebody1> 
    <td><br>
      <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="show.asp" method="post" target="_top">
          <input type="hidden" name="siteid" value="1">
          <tr> 
            <td width="150"><p>��ѡ��ʱ��<br>�п�ܮɰ�<br>
                Please Select TimeZone</p>
              <p>&nbsp; </p></td>
            <td><p> 
                <select name="TimeZone">
                  <%
		  dim i
		  for i=-12 to 12
			  response.write "<option value="""&i&""""
			  if cint(request.cookies("TimeZone")) = cint(i) then
			  response.write " selected"
			  end if
			  response.write ">"&i&"</option>"  
		  next%>
                </select>
              </p>
              <p>&nbsp; </p></td>
          </tr>
          <tr> 
            <td>��ѡ������<br>
			�п�ܻy��<br>
              Please Select Language<br> </td>
            <td><select name="Language">
                <option value="CHS" <%if request.cookies("Language")="CHS" then response.write " selected"%>>�������� 
                </option>
                <option value="CHT" <%if request.cookies("Language")="CHT" then response.write " selected"%>>�����c�^ 
                </option>
                <option value="ENG" <%if request.cookies("Language")="ENG" then response.write " selected"%>>English 
                </option>
              </select></td>
          </tr>
          <tr> 
            <td height="50" colspan="2"><input type="submit" name="Submit" value="OK"></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<%
hx.ShowFooter
set hx=nothing
%>
</body>
</html>