<!--#include file="conn.asp"-->
<!--#include file="skin.asp"-->
<!--#include file="languages.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style/style.css" rel="stylesheet" type="text/css">
<link href="style/style<%=skinid%>.css" rel="stylesheet" type="text/css">
<title>选择时区和语言 Language and TimeZone</title>
</head>
<body>
<table width="500" border="0" align="center" cellpadding="3" cellspacing="0" class="tableBorder2">
  <tr> 
    <th height="25" align="center"><b>选择时区和语言 Language and TimeZone</b></th>
  </tr>
  <tr class=tablebody1> 
    <td><br>
      <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="show.asp" method="post" target="_top">
          <input type="hidden" name="siteid" value="1">
          <tr> 
            <td width="150"><p>请选择时区<br>叫匡拒跋<br>
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
            <td>请选择语言<br>
			叫匡拒粂ē<br>
              Please Select Language<br> </td>
            <td><select name="Language">
                <option value="CHS" <%if request.cookies("Language")="CHS" then response.write " selected"%>>简体中文 
                </option>
                <option value="CHT" <%if request.cookies("Language")="CHT" then response.write " selected"%>>いゅ羉蔨 
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