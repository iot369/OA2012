<!-- #include file=conn.asp -->
<%
Dim style
style=Request("style")
%>
document.write("<script>var url='<%=hx.baseurl%>';var style='<%=style%>';</script>");
document.write("<script src="+url+"/stat.asp?style="+style+"&referer="+escape(document.referrer)+"&screenwidth="+(screen.width)+"></script>");


