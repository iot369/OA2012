<%
Dim conn,connstr
Dim DBPath
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("database/db.mdb")
'	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath

''====================================''
''说明    字符替换
''返回值  被替换后的字符串                  
''====================================''
Function htmlEncode2(str)
    If IsEmpty(str) Or str = "" Then
	    htmlEncode2 = "&nbsp;"
	Else
	    str = Replace(str,">","&gt;")
		str = Replace(str,"<","&lt;")
		str = Replace(str,"'","&quot;")
		str = Replace(str,Chr(13),"<br>")
		str = Replace(str,VBCrlf,"<br>")
		str = Replace(str," ","&nbsp;")
		htmlEncode2 = str
	End If
End Function
''====================================''
''说明    字符替换
''返回值  还原字符串                  
''====================================''
Function htmlEncode3(str)
	str = Replace(str,"&quot;","'")
	str = Replace(str,"<br>",Chr(13))
	str = Replace(str,"&nbsp;"," ")
	str = Replace(str,"&gt;",">")
	str = Replace(str,"&lt;","<")
	htmlEncode3 = str
End Function
%>
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
