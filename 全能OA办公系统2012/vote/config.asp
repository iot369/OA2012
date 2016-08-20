
<%
Dim DbConn
Dim dbPath

Sub OpenDB()
	Set DbConn = Server.CreateObject("ADODB.Connection")
	dbPath = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & server.mappath("moondowner_poll.mdb") 
	DbConn.open dbPath
End Sub

Sub CloseDB()
	DbConn.Close
	Set DbConn = Nothing
End Sub


Function RequestText(text)
	if text<>"" and not IsNull(text) then
		RequestText=Trim(Replace(text,"'",""))
	else
		RequestText=""
	end if
End Function

sub out(text)
	response.write "<br><br><center><font size=2>" & text & "<a href="""&_
	"javascript:history.back();"">[·µ»Ø]</a></font></center>"
	response.end
end sub

Sub CheckPage()
	rs.PageSize = PageNo
	mpage=rs.pagecount
	If Request("p") <> "" and IsNumeric("p") Then
		PageNum = CINT(Request("p"))
	Else
		PageNum = 1
	End If
End Sub 

Sub DisplayPage()
	Response.Write "Ò³´Î£º"
	If PageNum > 1 Then
		Response.Write "<a href=""" _
		& Request.ServerVariables("SCRIPT_NAME") _
		& "?p=" & PageNum - 1 _
		& """>[&lt;&lt;]</a>"
	End If

	For i = 1 To rs.PageCount
		If i = PageNum Then
			Response.Write "<font color=""red"">[" & i & "]</font>"
		else
			Response.Write "<a href=""" _
			& Request.ServerVariables("SCRIPT_NAME") _
			& "?p=" & i _
			& """>[" & i & "]</a>"
		End If
	Next

	If PageNum < rs.PageCount Then
		Response.Write "<a href=""" _
		& Request.ServerVariables("SCRIPT_NAME") _
		& "?p=" & PageNum + 1 _
		& """>[&gt;&gt;]</a>"
	End If
End Sub
%>