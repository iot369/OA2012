<%
Dim SelectedTimeZone,SelectedLanguage

if request.form("timeZone")<>"" then
	SelectedTimeZone = request.form("timeZone")
	response.cookies("TimeZone") = SelectedTimeZone
else
	if request.cookies("TimeZone")="" then
	SelectedTimeZone = TimeZone
	response.cookies("TimeZone") = TimeZone
	else
	SelectedTimeZone = request.cookies("TimeZone")
	end if
end if

if request.form("Language")<>"" then
	SelectedLanguage = request.form("Language")
	response.cookies("Language") = SelectedLanguage
else
	if request.cookies("Language") = "" then
	SelectedLanguage = Language
	response.cookies("Language") = SelectedLanguage
	else
	SelectedLanguage = request.cookies("Language")
	end if
end if

Dim Lang
Set Lang = CreateObject("Scripting.Dictionary")
select case SelectedLanguage
case "CHS"
'Chinese Simplified
%>
<!-- #include file="language/CHS.asp"-->
<%
case "CHT"
'Chinese Traditional 
%>
<!-- #include file="language/CHT.asp"-->
<%
case "ENG"
'English
%>
<!-- #include file="language/ENG.asp"-->
<%

end select


function clearLanguage()
	on error resume next
	clearLanguage 		= Lang.removeAll   
	set clearLanguage 	= nothing
end function
%>