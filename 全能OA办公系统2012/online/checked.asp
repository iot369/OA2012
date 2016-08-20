
<%

function selected(req,reqvalue)
	if req=reqvalue then
		selected=" selected"
	else
		selected=""
	end if
end function

function checked(req,reqvalue)
	if req=reqvalue then
		checked=" checked"
	else
		checked=""
	end if
end function

function checked1(req,reqvalue)
	if req=reqvalue then
		checked1="¡Ñ"
	else
		checked1="¡ð"
	end if
end function

function checked2(req,reqvalue)
	if req=reqvalue then
		checked2="¡Ì"
	else
		checked2="¡õ"
	end if
end function

function checked3(value)
	if value="" then
		checked3="&nbsp;"
	else
		checked3=value
	end if
end function
%>
