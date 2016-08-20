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
checked1="กั"
else
checked1="ก๐"
end if
end function

function checked2(req,reqvalue)
if req=reqvalue then
checked2="กฬ"
else
checked2="ก๕"
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