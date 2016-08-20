<%
Function LeftTrue(str,n)
If len(str)<=n/2 Then
LeftTrue=str
Else
Dim TStr
Dim l,t,c
Dim i
l=len(str)
t=l
TStr=""
t=0
for i=1 to l
c=asc(mid(str,i,1))
If c<0 then c=c+65536
If c>255 then
t=t+2
Else
t=t+1
End If
If t>n Then exit for
TStr=TStr&(mid(str,i,1))
next
LeftTrue = TStr+"..."
End If
End Function
function kbbs(stru)
if not isnull(stru) then
	stru = replace(stru, ">", "&gt;")
	stru = replace(stru, "<", "&lt;")
	stru = Replace(stru, CHR(32), " ")
	stru = Replace(stru, CHR(9), "&nbsp;")
	stru = Replace(stru, CHR(34), "&quot;")
	stru = Replace(stru, CHR(39), "&#39;")
	stru = Replace(stru, CHR(13), "")
	stru = Replace(stru, CHR(10), "&nbsp;")
		stru = Replace(stru, "script", "&#115cript")

kbbs = stru
end if
end function

%>