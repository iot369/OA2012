<%
class cls_show
	Public Sub ShowPageInfo(table,id,condition,PageNo,PageSize,LinkFile)
		dim StrSql,TotalCount,TotalPageCount,OutStr
		StrSql="SELECT count("&id&") FROM "&table&" "&condition&""
		Set rs = hx.Execute(StrSql)
		TotalCount=rs(0)
		Set rs=Nothing

	'�����¼��Ϊ0����ô�˳�
	If TotalCount=0 Then
		Exit Sub
	End If
		OutStr="<P align=Center>"
	'�����¼��>MaxRecord�����¼��ΪMaxRecord
	if TotalCount>MaxRecord then
		OutStr = OutStr & Lang.item("g_022") & TotalCount & Lang.item("g_024") & " " & Lang.item("g_023") & MaxRecord & Lang.item("g_024")
		TotalCount=MaxRecord
	else
		OutStr = OutStr & Lang.item("g_022") & TotalCount & Lang.item("g_024")		
	end if
	'�õ���ҳ��
	If (TotalCount mod PageSize)=0 Then
		TotalPageCount=TotalCount\PageSize
	Else
		TotalPageCount=(TotalCount\PageSize)+1
	End If
	'��ֹ�ύ��page�������ڵڶ����ύ����ҳ��
	if PageNo>TotalPageCount then 
		PageNo=TotalPageCount
	End if
		OutStr = OutStr & "&nbsp;<font color='#FF0000'>"&PageNo&"</font>/<font color='#FF0000'>"&TotalPageCount&"</font>"
	If PageNo>1 Then
		OutStr = OutStr & "&nbsp;<a Href='?"&LinkFile&"&PageNo=1'>"& Lang.item("g_025") & "</a>"
		OutStr = OutStr & "&nbsp;<a Href='?"&LinkFile&"&PageNo="&PageNo-1&"'>"& Lang.item("g_026") & "</a>"
	End If
	If PageNo<TotalPageCount Then
		OutStr = OutStr & "&nbsp;<a Href='?"&LinkFile&"&PageNo="&PageNo+1&"'>"& Lang.item("g_027") & "</a>"
		OutStr = OutStr & "&nbsp;<a Href='?"&LinkFile&"&PageNo="&TotalPageCount&"'>"& Lang.item("g_028") & "</a>"
	End If
		OutStr = OutStr & "</P>"
		Response.Write(OutStr)	
	End Sub

End class
%>