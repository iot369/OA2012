<%
'****************************************************************************************
'������(Program Name): Allyes ������ϴ�����											*
'����(Function)��	1.�������趨�ϴ��ļ���С											*
'					2.�����и�������Fso״̬����Fso��֧��״̬							*
'					3.�������趨�����ļ��ķ�ʽ(0=Ψһ��ʽ,1=����ʽ,2=���Ƿ�ʽ)		*
'����(Author):	Allyes��Mac																*
'����޸�����(The Date for last Modify):2003��6��21��									*
'�汾(Version):	1.003 build 205															*
'�޸�(Modify):	1���������ʾ�ļ���С(Build 204����ΪBuild 205)							*
'				2��������ϴ��ļ���ʽ����(Build 203 ����ΪBuild 204)					*
'����վ��(WebSite):	http://allyes@xfxd.com												*
'																						*
'ʹ�÷�ʽ(Option):																		*
'*���ϴ����ļ����浽path��ָ����Ŀ¼���档										        *
'Formfilefield  �ϴ�����"file"����                                                    *
'Path       Ҫ�����ļ��ķ���������·������ʽΪ��"d:\path\subpath"��"d:\path\subpath\"	*
'MaxSize    �����ϴ��ļ�����󳤶ȣ���KByteΪ��λ										*
'SavType    �����������ļ��ķ�ʽ��														*
'           0   Ψһ�ļ�����ʽ�������ͬ�����Զ�������									*
'           1   ����ʽ�������ͬ�������											*
'           2   ���Ƿ�ʽ�������ͬ���򸲸�ԭ�����ļ�									*
'FsoType	Fso֧��ģʽ																	*
'			0   ��֧��																	*
'			1   ֧��FSO																	*
'****************************************************************************************
Option Explicit

Dim FormData, FormSize, Divider, bCrLf
Dim FixFileExt

FormSize = Request.TotalBytes
FormData = Request.BinaryRead(FormSize)
bCrLf = ChrB(13) & ChrB(10)
Divider = LeftB(FormData, InStrB(FormData, bCrLf) - 1)
FixFileExt="asp|aspx|asa|asax|ascx|ashx|asmx|axd|cdx|cer|config|cs|csproj|licx|rem|resx|shtml|shtm|soap|stm|vb|vbproj|webinfo|cgi|pl|php|phtml|php3"		'����Ϊֻ����Щ�ļ������ϴ�(��"|"�Ÿ�)

Function SaveFile(FormFileField, Path, MaxSize, SavType, FsoType)
	If (SavType=0 or SavType=1) and FsoType=0 then
		SaveFile = "modeError"
		Exit function
	End if

    Dim ObjStream,Allyes_ObjStream
	Dim StartPos
	Dim Strlen, SearchStr
	Dim FileStart, FileLen, FileContent
	Dim Re_SavType
	Dim fnN
    Dim intfnN
	Dim FileExtName
    Dim FixFnN
	Dim intFix
	Dim i

    Set ObjStream = Server.CreateObject("ADODB.Stream")
    Set Allyes_ObjStream = Server.CreateObject("ADODB.Stream")
    ObjStream.Mode = 3
    ObjStream.Type = 1
    Allyes_ObjStream.Mode = 3
    Allyes_ObjStream.Type = 1
    SaveFile = ""
    StartPos = LenB(Divider) + 2
    FormFileField = Chr(34) & FormFileField & Chr(34)
	
	'-----------------------------------���·��------------------------------------
    If Right(Path,1) <> "\" Then		'���Ŀ¼������������
        Path = Path & "\"
    End If
	If FsoType = 1 then					'���֧��FSO���⡣���򲻼��
		CheckPath(path)					'���ָ��Ŀ¼�Ƿ���ڣ���������ڣ������д���
	End if
	'-------------------------------------------------------------------------------
	If len(trim(MaxSize)) = 0 then
		MaxSize=50*1024					'ָ��Ĭ������ϴ��ļ�Ϊ50M
	End if

    Do While StartPos > 0				'��ʼ����ÿ��file�ļ���������
        strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
        SearchStr = MidB(FormData, StartPos, strlen)
        If InStr(bin2str(SearchStr), FormFileField) > 0 Then
            FileName = bin2str(GetFileName(SearchStr,path,SavType,FsoType))
            
            ''----------------�ļ���ʽ����------------------------
            fnN = split(fileName,".")
            intfnN = Ubound(fnN)
            FileExtName = trim(fnN(intfnN))
            FixFnN = Split(FixFileExt,"|")
			intFix = Ubound(FixFnN)
			for i = 0 to intFix
				if lcase(FileExtName) = lcase(trim(FixFnN(i))) then
					SaveFile = "fileError"
					exit do
				end if
			next
            '------------------------------------------------------
            
            If FileName <> "" Then
                FileStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
                FileLen = InStrB(StartPos, FormData, Divider) - 2 - FileStart
                If FileLen <= MaxSize*1024 Then
                    FileContent = MidB(FormData, FileStart, FileLen)
					Allyes_ObjStream.Open
					With ObjStream
						.Open
						.Write FormData
						.Position=FileStart-1
						.CopyTo Allyes_ObjStream,FileLen
					End With

					Re_SavType = SavType
                    If SavType = 0 Then
                        SavType = 1
                    End If

                    On error resume next
					Allyes_ObjStream.SaveToFile Path & FileName, SavType
					if err.number<>0 then
						If Re_SavType=0 or Re_SavType=2 then
							FileName="pathError"
						else
							FileName="refileError"
						end if
					end if
                    ObjStream.Close
                    Allyes_ObjStream.Close

					If SaveFile <> "" Then
                        SaveFile = SaveFile & ","  & FileName &"|"& FileLen
                    Else
                        SaveFile = FileName &"|"& FileLen
                    End If
                Else
                    If SaveFile <> "" Then
                        SaveFile = SaveFile & ",refileError"
                    Else
                        SaveFile = "sizeError"
                    End If
                End If
            End If
        End If
        If InStrB(StartPos, FormData, Divider) < 1 Then
            Exit Do
        End If
        StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
    Loop
End Function

Function GetFormVal(FormName)						'ȡ������Ǳ���Ŀ�Ĺ���
	Dim StartPos
	Dim Strlen, SearchStr
	Dim ValStart, ValLen, ValContent

    GetFormVal = ""
    StartPos = LenB(Divider) + 2
    FormName = Chr(34) & FormName & Chr(34)
    Do While StartPos > 0
        Strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
        SearchStr = MidB(FormData, StartPos, strlen)
        If InStr(bin2str(SearchStr), FormName) > 0 Then
               ValStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
               ValLen = InStrB(StartPos, FormData, Divider) - 2 - ValStart
                  ValContent = MidB(FormData, ValStart, ValLen)
               If GetFormVal <> "" Then
                GetFormVal = GetFormVal & "," & bin2str(ValContent)
            Else
                GetFormVal = bin2str(ValContent)
            End If
        End If
        If InStrB(StartPos, FormData, Divider) < 1 Then
            Exit Do
        End If
        StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
    Loop
End Function

Function bin2str(binstr)
	Dim BytesStream,StringReturn

	Set BytesStream = Server.CreateObject("ADODB.Stream")
		With BytesStream
			.Type = 2
			.Open
			.WriteText binstr
			.Position = 0
			.Charset = "GB2312"
			.Position = 2
			StringReturn = .ReadText
			.close
		End With
		Set BytesStream = Nothing
	bin2str = StringReturn
End Function


Function str2bin(str)
	Dim i
    For i = 1 To Len(str)
        str2bin = str2bin & ChrB(Asc(Mid(str, i, 1)))
    Next
End Function

Function GetFileName(str,path,savtype,fsotype)
	Dim fs
	Dim i
	Dim hFileName
	Dim rFileName

    str = RightB(str,LenB(str)-InstrB(str,str2bin("filename="))-9)
    GetFileName = ""
    FileName = ""
    For i = LenB(str) To 1 Step -1
        If MidB(str, i, 1) = ChrB(Asc("\")) Then
            FileName = MidB(str, i + 1, LenB(str) - i - 1)
            Exit For
        End If
    Next
	
	If fsotype=1 then									'���֧��FSO,��ִ��FSO����
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		If savtype = 0 and fs.FileExists(path & bin2str(FileName)) = True Then
			hFileName = FileName
			rFileName = ""
			For i = LenB(FileName) To 1 Step -1
				If MidB(FileName, i, 1) = ChrB(Asc(".")) Then
					hFileName = LeftB(FileName, i-1)
					rFileName = RightB(FileName, LenB(FileName)-i+1)
					Exit For
				End If
			Next
			For i = 0 to 9999 
				If fs.FileExists(path & bin2str(hFileName) & i & bin2str(rFileName)) = False Then
					FileName = hFileName & str2bin(i) & rFileName
					Exit For
				End If
			Next
		End If
		Set fs = Nothing
	End If
	GetFileName = FileName
End Function

Function CheckPath(path)								'����Ŀ¼�Ƿ���ڣ���������ڣ�������Ŀ¼
	Dim Fs
	set Fs=server.CreateObject("scripting.filesystemobject")
	if not fs.FolderExists(path) then
		Fs.CreateFolder(path)
	end if
	set Fs = nothing
End function
%>