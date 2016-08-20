<%
'����ePASS������Ϣ
Function Return_Error_Info(Error_Code)
	Return_Error_Info=ePass.ErrToStr(Error_Code)
End Function

'ȡ��ePASSӲ����Ϣ
'InfoStr(1)---ePass�̼��汾��
'InfoStr(2)---ePassAPI������汾�Ų�Ʒ����
'InfoStr(3)---ePass��Ʒ����
'InfoStr(4)---ePass�����ռ�����
'InfoStr(5)---ePass�����㷨������
'InfoStr(6)---ePassӲ�����к�
Function Get_Epass_HardWare_Info()
	Dim ctx,InfoStr(6)
	Set ctx = CreateObject("EpsModu.EpsCtx")
	InfoStr(1)=Cstr(ctx.VerMajor)+"."+Cstr(ctx.VerMinor)
	InfoStr(2)=Cstr(ctx.LibVer)
	InfoStr(3)=Cstr(ctx.ProdCode)
	InfoStr(4)=Cstr(ctx.MemSize)
	InfoStr(5)=Cstr(ctx.Capabilities)
	InfoStr(6)=Hex(ctx.SerialNumber(1))+Hex(ctx.SerialNumber(0))
	Set ctx=Nothing
	Get_Epass_HardWare_Info=InfoStr
End Function

'�ڸ�Ŀ¼�´���һ����Ŀ¼
Function Create_Dir(DirID)
	Dim ErrCode
	ErrCode=ePass.CreateDir(DirID)
	Create_Dir=ErrCode
End Function

'ɾ��Ŀ¼
Function Delete_Dir(DirID)
	Dim ErrCode
	ErrCode=ePass.DeleteDir(DirID)
	Delete_Dir=ErrCode
End Function

'ɾ�����е��ļ���Ŀ¼����Ҫ����Ա����Ȩ��
Function Delete_All_Files()
	Dim ErrCode
	ErrCode=ePass.DeleteAllFiles()
	Delete_All_Files=ErrCode
End Function

'ɾ��ָ��Ŀ¼�µ��ļ�
Function Delete_File(DirID,FileID)
	Dim ErrCode
	ErrCode=ePass.DeleteFile(DirID,FileID)
	Delete_File=ErrCode
End Function

'����LEDָʾ�Ƶĵ���״̬
Function Led_Control(LedStatus)
	Dim ErrCode
	Errcode=ePass.LedControl(LedStatus)
	Led_Control=ErrCode
End Function

'����һ���ļ�
'DirId---Ŀ¼ID��
'FileId---�ļ�ID��
'FileSize---�ļ���С
'FileType---�����ļ�����
'ReadType---�ļ�������
'WriteType---�ļ�д����
'CryptType---�ļ���������
Function Create_File(DirId,FileId,FileSize,FileType,ReadType,WriteType,CryptType)
	Dim FileObject,ErrCode
	Set FileObject=CreateObject("EpsModu.EpsFileInfo")
	FileObject.DirId=DirId
	FileObject.Id=FileId
	FileObject.Size=FileSize
	FileObject.Type=FileType
	FileObject.ReadAccess=ReadType
	FileObject.WriteAccess=WriteType
	FileObject.CryptAccess=CryptType
	ErrCode=ePass.CreateFile()
	Set FileObject=Nothing
	Create_File=ErrCode
End Function

'��һ���ļ�
'DirId---Ŀ¼ID��
'FileId---�ļ�ID��
'FileInfo(1)---�������
'FileInfo(2)---�ļ���С����ֵ�ͣ�
'FileInfo(3)---�ļ����ͣ�ʮ�����ƣ�
'FileInfo(4)---�ļ������ԣ��߼��ͣ�
'FileInfo(5)---�ļ�д���ԣ��߼��ͣ�
'FileInfo(6)---�ļ��������ԣ��߼��ͣ�
Function Open_File(DirId,FileId)
	Dim ctx,FileInfo(6),ErrCode
	Set ctx = CreateObject("EpsModu.EpsCtx")
	ErrCode=ePass.SelectFile(DirId,FileId)
	FileInfo(1)=ErrCode
	If ErrCode=0 Then
		FileInfo(2)=ctx.FileSize
		FileInfo(3)=ctx.FileType
		If ctx.FileAccess And EPS_ACCESS_READ Then
			FileInfo(4)=True
		Else
			FileInfo(4)=False
		End If

		If ctx.FileAccess And EPS_ACCESS_WRITE Then
			FileInfo(5)=True
		Else
			FileInfo(5)=False
		End If

		If ctx.FileAccess And EPS_ACCESS_CRYPT Then
			FileInfo(6)=True
		Else
			FileInfo(6)=False
		End If
	End If
	Set ctx=Nothing
	Open_File=FileInfo
End Function

'���ļ�����
'Read_Offset---�����������
'Read_Length---��Ҫ��ȡ�����ݳ���
Function Read_File(Read_Offset,Read_Length)
	Dim ErrCode,i,sBuf,sText,GetInfo(3)
	ErrCode=ePass.ReadFile(Read_Offset,Read_Length)
	If ErrCode=0 Then
		For i=0 To ePass.FileBufLen-1
			sText=sText+Chr(ePass.FileBuf(i))
			sBuf=sBuf+Cstr(ePass.FileBuf(i))
		Next
	End If
	GetInfo(1)=ErrCode
	GetInfo(2)=sBuf
	GetInfo(3)=sText
	Read_File=GetInfo
End Function

'���ļ��ж�������
'Write_Offset---д���������
'Write_Length---��Ҫд������ݵĳ���
Function Write_File(Write_Offset,Write_Length)
	Dim i,ErrCode
	For i=0 to Write_Length
		ePass.FileBuf(i)=i+65
	Next
	ErrCode=ePass.WriteFile(Write_Offset,Write_Length)
	Write_File=ErrCode
End Function

'���ȫ�ֿ�����Ϣ
Function Get_Access_Info()
	Dim ai,ErrCode,AccessInfo(5)
	Set ai=CreateObject("EpsModu.EpsAccessInfo")
	ErrCode=ePass.GetAccessSettings
	AccessInfo(1)=ErrCode
	If ErrCode=0 Then
		AccessInfo(2)=ai.CreateAccess
		AccessInfo(3)=ai.DeleteAccess
		AccessInfo(4)=ai.CurPin
		AccessInfo(5)=ai.MaxPin
	End If
	Set ai=Nothing
	Get_Access_Info=AccessInfo
End Function

'����ȫ�ִ������ԣ��ò�����Ҫ����ԱȨ��
Function Set_Access(Create_Access,Delete_Access,CurPin_Info,MaxPin_Info)
	Dim ai,ErrCode
	Set ai=CreateObject("EpsModu.EpsAccessInfo")
	ai.CreateAccess=Create_Access
	ai.DeleteAccess=Delete_Access
	ai.CurPin=CurPin_Info
	ai.MaxPin=MaxPin_Info
	ErrCode=ePass.SetAccessSettings
	Set ai=Nothing
	Set_Access=ErrCode
End Function

'��֤PIN�룬����0��ʾ�ɹ��������ʾʧ��
Function Check_Pin(Pin_Code)
	Dim ErrCode
	Pin_Code="&h"+Pin_Code
	ErrCode=ePass.VerifyPin(CLng(Pin_Code))
	Check_Pin=ErrCode
End Function

'��֤����Ա���룬����0��ʾ�ɹ��������ʾʧ��
Function Check_Master(Master_Code)
	Dim ErrCode
	ErrCode=ePass.VerifyMasterKey(Master_Code)
	Check_Master=ErrCode
End Function

'�޸�PIN�룬����0��ʾ�ɹ��������ʾʧ��
Function Change_Pin(Pin_Code)
	Dim ErrCode
	Pin_Code="&h"+Pin_Code
	ErrCode=ePass.ModifyPin(CLng(Pin_Code))
	Change_Pin=ErrCode
End Function

'�޸Ĺ���Ա���룬����0��ʾ�ɹ��������ʾʧ��
Function Change_Master(Master_Code)
	Dim ErrCode
	ErrCode=ePass.ModifyMasterKey(Master_Code)
	Change_Master=ErrCode
End Function

'����ePassӦ���з���˼���128λ����ժҪ�Ľӿں�����������̵�Ч�ڿͻ��˵�ePass������Ӳ��
'Text_Str---����ַ���
'Key_Str---Key��Կ��
Function Get_Soft_HmacMd5(Text_Str,Key_Str)
	Dim i,Text_Len,Key_Len,One_Str,Buf
	Text_Len=Len(Text_Str)
	Key_Len=Len(Key_Str)
	For i=1 to Text_Len'����������е��ַ���
		One_Str=Mid(Text_Str,i,1)
		ePass.TextBuf(i-1)=Asc(One_Str)
	Next
	For i=1 to Key_Len'����������е���Կ��
		One_Str=Mid(Key_Str,i,1)
		ePass.KeyBuf(i-1)=Asc(One_Str)
	Next
	ePass.SoftHmacMd5 Text_Len,Key_Len
	For i=0 to 15
		Buf=Buf+Hex(ePass.DigestBuf(i))
	Next
	Get_Soft_HmacMd5=Buf
End Function

'����ePass�������ɵ���Կ�ļ����м���128λ����ժҪ
'KeyId1---key�ļ�1��ID��
'KeyDir2---key�ļ�2����Ŀ¼ID��
'KeyId2---key�ļ�2��ID��
'Key_Str---key��Կ�ַ���
'Text_Str---����ַ���
'Buf(1)---�������
'Buf(2)---128λ����ժҪ
Function Get_HmacMd5(KeyDir1,KeyId1,KeyDir2,KeyId2,Text_Str)
	Dim i,ErrCode,Text_Len,One_Str,Buf(2)
	Text_Len=Len(Text_Str)
	For i=1 to Text_Len
		One_Str=Mid(Text_Str,i,1)
		ePass.TextBuf(i-1)=Asc(One_Str)
	Next
	ErrCode=ePass.HmacMd5(KeyDir1,KeyId1,KeyDir2,KeyId2,Text_Len)
	Buf(1)=ErrCode
	If ErrCode=0 Then
		For i=0 to 15
			Buf(2)=Buf(2)+Hex(ePass.DigestBuf(i))
		Next
	End If
	Get_HmacMd5=Buf
End Function

'������Կ�ļ�
'KeyDir1---key�ļ�1����Ŀ¼ID��
'KeyId1---key�ļ�1��ID��
'KeyDir2---key�ļ�2����Ŀ¼ID��
'KeyId2---key�ļ�2��ID��
'Key_Str---key��Կ�ַ���
'Key_Crypt_Type---key���ܷ�ʽ
Function Create_Key_File(KeyDir1,KeyId1,KeyDir2,KeyId2,Key_Str,Key_Crypt_Type)
	Dim i,ErrCode,Key_Len,One_Str
	Key_Len=Len(Key_Str)
	For i=1 to Key_Len
		One_Str=Mid(Key_Str,i,1)
		ePass.KeyBuf(i-1)=Asc(One_Str)
	Next
	ePass.KeyCryptType=Key_Crypt_Type
	ErrCode=ePass.CreateKeyFile(KeyDir1,KeyId1,KeyDir2,KeyId2,Key_Len)
	Create_Key_File=ErrCode
End Function

'�����豸
Function Reset_Epass()
	Dim ErrCode
	ErrCode=ePass.ResetDevice()
	Reset_Epass=ErrCode
End Function

'ȡ��һ��64λ�����
Function Get_Rand_Number()
	Dim i,sRand
	ePass.GetChallenge
	For i=0 to 7
		sRand=sRand+Hex(ePass.ChallengeBuf(i))
	Next
	Get_Rand_Number=sRand
End Function
%>