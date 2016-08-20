<%
'返回ePASS错误信息
Function Return_Error_Info(Error_Code)
	Return_Error_Info=ePass.ErrToStr(Error_Code)
End Function

'取得ePASS硬件信息
'InfoStr(1)---ePass固件版本号
'InfoStr(2)---ePassAPI函数库版本号产品批号
'InfoStr(3)---ePass产品批号
'InfoStr(4)---ePass存贮空间总量
'InfoStr(5)---ePass内置算法的类型
'InfoStr(6)---ePass硬件序列号
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

'在根目录下创建一个子目录
Function Create_Dir(DirID)
	Dim ErrCode
	ErrCode=ePass.CreateDir(DirID)
	Create_Dir=ErrCode
End Function

'删除目录
Function Delete_Dir(DirID)
	Dim ErrCode
	ErrCode=ePass.DeleteDir(DirID)
	Delete_Dir=ErrCode
End Function

'删除所有的文件和目录，需要管理员级别权限
Function Delete_All_Files()
	Dim ErrCode
	ErrCode=ePass.DeleteAllFiles()
	Delete_All_Files=ErrCode
End Function

'删除指定目录下的文件
Function Delete_File(DirID,FileID)
	Dim ErrCode
	ErrCode=ePass.DeleteFile(DirID,FileID)
	Delete_File=ErrCode
End Function

'控制LED指示灯的点亮状态
Function Led_Control(LedStatus)
	Dim ErrCode
	Errcode=ePass.LedControl(LedStatus)
	Led_Control=ErrCode
End Function

'建立一个文件
'DirId---目录ID号
'FileId---文件ID号
'FileSize---文件大小
'FileType---创建文件类型
'ReadType---文件读属性
'WriteType---文件写属性
'CryptType---文件加密属性
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

'打开一个文件
'DirId---目录ID号
'FileId---文件ID号
'FileInfo(1)---错误代码
'FileInfo(2)---文件大小（数值型）
'FileInfo(3)---文件类型（十六进制）
'FileInfo(4)---文件读属性（逻辑型）
'FileInfo(5)---文件写属性（逻辑型）
'FileInfo(6)---文件加密属性（逻辑型）
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

'读文件内容
'Read_Offset---读操作的起点
'Read_Length---所要读取的数据长度
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

'向文件中定入内容
'Write_Offset---写操作的起点
'Write_Length---所要写入的内容的长度
Function Write_File(Write_Offset,Write_Length)
	Dim i,ErrCode
	For i=0 to Write_Length
		ePass.FileBuf(i)=i+65
	Next
	ErrCode=ePass.WriteFile(Write_Offset,Write_Length)
	Write_File=ErrCode
End Function

'获得全局控制信息
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

'设置全局创建属性，该操作需要管理员权限
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

'验证PIN码，返回0表示成功，否则表示失败
Function Check_Pin(Pin_Code)
	Dim ErrCode
	Pin_Code="&h"+Pin_Code
	ErrCode=ePass.VerifyPin(CLng(Pin_Code))
	Check_Pin=ErrCode
End Function

'验证管理员密码，返回0表示成功，否则表示失败
Function Check_Master(Master_Code)
	Dim ErrCode
	ErrCode=ePass.VerifyMasterKey(Master_Code)
	Check_Master=ErrCode
End Function

'修改PIN码，返回0表示成功，否则表示失败
Function Change_Pin(Pin_Code)
	Dim ErrCode
	Pin_Code="&h"+Pin_Code
	ErrCode=ePass.ModifyPin(CLng(Pin_Code))
	Change_Pin=ErrCode
End Function

'修改管理员密码，返回0表示成功，否则表示失败
Function Change_Master(Master_Code)
	Dim ErrCode
	ErrCode=ePass.ModifyMasterKey(Master_Code)
	Change_Master=ErrCode
End Function

'用于ePass应用中服务端计算128位数据摘要的接口函数，计算过程等效于客户端的ePass，无需硬件
'Text_Str---随机字符串
'Key_Str---Key密钥串
Function Get_Soft_HmacMd5(Text_Str,Key_Str)
	Dim i,Text_Len,Key_Len,One_Str,Buf
	Text_Len=Len(Text_Str)
	Key_Len=Len(Key_Str)
	For i=1 to Text_Len'输入计算所有的字符串
		One_Str=Mid(Text_Str,i,1)
		ePass.TextBuf(i-1)=Asc(One_Str)
	Next
	For i=1 to Key_Len'输入计算所有的密钥串
		One_Str=Mid(Key_Str,i,1)
		ePass.KeyBuf(i-1)=Asc(One_Str)
	Next
	ePass.SoftHmacMd5 Text_Len,Key_Len
	For i=0 to 15
		Buf=Buf+Hex(ePass.DigestBuf(i))
	Next
	Get_Soft_HmacMd5=Buf
End Function

'利用ePass中已生成的密钥文件进行计算128位数据摘要
'KeyId1---key文件1的ID号
'KeyDir2---key文件2所在目录ID号
'KeyId2---key文件2的ID号
'Key_Str---key密钥字符串
'Text_Str---随机字符串
'Buf(1)---错误代码
'Buf(2)---128位数据摘要
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

'建立密钥文件
'KeyDir1---key文件1所在目录ID号
'KeyId1---key文件1的ID号
'KeyDir2---key文件2所在目录ID号
'KeyId2---key文件2的ID号
'Key_Str---key密钥字符串
'Key_Crypt_Type---key加密方式
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

'重置设备
Function Reset_Epass()
	Dim ErrCode
	ErrCode=ePass.ResetDevice()
	Reset_Epass=ErrCode
End Function

'取得一个64位随机数
Function Get_Rand_Number()
	Dim i,sRand
	ePass.GetChallenge
	For i=0 to 7
		sRand=sRand+Hex(ePass.ChallengeBuf(i))
	Next
	Get_Rand_Number=sRand
End Function
%>