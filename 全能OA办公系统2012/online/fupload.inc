<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
'Sample multiple binary files upload via ASP - upload include 
'c1997-1999 Antonin Foller, DeskOffice, http://www.DeskOffice.com
'The file is part of ScriptUtilities library
'The file enables http upload to ASP without any components.
'But there is a small problem - ASP does not allow save binary data to the disk.
' So you can use the upload for :
' 1. Upload small text (or HTML) files to server-side disk (Save the data by filesystem object)
' 2. Upload binary/text files of any size to server-side database (RS("BinField") = Upload("FormField").Value


'Limit of upload size
Dim UploadSizeLimit

'********************************** GetUpload **********************************
'This function reads all form fields from binary input and returns it as a dictionary object.
'The dictionary object containing form fields. Each form field is represented by six values :
'.Name name of the form field (<Input Name="..." Type="File,...">)
'.ContentDisposition = Content-Disposition of the form field
'.FileName = Source file name for <input type=file>
'.ContentType = Content-Type for <input type=file>
'.Value = Binary value of the source field. 
'.Length = Len of the binary data field
Function GetUpload()
  Dim Result
  Set Result = Nothing
  If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 'Request method must be "POST"
    Dim CT, PosB, Boundary, Length, PosE
    CT = Request.ServerVariables("HTTP_Content_Type") 'reads Content-Type header
    If LCase(Left(CT, 19)) = "multipart/form-data" Then 'Content-Type header must be "multipart/form-data"
      'This is upload request.
      'Get the boundary and length from Content-Type header
      PosB = InStr(LCase(CT), "boundary=") 'Finds boundary
      If PosB > 0 Then Boundary = Mid(CT, PosB + 9) 'Separetes boundary
      Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
      if "" & UploadSizeLimit<>"" then
        UploadSizeLimit = clng(UploadSizeLimit)
        if Length > UploadSizeLimit then 
'          on error resume next 'Clears the input buffer
'            response.AddHeader "Connection", "Close"
'          on error goto 0
          Request.BinaryRead(Length)
          'Err.Raise 2, "GetUpload", "Upload size " & FormatNumber(Length,0) & "B exceeds limit of " & FormatNumber(UploadSizeLimit,0) & "B"
          'Err.Raise 6, "GetUpload", "Upload File too Long!"
		  set GetUpload=nothing
		  exit function
        end if
      end if
      
      If Length > 0 And Boundary <> "" Then 'Are there required informations about upload ?
        Boundary = "--" & Boundary
        Dim Head, Binary
        Binary = Request.BinaryRead(Length) 'Reads binary data from client
        
        'Retrieves the upload fields from binary data
        Set Result = SeparateFields(Binary, Boundary)
        Binary = Empty 'Clear variables
      Else
        Err.Raise 10, "GetUpload", "Zero length request ."
      End If
    Else
      Err.Raise 11, "GetUpload", "No file sent."
    End If
  Else
    Err.Raise 1, "GetUpload", "Bad request method."
  End If
  Set GetUpload = Result
End Function

'********************************** SeparateFields **********************************
'This function retrieves the upload fields from binary data and retuns the fields as array
'Binary is safearray of all raw binary data from input.
Function SeparateFields(Binary, Boundary)
  Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary
  Dim Fields
  Boundary = StringToBinary(Boundary)

    PosOpenBoundary = InstrB(Binary, Boundary)
    PosCloseBoundary = InstrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

  Set Fields = CreateObject("Scripting.Dictionary")

  Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)
    'Header and file/source field data
    Dim HeaderContent, FieldContent
    'Header fields
    Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
    'Helping variables
    Dim Field, TwoCharsAfterEndBoundary
    'Get end of header
        PosEndOfHeader = InstrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

    'Separates field header
        HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
        
    'Separates field content
        FieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)

    'Separates header fields from header
    GetHeadFields BinaryToString(HeaderContent), Content_Disposition, FormFieldName, SourceFileName, Content_Type

    'Create one field and assign parameters
    Set Field = CreateUploadField()
    Field.Name = FormFieldName
    Field.ContentDisposition = Content_Disposition
    Field.FilePath = SourceFileName
    Field.FileName = GetFileName(SourceFileName)
    Field.ContentType = Content_Type
    Field.Value = FieldContent
    Field.Length = LenB(FieldContent)


    Fields.Add FormFieldName, Field

    'Is this ending boundary ?
    TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
        'Binary.Mid(PosCloseBoundary + Len(Boundary), 2).String
    isLastBoundary = TwoCharsAfterEndBoundary = "--"
    If Not isLastBoundary Then 'This is not ending boundary - go to next form field.
      PosOpenBoundary = PosCloseBoundary
            PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary )
    End If
  Loop
  Set SeparateFields = Fields
End Function

'********************************** Utilities **********************************
Function BinaryToString(str)
strto = ""
for i=1 to lenb(str)
if AscB(MidB(str, i, 1)) > 127 then
        strto = strto & chr(Ascb(MidB(str, i, 1))*256+Ascb(MidB(str, i+1, 1)))
        i = i + 1
     else
        strto = strto & Chr(AscB(MidB(str, i, 1)))
     end if
next
BinaryToString=strto
End Function

Function StringToBinary(String)
    Dim I, B
    For I=1 to len(String)
        B = B & ChrB(Asc(Mid(String,I,1)))
    Next 
    StringToBinary = B
End Function

'Separates header fields from upload header
Function GetHeadFields(ByVal Head, Content_Disposition, Name, FileName, Content_Type)
  Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
  Name = (SeparateField(Head, "name=", ";")) 'ltrim
  If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)
  FileName = (SeparateField(Head, "filename=", ";")) 'ltrim
  If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)
  Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

'Separets one filed between sStart and sEnd
Function SeparateField(From, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE, sFrom
  sFrom = LCase(From)
  PosB = InStr(sFrom, sStart)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(From, PosB, PosE - PosB)
  Else
    SeparateField = Empty
  End If
End Function

'Separetes file name from the full path of file
Function GetFileName(FullPath)
  Dim Pos, PosF
  PosF = 0
  For Pos = Len(FullPath) To 1 Step -1
    Select Case Mid(FullPath, Pos, 1)
      Case "/", "\": PosF = Pos + 1: Pos = 0
    End Select
  Next
  If PosF = 0 Then PosF = 1
  GetFileName = Mid(FullPath, PosF)
End Function
</SCRIPT>
<SCRIPT RUNAT=SERVER LANGUAGE=JSCRIPT>
//The function creates Field object.
function CreateUploadField(){ return new uf_Init() }
function uf_Init(){
  this.Name = null
  this.ContentDisposition = null
  this.FileName = null
  this.FilePath = null
  this.ContentType = null
  this.Value = null
  this.Length = null
}
</SCRIPT>