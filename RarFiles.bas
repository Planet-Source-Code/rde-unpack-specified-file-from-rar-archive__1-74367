Attribute VB_Name = "mRarFiles"
Option Explicit ' Wrapper for unrar.dll freely available from www.rarlab.com

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'  32-bit Windows dynamic-link library providing file extraction from RAR archives
' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Private Enum ERAR
   ERAR_SUCCESS = 0          ' Success
   ERAR_NO_COMMENTS = 0      ' Comments not present
   ERAR_COMMENTS_READ = 1    ' Comments read completely
   ERAR_END_ARCHIVE = 10     ' End of archive
   ERAR_NO_MEMORY = 11       ' Not enough memory
   ERAR_BAD_DATA = 12        ' Archive header broken, file CRC error, broken comment
   ERAR_BAD_ARCHIVE = 13     ' File is not valid RAR archive
   ERAR_UNKNOWN_FORMAT = 14  ' Unknown archive format, unknown comment format
   ERAR_EOPEN = 15           ' Volume open error, file open error
   ERAR_ECREATE = 16         ' File create error
   ERAR_ECLOSE = 17          ' Archive close error
   ERAR_EREAD = 18           ' Read error
   ERAR_EWRITE = 19          ' Write error
   ERAR_SMALL_BUF = 20       ' Buffer too small, comments not completely read
   ERAR_UNKNOWN = 21         ' All errors which do not have a special ERAR code
End Enum

Private Enum eOpenMode ' RAROpenArchiveData struct OpenMode flags
   RAR_OM_LIST = 0     ' Open archive for reading file headers only
   RAR_OM_EXTRACT = 1  ' Open archive for testing and extracting files
End Enum

Private Enum eProcessFileOp
   RAR_SKIP = 0     ' Move to next file in archive. If the archive is solid and RAR_OM_EXTRACT mode was set when the archive was opened, the current file will be processed - the operation will be performed slower than a simple seek
   RAR_TEST = 1     ' Test current file and move to next file in the archive. If the archive was opened with RAR_OM_LIST mode, the operation is equal to RAR_SKIP
   RAR_EXTRACT = 2  ' Extract current file and move to next file. If the archive was opened with RAR_OM_LIST mode, the operation is equal to RAR_SKIP
End Enum

Private Enum UCM        ' Callback messages
   UCM_CHANGEVOLUME = 0 ' Process volume change
   UCM_PROCESSDATA = 1  ' Process unpacked data
   UCM_NEEDPASSWORD = 2 ' DLL needs a password to process archive
End Enum

Private Enum RAR_VOL
   RAR_VOL_ASK = 0    ' Required volume is absent. The callback function should prompt user
   RAR_VOL_NOTIFY = 1 ' Required volume is successfully opened. This is a notification call and volume name modification is not allowed
End Enum

Private Type RAROpenArchiveData
   ArcName As String   ' Input  - NULL terminated string containing the archive name
   OpenMode As Long    ' Input  - eOpenMode parameter - RAR_OM_LIST, RAR_OM_EXTRACT
   OpenResult As Long  ' Output - ERAR_SUCCESS, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_BAD_ARCHIVE, ERAR_UNKNOWN_FORMAT, or ERAR_EOPEN
   CmtBuf As String    ' Input  - Buffer for archive comments. Maximum size is limited to 64Kb. Comment text is NULL terminated. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long  ' Input  - Size of buffer for archive comments
   CmtSize As Long     ' Output - Size of comments actually read into the buffer, cannot exceed CmtBufSize
   CmtState As Long    ' Output - ERAR_NO_COMMENTS, ERAR_COMMENTS_READ, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_UNKNOWN_FORMAT, ERAR_SMALL_BUF
End Type

Private Type RARHeaderData
   ArcName As String * 260  ' Output - NULL terminated current archive name. May be used to determine the current volume name
   FileName As String * 260 ' Output - NULL terminated file name in OEM (DOS) encoding
   flags As Long            ' Output - Contains eFileFlags file flags
   PackSize As Long         ' Output - Packed file size or size of the file part if file was split between volumes
   UnpSize As Long          ' Output - Unpacked file size
   HostOS As Long           ' Output - eHost operating system used for archiving
   FileCRC As Long          ' Output - Unpacked file CRC. It should not be used for file parts which were split between volumes
   iFileTime As Integer '}  ' Output - Contains 16-bit time in standard MS DOS format
   iFileDate As Integer '}  ' Output - Contains 16-bit date in standard MS DOS format
   UnpVer As Long           ' Output - RAR version needed to extract file. It is encoded as 10 * Major version + minor version
   Method As Long           ' Output - Packing method
   FileAttr As Long         ' Output - File attributes
   CmtBuf As String '*      ' Input  - Buffer for file comments. Maximum size is limited to 64Kb. Comment text is a NULL terminated string in OEM encoding. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long       ' Input  - Size of buffer for archive comments
   CmtSize As Long          ' Output - Size of comments actually read into the buffer, will not exceed CmtBufSize
   CmtState As Long         ' Output - Comment state
End Type

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' RAROpenArchive
'  Open RAR archive and allocate memory structures.
'  ArchiveData points to RAROpenArchiveData structure.
'  Returns archive handle or NULL on error.

Private Declare Function RAROpenArchive Lib "unrar" (ArchiveData As RAROpenArchiveData) As Long 'Handle

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' RARCloseArchive
'  Close RAR archive and release allocated memory. It must be called when archive
'  processing is finished, even if the archive processing was stopped due to an error.
'  hArcData contains the archive handle obtained from the RAROpenArchive function call.
'  Returns ERAR_SUCCESS on success or ERAR_ECLOSE archive close error.

Private Declare Function RARCloseArchive Lib "unrar" (ByVal hArcData As Long) As Long

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' RARReadHeader
'  Read header of file in archive.
'  hArcData contains the archive handle obtained from RAROpenArchive.
'  HeaderData points to RARHeaderData structure.
'  Returns ERAR_SUCCESS on success, ERAR_END_ARCHIVE end of archive, or ERAR_BAD_DATA file header broken.

Private Declare Function RARReadHeader Lib "unrar" (ByVal hArcData As Long, HeaderData As RARHeaderData) As Long

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' RARProcessFile
'  Performs action and moves the current position in the archive to the next file.
'  Extract or test the current file from the archive opened in RAR_OM_EXTRACT mode.
'  If the mode RAR_OM_LIST is set, then a call to this function will simply skip
'  the archive position to the next file.
'  hArcData contains the archive handle obtained from RAROpenArchive.
'  Operation - File eProcessFileOp operation, RAR_SKIP, RAR_TEST, RAR_EXTRACT.
'  DestPath - Destination extract directory. If DestPath is NULL extracts to the current directory. This parameter has meaning only if DestName is NULL.
'  DestName - Full path and name of the file to be extracted or NULL. If DestName is defined (not NULL) it overrides the original file name saved in the archive and DestPath setting.
'  Both DestPath and DestName must be in OEM encoded. If necessary, use CharToOem to convert text to OEM before passing to this function.
'  Returns ERAR_SUCCESS, ERAR_BAD_DATA, ERAR_BAD_ARCHIVE, ERAR_UNKNOWN_FORMAT, ERAR_EOPEN, ERAR_ECREATE, ERAR_ECLOSE, ERAR_EREAD, or ERAR_EWRITE.
'  Note: if you wish to cancel extraction, return -1 when processing UCM_PROCESSDATA callback message.

Private Declare Function RARProcessFile Lib "unrar" (ByVal hArcData As Long, ByVal Operation As Long, ByVal pDestPath As String, ByVal pDestName As String) As Long

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' RARSetCallback
'  Set a user-defined callback function to process Unrar events.
'  hArcData contains the archive handle obtained from RAROpenArchive.
'  CallbackProc points to a user-defined callback function (AddressOf_Callbackfunc).
'  UserData - User data passed to callback function (ByRef param, ByVal ObjPtr_Me).
'  Other functions of UNRAR.DLL should not be called from the callback function.
'  This subroutine has no return value.

Private Declare Sub RARSetCallback Lib "unrar" (ByVal hArcData As Long, ByVal CallbackProc As Long, ByVal UserData As Long)

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Public Type tRarFiles
    sSpec As String
    vDate As Date
    CRC32 As Long
    CSize As Long
    USize As Long
    Attrb As Long
End Type

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' Use CharToOem to convert text to OEM before passing to RARProcessFile
Private Declare Sub CharToOem Lib "user32" Alias "CharToOemA" (ByVal sSrc As String, ByVal sDest As String)
Private Const MAX_PATH = 260&

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Private Declare Function DosDateTimeToVariantTime Lib "oleaut32" (ByVal lpFatDate As Long, ByVal lpFatTime As Long, vTime As Date) As Long
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
Private m_Size As Long

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Public Function GetRarFileInfo(sRarFile As String, tFiles() As tRarFiles) As Long
   On Error GoTo ErrHandler ' Return file count on success, zero on error

   Dim tRar As RAROpenArchiveData
   Dim tHeader As RARHeaderData

   Dim hRar As Long
   Dim cnt As Long
   Dim rc As Long

   tRar.OpenMode = RAR_OM_LIST ' Fill the RAR header structure
   tRar.ArcName = sRarFile
   tRar.CmtBufSize = 0&

   hRar = RAROpenArchive(tRar) ' Open the archive
   If tRar.OpenResult <> ERAR_SUCCESS Then Exit Function

   ReDim tFiles(50) As tRarFiles
   rc = RARReadHeader(hRar, tHeader) ' Read first file header

   Do While rc = ERAR_SUCCESS
      If cnt > UBound(tFiles) Then
         ReDim Preserve tFiles(cnt + 50) As tRarFiles
      End If

      With tHeader  ' Get the current file info
         tFiles(cnt).sSpec = Left$(.FileName, InStr(.FileName, vbNullChar) - 1&)
         DosDateTimeToVariantTime .iFileDate, .iFileTime, tFiles(cnt).vDate
         tFiles(cnt).CRC32 = .FileCRC
         tFiles(cnt).CSize = .PackSize
         tFiles(cnt).USize = .UnpSize
         tFiles(cnt).Attrb = .FileAttr ' Includes folders
      End With
      cnt = cnt + 1&

      rc = RARProcessFile(hRar, RAR_SKIP, vbNullString, vbNullString)
      If rc <> ERAR_SUCCESS Then Exit Do

      rc = RARReadHeader(hRar, tHeader) ' Read next file header
   Loop

ErrHandler:
   RARCloseArchive hRar
   If cnt Then
      ReDim Preserve tFiles(cnt - 1) As tRarFiles ' Zero based
   Else
      Erase tFiles
   End If
   GetRarFileInfo = cnt
End Function

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Public Function UnpackFile(sRarFile As String, sRelFile As String, sExtractDir As String) As Long
   On Error GoTo ErrHandler ' Return -1 on success, zero on error

   Dim tRar As RAROpenArchiveData
   Dim tHeader As RARHeaderData

   Dim sFile As String
   Dim sDir As String
   Dim hRar As Long
   Dim rc As Long

   tRar.OpenMode = RAR_OM_EXTRACT ' Fill the RAR header structure
   tRar.ArcName = sRarFile
   tRar.CmtBufSize = 0&

   hRar = RAROpenArchive(tRar)  ' Open the archive
   If tRar.OpenResult <> ERAR_SUCCESS Then Exit Function

   sDir = String$(MAX_PATH, vbNullChar)
   CharToOem sExtractDir, sDir

   rc = RARReadHeader(hRar, tHeader) ' Read first file header
   Do While rc = ERAR_SUCCESS

      ' Get the current file name
      sFile = Left$(tHeader.FileName, InStr(tHeader.FileName, vbNullChar) - 1&)

      If sFile = sRelFile Then ' Relative file spec as recorded in RAR archive
         rc = RARProcessFile(hRar, RAR_EXTRACT, sDir, vbNullString)
         UnpackFile = (rc = ERAR_SUCCESS) ' Return -1 on success, zero on error
         Exit Do
      Else
         rc = RARProcessFile(hRar, RAR_SKIP, vbNullString, vbNullString)
      End If
      If rc <> ERAR_SUCCESS Then Exit Do

      rc = RARReadHeader(hRar, tHeader) ' Read next file header
   Loop

ErrHandler:
   RARCloseArchive hRar
End Function

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Public Function UnpackToMemory(sRarFile As String, sRelFile As String, sRetFileText As String) As Long
   On Error GoTo ErrHandler ' Return file size on success, zero on error

   Dim tRar As RAROpenArchiveData
   Dim tHeader As RARHeaderData

   Dim sFile As String
   Dim hRar As Long
   Dim rc As Long

   tRar.OpenMode = RAR_OM_EXTRACT ' Fill the RAR header structure
   tRar.ArcName = sRarFile
   tRar.CmtBufSize = 0&
   m_Size = 0&

   hRar = RAROpenArchive(tRar)   ' Open the archive
   If tRar.OpenResult <> ERAR_SUCCESS Then Exit Function

   rc = RARReadHeader(hRar, tHeader) ' Read first file header

   Do While rc = ERAR_SUCCESS

      ' Get the current file name
      sFile = Left$(tHeader.FileName, InStr(tHeader.FileName, vbNullChar) - 1&)

      If sFile = sRelFile Then ' Relative file spec as recorded in RAR archive

         RARSetCallback hRar, AddressOf UnRARCallback, VarPtr(sRetFileText) ' Set the callback

         rc = RARProcessFile(hRar, RAR_EXTRACT, vbNullString, vbNullString)
         Exit Do
      Else
         rc = RARProcessFile(hRar, RAR_SKIP, vbNullString, vbNullString)
      End If
      If rc <> ERAR_SUCCESS Then Exit Do

      rc = RARReadHeader(hRar, tHeader) ' Read next file header
   Loop

   UnpackToMemory = m_Size ' Return file size on success, zero on error

ErrHandler:
   RARCloseArchive hRar
End Function

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Private Function UnRARCallback(ByVal Msg As Long, FileText As String, ByVal StringPtr As Long, ByVal Length As Long) As Long
   Dim aFileTxt() As Byte ' Return a positive value to continue process or -1 to cancel the archive operation

   If Msg = UCM_PROCESSDATA Then ' Process unpacked data

      ReDim aFileTxt(1 To Length) As Byte
      CopyMemByV VarPtr(aFileTxt(1)), StringPtr, Length
      FileText = StrConv(aFileTxt, vbUnicode)

      m_Size = Length
      UnRARCallback = -1& ' Cancel extraction to disk
   End If
End Function

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' Rd :)
