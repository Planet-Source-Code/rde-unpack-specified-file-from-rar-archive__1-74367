Attribute VB_Name = "mUnrar"
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
   RAR_SKIP = 0    ' Move to the next file in the archive. If the archive is solid and RAR_OM_EXTRACT mode was set when the archive was opened, the current file will be processed - the operation will be performed slower than a simple seek
   RAR_TEST = 1    ' Test the current file and move to the next file in the archive. If the archive was opened with RAR_OM_LIST mode, the operation is equal to RAR_SKIP
   RAR_EXTRACT = 2 ' Extract the current file and move to the next file. If the archive was opened with RAR_OM_LIST mode, the operation is equal to RAR_SKIP
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

Public Enum RarOperations
   OP_EXTRACT = 0 '
   OP_TEST = 1    '
   OP_LIST = 2    '
End Enum

Private Enum eFileFlags
   eContPrev = &H1& ' File continued from previous volume
   eContNext = &H2& ' File continued on next volume
   eEncrypt = &H4&  ' File encrypted with password
   eComment = &H8&  ' File comment present
   eSolid = &H10&   ' Compression of previous files is used (solid flag)
   '  bits 7 6 5
   '       0 0 0    ' Dictionary size   64 Kb
   '       0 0 1    ' Dictionary size  128 Kb
   '       0 1 0    ' Dictionary size  256 Kb
   '       0 1 1    ' Dictionary size  512 Kb
   '       1 0 0    ' Dictionary size 1024 Kb
   '       1 0 1    ' Dictionary size 2048 KB
   '       1 1 0    ' Dictionary size 4096 KB
   '       1 1 1    ' File is directory
   '  Other bits are reserved
End Enum

Private Enum eFlagsEx
   eVolumeEx = &H1&     ' Volume attribute (archive volume)
   eCommentEx = &H2&    ' Archive comment present
   eLockEx = &H4&       ' Archive lock attribute
   eSolidEx = &H8&       ' Solid attribute (solid archive)
   eNewNameEx = &H10&   ' New volume naming scheme ('volname.partN.rar')
   eAuthentEx = &H20&   ' Authenticity information present
   eRecoveryEx = &H40&  ' Recovery record present
   eEncryptEx = &H80&   ' Block headers are encrypted
   eFirstVolEx = &H100& ' First volume (set only by RAR 3.0 and later)
End Enum

Private Enum eHost
  eDOS = 0   ' MS DOS
  eOS2 = 1   ' OS/2
  eWin32 = 2 ' Win32
  eUnix = 3  ' Unix
End Enum

Private Type RARHeaderData
   ArcName As String * 260  ' Output Null terminated current archive name. May be used to determine the current volume name
   FileName As String * 260 ' Output Null terminated file name in OEM (DOS) encoding
   Flags As Long            ' Output parameter which contains eFileFlags file flags
   PackSize As Long         ' Output packed file size or size of the file part if file was split between volumes
   UnpSize As Long          ' Output parameter - unpacked file size
   HostOS As Long           ' Output eHost parameter - operating system used for archiving
   FileCRC As Long          ' Output unpacked file CRC. It should not be used for file parts which were split between volumes
   FileTime As Long         ' Output parameter contains date and time in standard MS DOS format
   UnpVer As Long           ' Output RAR version needed to extract file. It is encoded as 10 * Major version + minor version
   Method As Long           ' Output parameter - packing method
   FileAttr As Long         ' Output parameter - file attributes
   CmtBuf As String '*      ' Input buffer for file comments. Maximum comment size is limited to 64Kb. Comment text is a Null terminated string in OEM encoding. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long       ' Input size of buffer for archive comments
   CmtSize As Long          ' Output size of comments actually read into the buffer, will not exceed CmtBufSize
   CmtState As Long         ' Output comment state (File comments support is not implemented yet. CmtState is always 0).
End Type

Private Type RARHeaderDataEx
   ArcName As String * 1024   ' Output Null terminated current archive name. May be used to determine the current volume name
   ArcNameW As String * 1024  ' Output Null terminated current archive name. May be used to determine the current volume name
   FileName As String * 1024  ' Output Null terminated file name in OEM (DOS) encoding
   FileNameW As String * 1024 ' Output Null terminated Unicode file name
   Flags As Long              ' Output parameter which contains eFileFlags file flags
   PackSize As Long           ' Output parameter means packed file size or size of the file part if file was split between volumes
   PackSizeHigh As Long       ' Output parameter means packed file size or size of the file part if file was split between volumes
   UnpSize As Long            ' Output parameter - unpacked file size
   UnpSizeHigh As Long        ' Output parameter - unpacked file size
   HostOS As Long             ' Output eHost parameter - operating system used for archiving
   FileCRC As Long            ' Output unpacked file CRC. It should not be used for file parts which were split between volumes
   FileTime As Long           ' Output parameter contains date and time in standard MS DOS format
   UnpVer As Long             ' Output RAR version needed to extract file. It is encoded as 10 * Major version + minor version
   Method As Long             ' Output parameter - packing method
   FileAttr As Long           ' Output parameter - file attributes
   CmtBuf As String '*        ' Input parameter which should point to the buffer for file comments. Maximum comment size is limited to 64Kb. Comment text is a Null terminated string in OEM encoding. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long         ' Input parameter which should contain size of buffer for archive comments
   CmtSize As Long            ' Output containing size of comments actually read into the buffer, should not exceed CmtBufSize
   CmtState As Long           ' Output ERAR_NO_COMMENTS, ERAR_COMMENTS_READ, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_UNKNOWN_FORMAT, ERAR_SMALL_BUF
   Reserved(1024) As Long     ' Reserved for future use. Must be zero
End Type

Private Type RAROpenArchiveData
   ArcName As String   ' Input Null terminated string containing the archive name
   OpenMode As Long    ' Input eOpenMode parameter - RAR_OM_LIST, RAR_OM_EXTRACT
   OpenResult As Long  ' Output ERAR_SUCCESS, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_BAD_ARCHIVE, ERAR_UNKNOWN_FORMAT, or ERAR_EOPEN
   CmtBuf As String    ' Input buffer for archive comments. Maximum comment size is limited to 64Kb. Comment text is Null terminated. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long  ' Input size of buffer for archive comments
   CmtSize As Long     ' Output size of comments actually read into the buffer, cannot exceed CmtBufSize
   CmtState As Long    ' Output ERAR_NO_COMMENTS, ERAR_COMMENTS_READ, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_UNKNOWN_FORMAT, ERAR_SMALL_BUF
End Type

Private Type RAROpenArchiveDataEx
   ArcName As String    ' Input Null terminated string containing the archive name
   ArcNameW As String   ' Input Null terminated Unicode archive name or NULL if Unicode name is not specified
   OpenMode As Long     ' Input eOpenMode parameter - RAR_OM_LIST, RAR_OM_EXTRACT
   OpenResult As Long   ' Output ERAR_SUCCESS, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_BAD_ARCHIVE, ERAR_UNKNOWN_FORMAT, or ERAR_EOPEN
   CmtBuf As String     ' Input buffer for archive comments. Maximum comment size is limited to 64Kb. Comment text is zero terminated. If the comment text is larger than the buffer size, the comment text will be truncated. If CmtBuf is set to NULL, comments will not be read
   CmtBufSize As Long   ' Input size of buffer for archive comments
   CmtSize As Long      ' Output size of comments actually read into the buffer, cannot exceed CmtBufSize
   CmtState As Long     ' Output ERAR_NO_COMMENTS, ERAR_COMMENTS_READ, ERAR_NO_MEMORY, ERAR_BAD_DATA, ERAR_UNKNOWN_FORMAT, ERAR_SMALL_BUF
   Flags As Long        ' Output parameter. Combination of eFlagsEx bit flags
   Reserved(32) As Long ' Reserved for future use. Must be zero
End Type

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' Open RAR archive and allocate memory structures
' ArchiveData points to RAROpenArchiveData structure
' Returns archive handle or NULL in case of error
Private Declare Function RAROpenArchive Lib "unrar3" (ArchiveData As RAROpenArchiveData) As Long 'Handle

' Similar to RAROpenArchive, but uses RAROpenArchiveDataEx structure allowing to
' specify Unicode archive name and returning information about archive flags
' ArchiveDataEx points to RAROpenArchiveDataEx structure
' Returns archive handle or NULL in case of error
Private Declare Function RAROpenArchiveEx Lib "unrar3" (ArchiveDataEx As RAROpenArchiveDataEx) As Long 'Handle

' Close RAR archive and release allocated memory. It must be called when archive
' processing is finished, even if the archive processing was stopped due to an error
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' Returns ERAR_SUCCESS on success or ERAR_ECLOSE archive close error
Private Declare Function RARCloseArchive Lib "unrar3" (ByVal hArcData As Long) As Long

' Read header of file in archive
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' HeaderData points to RARHeaderData structure
' Returns ERAR_SUCCESS on success, ERAR_END_ARCHIVE end of archive, or ERAR_BAD_DATA file header broken
Private Declare Function RARReadHeader Lib "unrar3" (ByVal hArcData As Long, HeaderData As RARHeaderData) As Long

' Similar to RARReadHeader, but uses RARHeaderDataEx structure,
' containing information about Unicode file names and 64 bit file sizes.
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' HeaderDataEx points to RARHeaderDataEx structure
' Returns ERAR_SUCCESS on success, ERAR_END_ARCHIVE end of archive, or ERAR_BAD_DATA file header broken
Private Declare Function RARReadHeaderEx Lib "unrar3" (ByVal hArcData As Long, HeaderDataEx As RARHeaderDataEx) As Long

' Performs action and moves the current position in the archive to the next file.
' Extract or test the current file from the archive opened in RAR_OM_EXTRACT mode.
' If the mode RAR_OM_LIST is set, then a call to this function will simply skip
' the archive position to the next file.
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' Operation - File eProcessFileOp operation, RAR_SKIP, RAR_TEST, RAR_EXTRACT
' DestPath - Destination extract directory. If DestPath is Null extracts to the current directory. This parameter has meaning only if DestName is NULL.
' DestName - Full path and name of the file to be extracted or Null. If DestName is defined (not Null) it overrides the original file name saved in the archive and DestPath setting.
' Both DestPath and DestName must be in OEM encoded. If necessary, use CharToOem to convert text to OEM before passing to this function.
' Returns ERAR_SUCCESS, ERAR_BAD_DATA, ERAR_BAD_ARCHIVE, ERAR_UNKNOWN_FORMAT, ERAR_EOPEN, ERAR_ECREATE, ERAR_ECLOSE, ERAR_EREAD, or ERAR_EWRITE
' Note: if you wish to cancel extraction, return -1 when processing UCM_PROCESSDATA callback message.
Private Declare Function RARProcessFile Lib "unrar3" (ByVal hArcData As Long, ByVal Operation As Long, ByVal pDestPath As String, ByVal pDestName As String) As Long

' Converts strings from Unicode to OEM encoding to make sure
' certain characters in paths are handled properly by RARProcessFile
' Both DestPath and DestName must be in OEM encoded.
' Use CharToOem to convert text to OEM before passing to RARProcessFile.
Private Declare Sub CharToOem Lib "user32" Alias "CharToOemA" (ByVal sSrc As String, ByVal sDest As String)

' Unicode version of RARProcessFile. It uses Unicode DestPath and DestName parameters,
' other parameters and return values are the same as in RARProcessFile.
Private Declare Function RARProcessFileW Lib "unrar3" (ByVal hArcData As Long, ByVal Operation As Long, ByVal StrPtr_DestPath As Long, ByVal StrPtr_DestName As Long) As Long

' Set a user-defined callback function to process Unrar events.
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' CallbackProc points to a user-defined callback function (AddressOf_Callbackfunc).
' UserData - User data passed to callback function (ObjPtr_Me).
' Other functions of UNRAR.DLL should not be called from the callback function.
' This subroutine has no return value.
Private Declare Sub RARSetCallback Lib "unrar3" (ByVal hArcData As Long, ByVal CallbackProc As Long, ByVal UserData As Long)

'Obsoleted, use RARSetCallback instead.
'void   PASCAL RARSetChangeVolProc(HANDLE hArcData,CHANGEVOLPROC ChangeVolProc);
'void   PASCAL RARSetProcessDataProc(HANDLE hArcData,PROCESSDATAPROC ProcessDataProc);
'typedef int (PASCAL *CHANGEVOLPROC)(char *ArcName,int Mode);
'typedef int (PASCAL *PROCESSDATAPROC)(unsigned char *Addr,int Size);

' Set a password to decrypt files.
' hArcData contains the archive handle obtained from the RAROpenArchive function call
' Password string containing a Null terminated password.
' This subroutine has no return value.
Private Declare Sub RARSetPassword Lib "unrar3" (ByVal hArcData As Long, ByVal Password As String)

' Returns unrar.dll version.
' Returns a value denoting DLL version. The current version value is defined in unrar.h as RAR_DLL_VERSION
' This function is absent in old versions of unrar.dll, so it may be wise
' to use LoadLibrary and GetProcAddress to access this function.
Private Declare Function RARGetDllVersion Lib "unrar3" () As Long

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Private Const RAR_ABORT As Long = -1
Private Const RAR_CONTINUE As Long = 1
'
'UNRARCALLBACK_MESSAGES Callback messages
'    UCM_CHANGEVOLUME
'        RAR_VOL_ASK
'        RAR_VOL_NOTIFY
'    UCM_PROCESSDATA
'    UCM_NEEDPASSWORD

Private Function UnRARCallback(ByVal Msg As Long, ByVal UserData As Long, ByVal Param1 As Long, ByVal Param2 As Long) As Long

' Note: if you wish to cancel extraction, return -1 when processing UCM_PROCESSDATA callback message.

   ' The function will be passed four parameters:
   '  Msg                    Type of event. Described below.
   '  UserData               User defined value passed to RARSetCallback.
   '  P1 and P2              Event dependent parameters. Described below.
   '
   ' Possible events
   '
   '    UCM_CHANGEVOLUME     Process volume change.
   '
   '      P1                   Points to the zero terminated name of the next volume.
   '
   '      P2                   The function call mode:
   '
   '        RAR_VOL_ASK          Required volume is absent. The function should
   '                             prompt user and return a positive value
   '                             to retry or return -1 value to terminate
   '                             operation. The function may also specify a new
   '                             volume name, placing it to the address specified
   '                             by P1 parameter.
   '
   '        RAR_VOL_NOTIFY       Required volume is successfully opened.
   '                             This is a notification call and volume name
   '                             modification is not allowed. The function should
   '                             return a positive value to continue or -1
   '                             to terminate operation.
   '
   '    UCM_PROCESSDATA          Process unpacked data. It may be used to read
   '                             a file while it is being extracted or tested
   '                             without actual extracting file to disk.
   '                             Return a positive value to continue process
   '                             or -1 to cancel the archive operation
   '
   '      P1                   Address pointing to the unpacked data.
   '                           Function may refer to the data but must not
   '                           change it.
   '
   '      P2                   Size of the unpacked data. It is guaranteed
   '                           only that the size will not exceed the maximum
   '                           dictionary size (4 Mb in RAR 3.0).
   '
   '    UCM_NEEDPASSWORD         DLL needs a password to process archive.
   '                             This message must be processed if you wish
   '                             to be able to handle archives with encrypted
   '                             file names. It can be also used as replacement
   '                             of RARSetPassword function even for usual
   '                             encrypted files with non-encrypted names.
   '
   '      P1                   Address pointing to the buffer for a password.
   '                           You need to copy a password here.
   '
   '      P2                   Size of the password buffer.

End Function

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

'RARExecute OP_EXTRACT, sRarFile

Public Sub RARExecute(ByVal Mode As RarOperations, RarFile As String, Optional Password As String)
    ' Description:-
    ' Extract file(s) from RAR archive.
    ' Parameters:-
    ' Mode = Operation to perform on RAR Archive
    ' RARFile = RAR Archive filename
    ' Password = Password (Optional)
    Dim lHandle As Long
    Dim iStatus As Integer
    Dim uRAR As RAROpenArchiveData
    Dim uHeader As RARHeaderData
    Dim sStat As String, Ret As Long

    uRAR.ArcName = RarFile
    uRAR.CmtBuf = Space$(16384)
    uRAR.CmtBufSize = 16384

    If Mode = OP_LIST Then
        uRAR.OpenMode = RAR_OM_LIST
    Else
        uRAR.OpenMode = RAR_OM_EXTRACT
    End If

    lHandle = RAROpenArchive(uRAR)
    If uRAR.OpenResult <> 0 Then
        'Kill RarFile
        OpenError uRAR.OpenResult, RarFile
    End If

    If Password <> "" Then RARSetPassword lHandle, Password

    If (uRAR.CmtState = 1) Then MsgBox uRAR.CmtBuf, vbApplicationModal + vbInformation, "Comment"

    iStatus = RARReadHeader(lHandle, uHeader)

    Do Until iStatus <> 0
       'sStat = Left$(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1)
        Select Case Mode
          Case RarOperations.OP_EXTRACT
            Ret = RARProcessFile(lHandle, RAR_EXTRACT, "", uHeader.FileName)
          Case RarOperations.OP_TEST
            Ret = RARProcessFile(lHandle, RAR_TEST, "", uHeader.FileName)
          Case RarOperations.OP_LIST
            Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
        End Select

        If Ret <> 0 Then ProcessError Ret

        iStatus = RARReadHeader(lHandle, uHeader)
    Loop

    If iStatus = ERAR_BAD_DATA Then MsgBox "File header broken", vbCritical

    RARCloseArchive lHandle
End Sub

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

' Error handling
Private Sub OpenError(ErroNum As Long, ArcName As String)
Dim erro As String

    Select Case ErroNum
        Case ERAR_NO_MEMORY
            erro = "Not enough memory"
            GoTo errorbox
        Case ERAR_EOPEN:
            erro = "Cannot open " & ArcName
            GoTo errorbox
        Case ERAR_BAD_ARCHIVE:
            erro = ArcName & " is not RAR archive"
            GoTo errorbox
        Case ERAR_BAD_DATA:
            erro = ArcName & ": archive header broken"
            GoTo errorbox
    End Select

    Exit Sub

errorbox:
    MsgBox erro, vbCritical
End Sub

Private Sub ProcessError(ErroNum As Long)
Dim erro As String

    Select Case ErroNum
        Case ERAR_UNKNOWN_FORMAT
            erro = "Unknown archive format"
            GoTo errorbox
        Case ERAR_BAD_ARCHIVE:
            erro = "Bad volume"
            GoTo errorbox
        Case ERAR_ECREATE:
            erro = "File create error"
            GoTo errorbox
        Case ERAR_EOPEN:
            erro = "Volume open error"
            GoTo errorbox
        Case ERAR_ECLOSE:
            erro = "File close error"
            GoTo errorbox
        Case ERAR_EREAD:
            erro = "Read error"
            GoTo errorbox
        Case ERAR_EWRITE:
            erro = "Write error"
            GoTo errorbox
        Case ERAR_BAD_DATA:
            erro = "CRC error"
            GoTo errorbox
    End Select

    Exit Sub

errorbox:
    MsgBox erro, vbCritical
End Sub

' =*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

