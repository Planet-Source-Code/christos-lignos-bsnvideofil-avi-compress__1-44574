Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function AVISaveOptions Lib "avifil32.dll" (ByVal hWnd As Long, _
               ByVal uiFlags As Long, ByVal nStreams As Long, ByRef ppavi As Long, _
               ByRef ppOptions As Long) As Long

Public Declare Function AVISave Lib "avifil32.dll" Alias "AVISaveVA" (ByVal szFile As String, _
               ByVal pclsidHandler As Long, ByVal lpfnCallback As Long, ByVal nStreams As Long, _
               ByRef ppaviStream As Long, ByRef ppCompOptions As Long) As Long

Public Declare Function AVISaveOptionsFree Lib "avifil32.dll" (ByVal nStreams As Long, _
               ByRef ppOptions As Long) As Long

Public Declare Function AVIMakeCompressedStream Lib "avifil32.dll" (ByRef ppsCompressed As Long, _
               ByVal psSource As Long, ByRef lpOptions As AVI_COMPRESS_OPTIONS, _
               ByVal pclsidHandler As Long) As Long '

Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long

Public Declare Sub AVIFileExit Lib "avifil32.dll" ()
Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT

Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long

'  Bitmap
Public Type AVI_RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount  As Long
    dwFormatChangeCount As Long
    szName As String * 64
End Type

'  AVIFIle   Info
Public Type AVI_FILE_INFO
    dwMaxBytesPerSecond As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

Public Type AVI_COMPRESS_OPTIONS
    fccType As Long            '/* stream type, for consistency */
    fccHandler As Long         '/* compressor */
    dwKeyFrameEvery As Long    '/* keyframe rate */
    dwQuality As Long          '/* compress quality 0-10,000 */
    dwBytesPerSecond As Long   '/* bytes per second */
    dwFlags As Long            '/* flags... see below */
    lpFormat As Long           '/* save format */
    cbFormat As Long
    lpParms As Long            '/* compressor options */
    cbParms As Long
    dwInterleaveEvery As Long  '/* for non-video streams only */
End Type

Private Const SEVERITY_ERROR    As Long = &H80000000
Private Const FACILITY_ITF      As Long = &H40000
Private Const AVIERR_BASE       As Long = &H4000

Global Const AVIERR_OK As Long = 0&
Global Const AVIERR_USERABORT   As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 198) '-2147204922
Global Const OF_SHARE_DENY_WRITE As Long = &H20
Global Const streamtypeVIDEO       As Long = 1935960438
Global Const ICMF_CHOOSE_KEYFRAME           As Long = &H1
Global Const ICMF_CHOOSE_DATARATE           As Long = &H2
Global Const ICMF_CHOOSE_PREVIEW            As Long = &H4
Global gfAbort As Boolean

Global MaximumS As Long
Global MetrishS As Long

Public Function TestAVISave(ByVal nPercent As Long) As Long

  MetrishS = MetrishS + 1

fMain.BarColor1.Value = MetrishS
DoEvents

If gfAbort = True Then
    TestAVISave = AVIERR_USERABORT
Else
    TestAVISave = AVIERR_OK
End If

End Function


