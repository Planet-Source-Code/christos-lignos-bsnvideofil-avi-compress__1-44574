VERSION 5.00
Object = "*\ABsnProgBar.vbp"
Begin VB.Form fMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Compress   AVI   File"
   ClientHeight    =   2925
   ClientLeft      =   3555
   ClientTop       =   2535
   ClientWidth     =   4185
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   4185
   Begin BsnProgressBar.BarColor BarColor1 
      Height          =   420
      Left            =   180
      TabIndex        =   3
      Top             =   1755
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   741
      Appearance      =   0
      BackColor       =   -2147483643
      Caption         =   "0%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancelSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   915
      Width           =   1710
   End
   Begin VB.CommandButton cmdOpenAVIFile 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select File"
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   390
      Width           =   1710
   End
   Begin VB.Label Status 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2595
      Width           =   4170
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    
    Call AVIFileInit
    MetrishS = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If cmdCancelSave.Enabled = True Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call AVIFileExit
End Sub

Private Sub cmdOpenAVIFile_Click()
    
    Dim res As Long
    Dim ofd As FileCls
    Dim szFile As String
    Dim pAVIFile As Long
    Dim pAVIStream As Long
    Dim szFileOut As String
    Dim pAVIStreamOut As Long
    Dim FilCompress As AVI_COMPRESS_OPTIONS
    Dim pFilCompress As Long
    
    Dim numFrames As Long



    MetrishS = 0
    MaximumS = 0
    
    Set ofd = New FileCls
    With ofd
        .OwnerHwnd = Me.hWnd
        .Filter = "AVI Files|*.avi"
        .DlgTitle = "Select AVI File to Copy Video From"
    End With
    res = ofd.VBGetOpenFileNamePreview(szFile)
    
    If res = False Then GoTo ErrorOut

    'Open the AVI File
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut

    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut

    ofd.DlgTitle = "Choose Location and Name to Save New AVI File"
    ofd.DefaultExt = "avi"
    szFileOut = "Out_File.avi"
    res = ofd.VBGetSaveFileName(szFileOut)
    If res = False Then
        MsgBox "User cancelled - no file saved.", vbInformation, App.title
        GoTo ErrorOut
    End If

    DoEvents

    pFilCompress = VarPtr(FilCompress)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE Or ICMF_CHOOSE_PREVIEW, _
                        1, _
                        pAVIStream, _
                        pFilCompress)
    If res <> 1 Then
        MsgBox "AVI SaveOptions returned an error !", vbCritical, App.title
        res = 0
        GoTo ErrorOut
    End If

    DoEvents
    
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut
    
MaximumS = numFrames
BarColor1.Max = MaximumS
    

    'recompress
    res = AVIMakeCompressedStream(pAVIStreamOut, pAVIStream, FilCompress, 0&)
    If res <> AVIERR_OK Then
        Call AVISaveOptionsFree(1, pFilCompress)
        GoTo ErrorOut
    End If


    gfAbort = False
    cmdOpenAVIFile.Enabled = False
    cmdCancelSave.Enabled = True
    pFilCompress = VarPtr(FilCompress)
    
    res = AVISave(szFileOut, 0&, AddressOf TestAVISave, 1, pAVIStreamOut, pFilCompress)
    DoEvents
    
    If res = AVIERR_USERABORT Then
        Status = "User cancelled .."
    Else
        Status = "Finished .."
    End If

    Call AVISaveOptionsFree(1, pFilCompress)


ErrorOut:
    cmdCancelSave.Enabled = False
    cmdOpenAVIFile.Enabled = True

    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream)
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile)
    End If

    If (res <> AVIERR_OK) Then
        If res <> AVIERR_USERABORT Then
            MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
        End If
    End If
    
End Sub

Private Sub cmdCancelSave_Click()
    gfAbort = True
End Sub

