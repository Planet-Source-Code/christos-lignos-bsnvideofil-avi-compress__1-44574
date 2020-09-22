VERSION 5.00
Begin VB.UserControl BarColor 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   510
   ScaleWidth      =   3870
   ToolboxBitmap   =   "ColorBar.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   105
         Width           =   270
      End
   End
End
Attribute VB_Name = "BarColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const m_def_AutoCaption = True
Const m_def_StrExtend = "%"
Const m_def_Interactive = True
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
Dim m_AutoCaption As Boolean
Dim m_StrExtend As String
Dim m_Interactive As Boolean
Dim m_Min As Variant
Dim m_Max As Variant
Dim m_Value As Variant
Event Change()
Attribute Change.VB_Description = "Se produit lorsque le contenu d'un contrôle a été modifié."
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)




'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Renvoie ou définit si un objet apparaît ou non en 3D au moment de l'exécution."
    Appearance = Picture1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    Picture1.Appearance() = New_Appearance
    UserControl_Resize
    PropertyChanged "Appearance"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    Value = m_Value
    PropertyChanged "BackColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Renvoie ou définit le style de la bordure d'un objet."
    BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Picture1.BorderStyle() = New_BorderStyle
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Renvoie ou définit le texte affiché dans la barre de titre d'un objet ou sous l'icône d'un objet."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption & IIf(m_AutoCaption, m_StrExtend, "")
    Label1.Move (Picture1.ScaleWidth - Label1.Width) \ 2, (Picture1.ScaleHeight - Label1.Height) \ 2
    PropertyChanged "Caption"
End Property

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X + Label1.Left, Y + Label1.Top)
    Call Picture1_MouseDown(Button, Shift, X + Label1.Left, Y + Label1.Top)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X + Label1.Left, Y + Label1.Top)
    Call Picture1_MouseMove(Button, Shift, X + Label1.Left, Y + Label1.Top)
End Sub

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Renvoie un objet Font."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    Value = m_Value
    PropertyChanged "ForeColor"
End Property

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If m_Interactive Then
            If Button = 1 Then
                Value = XinValue(X)
            End If
     End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If m_Interactive Then
            If Button = 1 Then
                Value = XinValue(X)
            End If
     End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=14,0,0,0
Public Property Get Min() As Variant
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Variant)
    m_Min = New_Min
    Value = m_Value
    PropertyChanged "Min"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=14,0,0,100
Public Property Get Max() As Variant
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Variant)
    m_Max = New_Max
    Value = m_Value
    PropertyChanged "Max"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=14,0,0,50
Public Property Get Value() As Variant
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    m_Value = New_Value
    Picture1.Line (m_Value / m_Max * Picture1.ScaleWidth, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), Picture1.BackColor, BF
    Picture1.Line (0, 0)-(m_Value / m_Max * Picture1.ScaleWidth, Picture1.ScaleHeight), Picture1.FillColor, BF
    If m_AutoCaption Then
        Label1.Caption = CStr(m_Value) & m_StrExtend
        Label1.Move (Picture1.ScaleWidth - Label1.Width) \ 2, (Picture1.ScaleHeight - Label1.Height) \ 2
    End If
    RaiseEvent Change
    PropertyChanged "Value"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=14
Public Function XinValue(ByVal X As Long) As Long
Dim tempo As Long
    tempo = X / Picture1.ScaleWidth * m_Max
    If tempo > m_Max Then tempo = m_Max
    If tempo < m_Min Then tempo = m_Min
    XinValue = tempo
End Function

'Initialiser les propriétés pour le contrôle utilisateur
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_Interactive = m_def_Interactive
    m_StrExtend = m_def_StrExtend
    m_AutoCaption = m_def_AutoCaption
End Sub



'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Picture1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Label1.Caption = PropBag.ReadProperty("Caption", "1")
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFF00&)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    Picture1.FillColor = PropBag.ReadProperty("FillColor", &HC00000)
    
    m_Interactive = PropBag.ReadProperty("Interactive", m_def_Interactive)
    m_StrExtend = PropBag.ReadProperty("StrExtend", m_def_StrExtend)
    m_AutoCaption = PropBag.ReadProperty("AutoCaption", m_def_AutoCaption)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Picture1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
    Picture1.Move 0, 0, UserControl.Width, UserControl.Height
    Label1.Move (Picture1.ScaleWidth - Label1.Width) \ 2, (Picture1.ScaleHeight - Label1.Height) \ 2
    Value = m_Value
End Sub

'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", Picture1.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "50")
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFF00&)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("FillColor", Picture1.FillColor, &HC00000)
    Call PropBag.WriteProperty("Interactive", m_Interactive, m_def_Interactive)
    Call PropBag.WriteProperty("StrExtend", m_StrExtend, m_def_StrExtend)
    Call PropBag.WriteProperty("AutoCaption", m_AutoCaption, m_def_AutoCaption)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", Picture1.MousePointer, 0)
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Picture1,Picture1,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Renvoie ou définit la couleur de remplissage des formes, des cercles et des boîtes."
    FillColor = Picture1.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    Picture1.FillColor() = New_FillColor
    Value = m_Value
    PropertyChanged "FillColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,True
Public Property Get Interactive() As Boolean
    Interactive = m_Interactive
End Property

Public Property Let Interactive(ByVal New_Interactive As Boolean)
    m_Interactive = New_Interactive
    PropertyChanged "Interactive"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=13,0,0,%
Public Property Get StrExtend() As String
    StrExtend = m_StrExtend
End Property

Public Property Let StrExtend(ByVal New_StrExtend As String)
    m_StrExtend = New_StrExtend
    Value = m_Value
    PropertyChanged "StrExtend"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,True
Public Property Get AutoCaption() As Boolean
    AutoCaption = m_AutoCaption
End Property

Public Property Let AutoCaption(ByVal New_AutoCaption As Boolean)
    m_AutoCaption = New_AutoCaption
    PropertyChanged "AutoCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Picture1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Picture1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Picture1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Picture1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

