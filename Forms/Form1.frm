VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "What' s In the Variant"
   ClientHeight    =   6015
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCreate 
      Caption         =   "Create"
      Height          =   735
      Left            =   2400
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   5175
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   15
      Top             =   840
      Width           =   9015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   2175
      TabIndex        =   11
      Top             =   840
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   0
         TabIndex        =   18
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox Combo6 
         Height          =   345
         Left            =   0
         TabIndex        =   20
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Datatype:"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox PnlTextBox 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12855
      TabIndex        =   8
      Top             =   0
      Width           =   12855
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton BtnDebugTest8 
      Caption         =   "Debug Test8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest7 
      Caption         =   "Debug Test7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest6 
      Caption         =   "Debug Test6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest5 
      Caption         =   "Debug Test5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest4 
      Caption         =   "Debug Test4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest3 
      Caption         =   "Debug Test3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest2 
      Caption         =   "Debug Test2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton BtnDebugTest1 
      Caption         =   "Debug Test1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand TextBox"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'What's in the Variant?
'======================
'the Variant
Dim VarVal

'Hilfsvariablen
Dim BytVal As Byte
Dim IntVal As Integer
Dim LngVal As Long
Dim SngVal As Single
Dim DblVal As Double
Dim CurVal As Currency
Dim DecVal As Variant
Dim BolVal As Boolean
Dim DatVal As Date
Dim DObVal As IUnknown 'DataObject
Dim StrVal As String
Dim ErrVal As Long 'ErrObject
Dim UdtVal As MSFlexGridWizard.GridSettingsType 'just as an example udt
'Dim CArVal As TCArr
Dim ArrVal 'an arbitrary Array as Array as Array . . .
Dim ClsVal As Dummy       ' serves as a typical user Class
Dim ColVal As New Collection ' serves as a typical VBA Class

Private Sub Form_Load()
    InitData
    
    VarTypes_ToCombo Combo1
    VarTypes_ToCombo Combo2, vbEmpty, vbNull
    VarTypes_ToCombo Combo3, vbEmpty, vbNull
    VarTypes_ToCombo Combo4, vbEmpty, vbNull
    VarTypes_ToCombo Combo5, vbEmpty, vbNull
    VarTypes_ToCombo Combo6, vbEmpty, vbNull, vbArray
    
    Combo2.Visible = False
    Combo3.Visible = False
    Combo4.Visible = False
    Combo5.Visible = False
    Combo6.Visible = False
    Combo1.ListIndex = 14
    
End Sub


Private Sub Command1_Click()
    Dim v1, v2
    v1 = CByte(123) 'CSByte("123")
    v2 = CByte(23) 'CSByte("-23")
    
    MsgBox EVarType_ToStr(VarType(v1))
    
    Dim v: v = VAdd(v1, v2)
    
    MsgBox EVarType_ToStr(VarType(v))
    
    'leider kann man nicht damit Rechnen, zum Rechnen muss man erst wieder in einen von VB unterstützten Typ umwandeln
    'geht nicht
    'sb = sb + 1
    
    'geht
    Dim sb As Byte
    sb = CByte(sb)
    sb = sb + CByte(1)
    
    MsgBox EVarType_ToStr(VarType(sb))
    
End Sub

'Private Sub Command2_Click()
''    Dim pfn
'    'pfn = App.Path & "\test.png"
''    pfn = "C:\Testdir\testpics\test.bmp"
''    Debug.Print pfn
'
''    'Dim vpic As Long 'As IPictureDisp
''    Dim vpic As IPictureDisp
''    OleLoadPictureFile pfn, vpic
'    'Debug.Print vpic
'    'Dim pic As IPictureDisp: 'Set pic = vpic
'    'If pic Is Nothing Then Debug.Print "pic is nothing"
'    'Set Picture2.Picture = vpic
''    Picture2.Picture = vpic
'    'Picture1.Picture = LoadPicture(pfn)
'    'Set Picture2.Picture = LoadPic("C:\Testdir\testpics\test.gif")
'    Set Picture2.Picture = LoadPic(App.Path & "\Resources\test3.png")
'End Sub

Private Sub InitData()
    BolVal = True
    StrVal = "Text"
    Set ClsVal = New Dummy: ClsVal.Value = "Testtest"
    BytVal = 123
    IntVal = 12345
    LngVal = 123456789
    SngVal = 3.141593
    DblVal = 3.14159265358979
    CurVal = CCur("123456789012,3456")
    DatVal = Now
    Set DObVal = ColVal.[_NewEnum]
    DecVal = CDec("12345678901234567890,1234567890")
    'Set ErrVal = Err: ErrVal.Number = 9
    ErrVal = 9
    With UdtVal
        .AllowColDragging = True
        .GridStyle = MSFlexGridWizard.GridTypeSettings.gtOutline
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PopupMenu mnuOptions
    End If
End Sub
Private Sub PnlTextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = 8 * Screen.TwipsPerPixelX
    Dim t As Single: t = l
    Dim W As Single: W = Me.ScaleWidth '- 2 * L
    Dim H As Single: H = PnlTextBox.Height
    If W > 0 And H > 0 Then
        PnlTextBox.Move 0, 0, W, H
        Text1.Move l, Text1.Top, IIf(mnuExpand.Checked, W - 2 * l, 3615), Text1.Height
        Text2.Move Text2.Left, Text2.Top, W - Text2.Left, Me.ScaleHeight - Text2.Top
    End If
End Sub

Private Sub mnuExpand_Click()
    mnuExpand.Checked = Not mnuExpand.Checked
    Form_Resize
End Sub

Private Sub Combo1_Click()
    Combo2.Visible = False
    Combo3.Visible = False
    Combo4.Visible = False
    Combo5.Visible = False
    Combo6.Visible = False
    cmbinit Combo1, Combo2
End Sub
Private Sub Combo2_Click()
    Combo3.Visible = False
    Combo4.Visible = False
    Combo5.Visible = False
    Combo6.Visible = False
    cmbinit Combo2, Combo3
End Sub
Private Sub Combo3_Click()
    Combo4.Visible = False
    Combo5.Visible = False
    Combo6.Visible = False
    cmbinit Combo3, Combo4
End Sub
Private Sub Combo4_Click()
    Combo5.Visible = False
    Combo6.Visible = False
    cmbinit Combo4, Combo5
End Sub
Private Sub Combo5_Click()
    Combo6.Visible = False
    cmbinit Combo5, Combo6
End Sub
Private Sub Combo6_Click()
    cmbinit Combo6, Nothing
End Sub

Sub cmbinit(cmb As ComboBox, cmbnext As ComboBox)
    Dim vt As EVbVarType: vt = cmb.ItemData(cmb.ListIndex)
    InitTextBox vt
    If Not cmbnext Is Nothing Then
        cmbnext.Visible = vt = vbArray
        If cmbnext.Visible Then cmbnext.ListIndex = 0
    End If
End Sub

Private Sub InitTextBox(ByVal vt As EVbVarType)
    
    Dim a As AlignmentConstants

    Select Case vt
    Case vbEmpty, vbString, vbLPStr To vbWChar, vbDataObject, vbUserDefinedType
                                                                a = vbLeftJustify
    Case vbNull, vbBoolean, vbError, vbVoid, vbVariant:         a = vbCenter
    Case vbInteger To vbDate, vbDecimal To vbPtr, vbIntPtr, vbUIntPtr
                                                                a = vbRightJustify
    Case Else:                                                  a = vbLeftJustify
    End Select
    
    Dim s As String
    Select Case vt
    Case vbEmpty:                        s = ""
    Case vbNull, vbVoid:                 s = "0"
    Case vbBoolean:                      s = BolVal
    Case vbError:                        s = ErrVal '.Number
    Case vbString, vbLPStr, vbLPWStr
                                         s = StrVal
    Case vbWChar:                        s = "A"
    Case vbObject:                       s = ClsVal.Value
    Case vbDataObject:                   s = "[_NewEnum] As IUnknown" 'DObVal.GetData(0))
    Case vbInteger, vbUInteger:          s = IntVal
    Case vbLong, vbULong, vbInt, vbUInt: s = LngVal
    Case vbPtr, vbIntPtr, vbUIntPtr, vbHResult
                                         s = LngVal
    Case vbSingle:                       s = SngVal
    Case vbDouble:                       s = DblVal
    Case vbLongLong, vbULongLong:        s = Int(CurVal)
    Case vbCurrency:                     s = CurVal
    Case vbDate:                         s = DatVal
    Case vbDecimal:                      s = DecVal
    Case vbByte, vbSByte:                s = BytVal
    'Case vbUserDefinedType: s = GridSettingsType_ToStr(UdtVal)
    Case vbUserDefinedType:              s = Udt_ToStr(UdtVal)
    Case Else:              s = ""
    End Select
    Text1.Alignment = a
    Text1.Text = s
    
End Sub

Private Sub BtnCreate_Click()
Try: On Error GoTo Catch
    Dim v
    VarVal = Empty
    Dim s As String: s = Text1.Text
    Dim vt As EVbVarType: vt = Combo1.ItemData(Combo1.ListIndex)
    Dim vtx As EVbVarType
    If vt = vbArray Then
        vtx = Combo2.ItemData(Combo2.ListIndex)
        CreateArray VarVal, vtx, s
        If Combo3.Visible Then
            vtx = Combo3.ItemData(Combo3.ListIndex)
            CreateArray VarVal(0), vtx, s
            If Combo4.Visible Then
                vtx = Combo4.ItemData(Combo4.ListIndex)
                CreateArray VarVal(0)(0), vtx, s
                If Combo5.Visible Then
                    vtx = Combo5.ItemData(Combo5.ListIndex)
                    CreateArray VarVal(0)(0)(0), vtx, s
                    If Combo6.Visible Then
                        vtx = Combo6.ItemData(Combo6.ListIndex)
                        CreateArray VarVal(0)(0)(0)(0), vtx, s
                        AssignVar v, VarVal(0)(0)(0)(0)
                    Else
                        AssignVar v, VarVal(0)(0)(0)
                    End If
                Else
                    AssignVar v, VarVal(0)(0)
                End If
            Else
                AssignVar v, VarVal(0)
            End If
        Else
            AssignVar v, VarVal
        End If
    Else
        GetValue vt, v, s
        AssignVar VarVal, v
    End If
    s = VarType2_ToStr(VarVal) & " = " & VarValue_ToStr(VarVal)
    TextAdd s
    Exit Sub
Catch:
    ErrHandler Me, "BtnCreate", EVarType_ToStr(vt) & " = " & Text1.Text
End Sub

Sub GetValue(ByVal vt As EVbVarType, ByRef v_out, s As String)
    Select Case vt
    Case vbEmpty:       v_out = Empty
    Case vbNull, vbVoid
                        v_out = Null
    Case vbInteger:     IntVal = CInt(s):  v_out = IntVal
    Case vbUInteger:    v_out = CUInt(s)
    
    Case vbLong, vbPtr, vbHResult, vbInt, vbUInt, vbIntPtr, vbUIntPtr
                        LngVal = CLng(s):  v_out = LngVal
    Case vbULong:       v_out = CULng(s)
    Case vbSingle:
                    If Single_TryParse(s, SngVal) Then
                        v_out = SngVal
                    Else
                        MsgBox "Bitte eine Zahl eingeben: " & s
                        v_out = CSng(0)
                    End If
    Case vbDouble:
                    If Double_TryParse(s, DblVal) Then
                        v_out = DblVal
                    Else
                        MsgBox "Bitte eine Zahl eingeben: " & s
                        v_out = CDbl(0)
                    End If
    Case vbCurrency:    CurVal = CCur(s):  v_out = CurVal
    Case vbLongLong:    v_out = CLngLng(s)
    Case vbULongLong:   v_out = CULngLng(s)
    Case vbDate:        DatVal = CDate(s): v_out = DatVal
    Case vbString, vbLPStr, vbLPWStr
                        StrVal = CStr(s):  v_out = StrVal
    Case vbWChar:       IntVal = AscW(s):  v_out = IntVal
    Case vbObject:      ClsVal.Value = s: Set v_out = ClsVal
    Case vbError:       ErrVal = CLng(s): v_out = CVErr(ErrVal)
                    
    Case vbBoolean:     BolVal = CBool(s): v_out = BolVal
    Case vbVariant:
                        If TryRecognizeVar(s, v_out) Then:
                        
    Case vbDataObject:  v_out = DObVal 'no Set needed!
    Case vbDecimal:
                        If Decimal_TryParse(s, DecVal) Then
                            v_out = DecVal
                        Else
                            MsgBox "Bitte eine Zahl eingeben: " & s
                            v_out = CDec(0)
                        End If
    Case vbByte:        BytVal = CByte(s): v_out = BytVal
    Case vbSByte:       v_out = CSByte(s)
    Case vbCArray:
                        'v_out = CArVal.bytes
    Case vbUserDefinedType:
                        'UdtVal = GridSettingsType_Parse(s)
                        UdtVal = Udt_Parse(s)
                        v_out = UdtVal
    End Select
End Sub

Sub TextAdd(s As String)
    Dim t As String: t = Text2.Text
    Text2.Text = t & s & vbCrLf
End Sub

Private Sub BtnDebugTest1_Click()
    Dim v
    
    v = "255":      Debug_Print VarType2_ToStr(v) & ": " & v
    v = 254:        Debug_Print VarType2_ToStr(v) & ": " & v
    v = CByte(253): Debug_Print VarType2_ToStr(v) & ": " & v
    v = v + 1:      Debug_Print VarType2_ToStr(v) & ": " & v
    v = CDec(252):  Debug_Print VarType2_ToStr(v) & ": " & v
    
    Dim v2$
    
    v2 = 4:         Debug_Print VarType2_ToStr(v2) & ": " & v2
    
End Sub

Private Sub BtnDebugTest2_Click()
    Dim v0: ReDim v0(0 To 10)
    Dim v1: ReDim v1(0 To 10)
    Dim v2: ReDim v2(0 To 10)
    Dim v3: ReDim v3(0 To 10)
    Dim v4: ReDim v4(0 To 10) As Long
    
    v4(0) = 123
    v4(1) = 456
    v4(2) = 879123
    
    v3(0) = v4
    v2(0) = v3
    v1(0) = v2
    v0(0) = v1
    
    v0(0)(0)(0)(0)(0) = 123
    v0(0)(0)(0)(0)(1) = 456
    v0(0)(0)(0)(0)(2) = 789
    Debug_Print VarType2_ToStr(v0)
    'Debug_Print v0(0)(0)(0)(0)(0)

End Sub

Private Sub BtnDebugTest3_Click()
    Dim v0: ReDim v0(0 To 10)
    Dim v1: ReDim v1(0 To 10)
    Dim v2: ReDim v2(0 To 10)
    Dim v3: ReDim v3(0 To 10)
    Dim v4: ReDim v4(0 To 10) As Collection
    
    Set v4(0) = New Collection
    
    v4(0).Add 123
    v4(0).Add 456
    v4(0).Add 879123
    
    v3(0) = v4
    v2(0) = v3
    v1(0) = v2
    v0(0) = v1
    
    Debug_Print VarType2_ToStr(v0)
    'Debug.Print v0(0)(0)(0)(0)(0)(2)
End Sub

Private Sub BtnDebugTest4_Click()
    Dim v0: ReDim v0(0 To 10)
    Dim v1: ReDim v1(0 To 10)
    Dim v2: ReDim v2(0 To 10)
    Dim v3: ReDim v3(0 To 10)
    Dim v4: ReDim v4(0 To 10) As Dummy
    
    Set v4(0) = New Dummy
    
    v3(0) = v4
    v2(0) = v3
    v1(0) = v2
    v0(0) = v1
    
    Debug_Print VarType2_ToStr(v0)
    
    Dim c1 'As Class1
    Set c1 = v4(0)
    'Debug_Print TypeName(c1)
    Debug_Print VarType2_ToStr(c1)
End Sub

Private Sub BtnDebugTest5_Click()
    Dim v1 As Dummy:      Set v1 = New Dummy:      Debug_Print VarType2_ToStr(v1)
    Dim v2 As Collection: Set v2 = New Collection: Debug_Print VarType2_ToStr(v2)
    Dim v3 As New Dummy:                           Debug_Print VarType2_ToStr(v3)
    Dim v4 As New Collection:                      Debug_Print VarType2_ToStr(v4)
    Dim v5:               Set v5 = New Dummy:      Debug_Print VarType2_ToStr(v5)
    Dim v6:               Set v6 = New Collection: Debug_Print VarType2_ToStr(v6)
    Dim v7 As Object:     Set v7 = New Dummy:      Debug_Print VarType2_ToStr(v7)
    Dim v8 As Object:     Set v8 = New Collection: Debug_Print VarType2_ToStr(v8)

End Sub

Private Sub BtnDebugTest6_Click()
    ReDim v1(0 To 3) As Dummy:          Debug_Print VarType2_ToStr(v1)
    Set v1(0) = New Dummy:              Debug_Print VarType2_ToStr(v1)
    
    Dim v2: ReDim v2(0 To 3) As Dummy:  Debug_Print VarType2_ToStr(v2)
    Set v2(0) = New Dummy:              Debug_Print VarType2_ToStr(v2)
    
    Dim v3: ReDim v3(0 To 3):           Debug_Print VarType2_ToStr(v3)
    Set v3(0) = New Dummy:              Debug_Print VarType2_ToStr(v3)
    Set v3 = New Dummy:                 Debug_Print VarType2_ToStr(v3)
    
    ReDim v4(0 To 3) As Collection:     Debug_Print VarType2_ToStr(v4)
    Set v4(0) = New Collection:         Debug_Print VarType2_ToStr(v4)
End Sub

Private Sub BtnDebugTest7_Click()
    ' Type GridSettingsType is from the component "MSFlexGrid Wizard"
    ' C:\... \.. \VB98\VB98\Wizards\FLEXWIZ.OCX
    Dim g As GridSettingsType
    g.AllowColDragging = True
    Debug_Print VarType2_ToStr(g)
    Dim v
    v = g
    Debug_Print VarType2_ToStr(v)
    
End Sub

Private Sub BtnDebugTest8_Click()
    Dim v0: ReDim v0(0 To 10)
    Dim v1: ReDim v1(0 To 10)
    Dim v2: ReDim v2(0 To 10)
    Dim v3: ReDim v3(0 To 10)
    Dim v4: ReDim v4(0 To 10) As Collection
    
    Set v4(0) = New Collection
    
    v4(0).Add 123
    v4(0).Add 456
    v4(0).Add 879123
    
    v3(0) = v4
    v2(0) = v3
    v1(0) = v2
    v0(0) = v1
    
    Debug_Print VarType2_ToStr(v0)
    'Debug.Print v0(0)(0)(0)(0)(0)(2)

End Sub

Sub Debug_Print(s As String)
    If IsInIDE Then
        Debug.Print s
    Else
        Text2.Text = Text2.Text & vbCrLf & s
    End If
End Sub
