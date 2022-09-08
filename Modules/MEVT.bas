Attribute VB_Name = "MEVT"
Option Explicit

'see:
'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-irecordinfo
Public Type TUdtVar
    vt   As Integer ' 2
    res1 As Integer ' 2
    res2 As Long    ' 4
    pvData  As Long 'Pointer to the new udt on the heap
    RecInfo As IRecordInfoVB
End Type

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLen As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal BytLen As Long)

'HRESULT VariantChangeTypeEx(
'  VARIANTARG       *pvargDest,
'  const VARIANTARG *pvarSrc,
'  LCID             lcid,
'  USHORT           wFlags,
'  VarType vt
');
'wflags:
'VARIANT_NOVALUEPROP    =  1: Prevents the function from attempting to coerce an object to a fundamental type by getting the Value property.
'                             Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.
Public Const VARIANT_NOVALUEPROP    As Integer = 1
'VARIANT_ALPHABOOL      =  2: Converts a VT_BOOL value to a string containing either "True" or "False".
Public Const VARIANT_ALPHABOOL      As Integer = 2
'VARIANT_NOUSEROVERRIDE =  4: For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.
Public Const VARIANT_NOUSEROVERRIDE As Integer = 4
'VARIANT_LOCALBOOL      = 16: For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer.
Public Const VARIANT_LOCALBOOL      As Integer = 16

'Return-Value:
'S_OK:    Success.
Public Const S_OK                As Long = 0
'DISP_E_BADVARTYPE   : The variant type is not a valid type of variant.
Public Const DISP_E_BADVARTYPE   As Long = &H80020008
'DISP_E_OVERFLOW     : The data pointed to by pvarSrc does not fit in the destination type.
Public Const DISP_E_OVERFLOW     As Long = &H8002000A
'DISP_E_TYPEMISMATCH : The argument could not be coerced to the specified type.
Public Const DISP_E_TYPEMISMATCH As Long = &H80020005
'E_INVALIDARG        : One of the arguments is not valid.
Public Const E_INVALIDARG        As Long = &H80070057
'E_OUTOFMEMORY       : Insufficient memory to complete the operation.
Public Const E_OUTOFMEMORY       As Long = &H8007000E

Public Declare Function VariantChangeTypeEx Lib "oleaut32" (ByRef Dst As Any, ByRef Src As Any, ByVal Lcid As Long, ByVal wflags As Integer, ByVal vt As Integer) As Long

Public Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long

'INVOKE_FUNC           1: Der Member wird mit der üblichen Aufrufsyntax für Funktionen aufgerufen.
'INVOKE_PROPERTYGET    2: Die Funktion wird mit der üblichen Syntax für den Zugriff auf Eigenschaften aufgerufen.
'INVOKE_PROPERTYPUT    4: Die Funktion wird mit Syntax für das Zuweisen von Eigenschaftswerten aufgerufen.
'INVOKE_PROPERTYPUTREF 8: Die Funktion wird mit Syntax für das Zuweisen von Verweisen auf Eigenschaften aufgerufen.
'Enum VbCallType
'    VbMethod = 1
'    VbGet = 2
'    VbLet = 4
'    VbSet = 8
'End End



Public Function TUdtVar(var) As TUdtVar
    RtlMoveMemory TUdtVar, var, 16
End Function
Public Sub TUdtVar_Zero(var As TUdtVar)
    RtlZeroMemory var, 16
End Sub

Function VarValue_ToStr(aVar)
    Dim s As String
    Dim vt As EVbVarType: vt = VVarType(aVar)
    If vt And vbByRef Then vt = vt - vbByRef
    Select Case vt
    Case vbEmpty:   s = ""
    Case vbNull:    s = "0"
    Case vbError:   s = ErrorStr(aVar) '"Error"
    Case vbInteger, vbUInteger, vbLong, vbULong, vbLongLong, vbULongLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString, vbBoolean, vbVariant, vbDecimal, vbByte, vbSByte
        s = CStr(aVar)
    Case vbError
        s = CStr(aVar)
    Case vbDataObject
        s = "DataObject" 'CStr(aVar)
    Case vbUserDefinedType
        If TypeName(aVar) = "GridSettingsType" Then
            Dim t As GridSettingsType: t = aVar
            's = GridSettingsType_ToStr(t)
            s = Udt_ToStr(t)
        End If
    'Case vbArray:      ReDim vArr(0 To 3)
    Case vbObject
        If TypeName(aVar) = "Dummy" Then
            Dim d As Dummy: Set d = aVar
            s = d.ToStr
        End If
    Case Else
        If vt And vbArray = vbArray Then
            Dim u As Long: u = UBound(aVar)
            Dim i As Long
            For i = 0 To u - 1
                s = s & VarValue_ToStr(aVar(i)) & ", "
            Next
            s = s & VarValue_ToStr(aVar(u))
        Else
            s = "Error, type could not be recognized: " & vt
        End If
    End Select
    VarValue_ToStr = s
End Function

Function ErrorStr(e) As String
Try: On Error Resume Next 'GoTo Catch
    Dim l As Long: l = CLng(e)
    ErrorStr = "(" & l & ") " & Error$(l)
Catch:
End Function

Public Sub AssignVar(ByRef Dst, ByRef Src As Variant)
    'never assign two variants like this: aVar = bVar
    'because bVar could also be an object with a default property
    'so you will get only the default property
    'e.g. if the default property is of type String aVar will
    'also be of type String, not of object
    If IsObject(Src) Then
        Set Dst = Src
    Else
        Dst = Src
    End If
End Sub

Sub CreateArray(vArr, ByVal vt As EVbVarType, s As String)

    Dim sep As String: sep = GetSeparator(s)
    Dim sa() As String: sa = Split(s, sep)
    Dim u As Long: u = Max(UBound(sa), 0)
    Dim i As Long
    Select Case vt
    'Case vbEmpty:    'ReDim vArr(0 To 3) As Empty ' nopo
    'Case vbNull:     'ReDim vArr(0 To 3) As Error ' nopo
    Case vbInteger, vbUInteger, vbWChar
                                ReDim vArr(0 To u) As Integer
                                If vt = vbInteger Or vt = vbUInteger Then
                                    For i = 0 To u: vArr(i) = CInt(sa(i)): Next
                                ElseIf vt = vbWChar Then
                                    For i = 0 To u: vArr(i) = AscW(Mid(s, i, 1)): Next
                                End If
    Case vbLong, vbULong, vbHResult, vbInt, vbUInt, vbPtr, vbIntPtr, vbUIntPtr, vbVoid
                                ReDim vArr(0 To u) As Long
                                For i = 0 To u: vArr(i) = CLng(sa(i)): Next
                                
    Case vbSingle:              ReDim vArr(0 To u) As Single
                                For i = 0 To u: vArr(i) = CSng(sa(i)): Next
    Case vbDouble:              ReDim vArr(0 To u) As Double
                                For i = 0 To u: vArr(i) = CDbl(sa(i)): Next
    Case vbCurrency, vbLongLong, vbULongLong
                                ReDim vArr(0 To u) As Currency
                                For i = 0 To u: vArr(i) = CCur(sa(i)): Next
    Case vbDate:                ReDim vArr(0 To u) As Date
                                For i = 0 To u: vArr(i) = CDate(sa(i)): Next
    Case vbString, vbLPStr, vbLPWStr
                                ReDim vArr(0 To u) As String
                                For i = 0 To u: vArr(i) = sa(i): Next
    Case vbObject:              ReDim vArr(0 To u) As Object
                                'For i = 0 To u: vArr(i) = CDate(sa(i)): Next
    Case vbError:               ReDim vArr(0 To u) As ErrObject
    Case vbBoolean:             ReDim vArr(0 To u) As Boolean
                                For i = 0 To u: vArr(i) = CBool(sa(i)): Next
    Case vbVariant:             ReDim vArr(0 To u) As Variant
    Case vbDataObject:          ReDim vArr(0 To u) As DataObject
    Case vbDecimal:             ReDim vArr(0 To u) As Variant
                                For i = 0 To u: vArr(i) = CDec(sa(i)): Next
    Case vbSByte, vbByte:       ReDim vArr(0 To u) As Byte
                                For i = 0 To u: vArr(i) = CByte(sa(i)): Next
    Case vbUserDefinedType:     ReDim vArr(0 To u) As MSFlexGridWizard.GridSettingsType
                                'For i = 0 To u: vArr(i) = MUdt.GridSettingsType_Parse(sa(i)): Next
                                For i = 0 To u: vArr(i) = MUdt.Udt_Parse(sa(i)): Next
    Case vbArray:               ReDim vArr(0 To u)
    Case vbSafeArray:
    Case vbUserdefined:
    End Select
End Sub

Function GetSeparator(ByVal str As String) As String
    Dim s As String
    If Len(str) = 0 Then GetSeparator = ", ": Exit Function
    Dim u1 As Long: u1 = UBound(Split(str, "; "))
    Dim u2 As Long: u2 = UBound(Split(str, ", "))
    Dim u3 As Long: u3 = UBound(Split(str, " "))
    Dim u4 As Long: u4 = UBound(Split(str, ";"))
    Dim u  As Long: u = Max(Max(Max(u1, u2), u3), u4)
    Select Case u
    Case u1: s = "; "
    Case u2: s = ", "
    Case u3: s = " "
    Case u4: s = ";"
    End Select
    GetSeparator = s
End Function

Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function

Function TryRecognizeVar(s As String, r_out As Variant) As Boolean
'just for demonstration, maybe not complete
'must not be compared with type inference
' Int32.Min =          -2147483648
' Int32.Max =           2147483647
'UInt32.Min =                    0
'UInt32.Max =           4294967295
' Int64.Min = -9223372036854775808
' Int64.Max =  9223372036854775807
'UInt64.Min =                    0
'UInt64.Max = 18446744073709551615
Try: On Error GoTo Catch
    Dim c As Currency
    If IsNumeric(s) Then
        s = Replace(s, ".", ",")
        Dim d: d = CDec(s)
        If IsInt(s) Then
            If -128 < d And d < 0 Then
                r_out = CSByte(s)
            ElseIf 0 <= d And d <= 255 Then
                r_out = CByte(d)
            ElseIf -32768 <= d And d <= 32767 Then
                r_out = CInt(d)
            ElseIf 32768 <= d And d <= 65535 Then
                r_out = CUInt(s)
            ElseIf -2147483647 <= d And d <= 2147483647 Then
                r_out = CLng(c)
            ElseIf 2147483648# <= d And d <= 4294967294# Then
                r_out = CULng(s)
            ElseIf CDec("-9223372036854775807") <= d And d <= CDec("9223372036854775807") Then
                r_out = CLngLng(s)
            ElseIf 0 <= d And d <= CDec("18446744073709551615") Then
                r_out = CULngLng(s)
            End If
        Else
            If IsSingle(s) Then
                r_out = CSng(d)
            ElseIf IsDouble(s) Then
                r_out = CDbl(d)
            Else
                r_out = d
            End If
        End If
    Else
        If Len(s) = 0 Then
            r_out = Empty
        Else
            Select Case UCase(s)
            Case "TRUE", "WAHR":    r_out = True  ', "OK", "RIGHT", "RICHTIG"
            Case "FALSE", "FALSCH": r_out = False ', "WRONG"
            Case "EMPTY":           r_out = Empty
            Case "NULL":            r_out = Null
            Case Else: r_out = s
            End Select
        End If
    End If
    TryRecognizeVar = True: Exit Function
Catch:
End Function

Function IsInt(ByVal s As String) As Boolean
    If InStr(1, s, ",") Then Exit Function
    If InStr(1, s, ".") Then Exit Function
    IsInt = Int(CDec(s)) = CDec(s)
End Function

Function IsSingle(s As String) As Boolean
    If Len(s) > 9 Then Exit Function 'may be wrong!
    Dim sng As Single
    IsSingle = Single_TryParse(s, sng)
End Function
Function Single_TryParse(ByVal s As String, s_out As Single) As Boolean
Try: On Error GoTo Catch
    If IsNumeric(s) Then
        s = Replace(s, ",", ".")
        s_out = CSng(Val(s))
        Single_TryParse = True
    End If
Catch:
End Function

Function IsDouble(s As String) As Boolean
    If Len(s) > 17 Then Exit Function 'may be wrong!
    Dim dbl As Double
    IsDouble = Double_TryParse(s, dbl)
End Function
Function Double_TryParse(ByVal s As String, d_out As Double) As Boolean
Try: On Error GoTo Catch
    If IsNumeric(s) Then
        s = Replace(s, ",", ".")
        d_out = Val(s)
        Double_TryParse = True
    End If
Catch:
End Function

Function Decimal_TryParse(ByVal s As String, d_out) As Boolean
Try: On Error GoTo Catch
    If IsNumeric(s) Then
        s = Replace(s, ".", ",") 'hier "," und "." andersrum
        d_out = CDec(s)
        Decimal_TryParse = True
    End If
Catch:
End Function

Function Currency_TryParse(ByVal s As String, c_out As Currency) As Boolean
Try: On Error GoTo Catch
    If IsNumeric(s) Then
        s = Replace(s, ".", ",") 'hier "," und "." andersrum
        c_out = CCur(s)
        Currency_TryParse = True
    End If
Catch:
End Function

Function ErrHandler(obj, fncname As String, msg As String) As VbMsgBoxResult
    Dim objName As String
    objName = IIf(IsObject(obj) Or IsRecord(obj), TypeName(obj), obj)
    ErrHandler = MsgBox("Fehler: " & Err.Number & " in " & objName & "::" & fncname & IIf(Len(msg), vbCrLf & msg, "") & IIf(Len(Err.Description), vbCrLf & Err.Description, ""))
End Function

Function IsRecord(obj) As Boolean
    IsRecord = VarType(obj) = vbUserDefinedType
End Function

'##############################'    VariantChangeTypeEx    '##############################'
Function CXVar(Value, vt As EVbVarType)
    Dim f As Long
    Select Case vt
    Case EVbVarType.vbBoolean: f = VARIANT_LOCALBOOL Or VARIANT_ALPHABOOL
    End Select
    Static Lcid As Long: If Lcid = 0 Then Lcid = GetUserDefaultLCID
    Dim hr As Long: hr = VariantChangeTypeEx(CXVar, Value, Lcid, f, CInt(vt))
    If hr = S_OK Then Exit Function
    Dim s As String
    Select Case hr
    Case DISP_E_BADVARTYPE:   s = "The variant type is not a valid type of variant."
    Case DISP_E_OVERFLOW:     s = "The data pointed to by pvarSrc does not fit in the destination type."
    Case DISP_E_TYPEMISMATCH: s = "The argument could not be coerced to the specified type."
    Case E_INVALIDARG:        s = "One of the arguments is not valid."
    Case E_OUTOFMEMORY:       s = "Insufficient memory to complete the operation."
    End Select
    MsgBox s
End Function

'signed byte
Function CSByte(Value)
    'Dim hr As Long: hr = VariantChangeTypeEx(CSByte, Value, Lcid, 0, CInt(EVbVarType.vbSByte))
    CSByte = CXVar(Value, EVbVarType.vbSByte)
    'or maybe call CXVar Value, EVbVarType.vbSByte
End Function

'unsigned byte
'Function CByte(ByVal s As String)

'signed int16
'Function CInt(Value)

'unsigned int16
Function CUInt(Value)
    'Dim hr As Long: hr = VariantChangeTypeEx(CUInt, Value, Lcid, 0, CInt(EVbVarType.vbUInteger))
    CUInt = CXVar(Value, EVbVarType.vbUInteger)
    'or maybe call CXVar Value, EVbVarType.vbUInteger
End Function

'signed int32
'function CLng(Value)

'unsigned int32
Function CULng(Value)
    'Dim hr As Long: hr = VariantChangeTypeEx(CULng, Value, Lcid, 0, CInt(EVbVarType.vbULong))
    CULng = CXVar(Value, EVbVarType.vbULong)
    'or maybe call CXVar Value, EVbVarType.vbULong
End Function

'signed int64
Function CLngLng(Value)
    'Dim hr As Long: hr = VariantChangeTypeEx(CLngLng, Value, Lcid, 0, CInt(EVbVarType.vbLongLong))
    CLngLng = CXVar(Value, EVbVarType.vbLongLong)
    'or maybe call CXVar Value, EVbVarType.vbLongLong
End Function

'unsigned int64
Function CULngLng(Value)
    'Dim hr As Long: hr = VariantChangeTypeEx(CULngLng, Value, Lcid, 0, CInt(EVbVarType.vbULongLong))
    CULngLng = CXVar(Value, EVbVarType.vbULongLong)
    'or maybe call CXVar Value, EVbVarType.vbULongLong
End Function
