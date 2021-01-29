Attribute VB_Name = "MEVT"
Option Explicit
'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/bbb05720-f724-45c7-8d17-f83c3d1a3961
'https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-OAUT/%5bMS-OAUT%5d.pdf
' typedef enum tagVARENUM
' {
'                                                                       V: may be used in a VARIANT
'                                                                         S: may be used in a SAFEARRAY
'                                                                           T: may be used in a TYPEDESC
'Extended VbVarType                             Wert
Public Enum EVbVarType                      'Dec    Hex
    vbEmpty = VbVarType.vbEmpty             '   0   &H0&    VT_EMPTY    V       The type of the contained field is undefined. When this flag is specified, the VARIANT MUST NOT contain a data field. The VARIANT definition is specified in section 2.2.29.2.
    vbNull = VbVarType.vbNull               '   1   &H1&    VT_NULL     V       The type of the contained field is NULL.      When this flag is specified, the VARIANT MUST NOT contain a data field. The VARIANT definition is specified in section 2.2.29.2.
    
                                                                              ' Either the specified type,
                                                                              ' or the type of the element
                                                                              ' or contained field MUST . . .
    vbInteger = VbVarType.vbInteger         '   2   &H2&    VT_I2       V,S,T ' * be a  2-byte signed integer.
    vbLong = VbVarType.vbLong               '   3   &H3&    VT_I4       V,S,T ' * be a  4-byte signed integer.
    vbSingle = VbVarType.vbSingle           '   4   &H4&    VT_R4       V,S,T ' * be a  4-byte IEEE floating-point number.
    vbDouble = VbVarType.vbDouble           '   5   &H5&    VT_R8       V,S,T ' * be an 8-byte IEEE floating-point number.
    vbCurrency = VbVarType.vbCurrency       '   6   &H6&    VT_CY       V,S,T ' * be CURRENCY (see section 2.2.24).
    vbDate = VbVarType.vbDate               '   7   &H7&    VT_DATE     V,S,T ' * be DATE (see section 2.2.25).
    vbString = VbVarType.vbString           '   8   &H8&    VT_BSTR     V,S,T ' * be BSTR (see section 2.2.23).
    vbObject = VbVarType.vbObject           '   9   &H9&    VT_DISPATCH V,S,T ' * be a pointer to IDispatch (see section 3.1.4).
    vbError = VbVarType.vbError             '   10  &HA&    VT_ERROR    V,S,T ' * be HRESULT.
    vbBoolean = VbVarType.vbBoolean         '   11  &HB&    VT_BOOL     V,S,T ' * be VARIANT_BOOL (see section 2.2.27).
    vbVariant = VbVarType.vbVariant         '   12  &HC&    VT_VARIANT  V,S,T ' * be VARIANT (see section 2.2.29). It MUST appear with the bit flag VT_BYREF.
    vbDataObject = VbVarType.vbDataObject   '   13  &HD&    VT_UNKNOWN  V,S,T ' * be a pointer to IUnknown.
    vbDecimal = VbVarType.vbDecimal         '   14  &HE&    VT_DECIMAL  V,S,T ' * be DECIMAL (see section 2.2.26).
                                            '   15  &HF&                      '
    vbSByte = 16                            '   16  &H10&   VT_I1       V,S,T ' * be a 1-byte integer.
    vbByte = VbVarType.vbByte               '   17  &H11&   VT_UI1      V,S,T ' * be a 1-byte unsigned integer.
    vbUInteger = 18                         '   18  &H12&   VT_UI2      V,S,T ' * be a 2-byte unsigned integer.
    vbULong = 19                            '   19  &H13&   VT_UI4      V,S,T ' * be a 4-byte unsigned integer.
    vbLongLong = 20                         '   20  &H14&   VT_I8       V,S,T ' * be an 8-byte  signed integer.
    vbULongLong = 21                        '   21  &H15&   VT_UI8      V,S,T ' * be an 8-byte unsigned integer.
    vbInt = 22                              '   22  &H16&   VT_INT      V,S,T ' * be a 4-byte signed integer.
    vbUInt = 23                             '   23  &H17&   VT_UINT     V,S,T ' * be a 4-byte unsigned integer.
                                                                              ' The specified type MUST . . .
    vbVoid = 24                             '   24  &H18&   VT_VOID     T     ' * be void.
    vbHResult = 25                          '   25  &H19&   VT_HRESULT  T     ' * be HRESULT.
    vbPtr = 26                              '   26  &H1A&   VT_PTR      T     ' * be a unique pointer, as specified in [C706] section 4.2.20.2.
    vbSafeArray = 27                        '   27  &H1B&   VT_SAFEARRAY T    ' * be SAFEARRAY (section 2.2.30).
    vbCArray = 28                           '   28  &H1C&   VT_CARRAY   T     ' * be a fixed-size array.
    vbUserdefined = 29                      '   29  &H1D&   VT_USERDEFINED T  ' * be user defined.
    vbLPStr = 30                            '   30  &H1E&   VT_LPSTR    T     ' * be a NULL-terminated string, as specified in [C706] section 14.3.4.
    vbLPWStr = 31                           '   31  &H1F&   VT_LPWSTR   T     ' * be a zero-terminated string of UNICODE characters, as specified in [C706], section 14.3.4.
    vbWChar = 32                            '   32  &H20&               V,S,T ' self defined :)
                                            '   33  &H21&
                                            '   34  &H22&
                                            '   35  &H23&
    vbUserDefinedType = VbVarType.vbUserDefinedType '36 &H24& VT_RECORD V,S    The type of the element or contained field MUST be a BRECORD (see section 2.2.28.2).
    vbIntPtr = 37                           '   37  &H25&               T      The specified type MUST be either a 4-byte or an 8-byte signed integer. The size of the integer is platform specific and determines the system pointer size value, as specified in section 2.2.21.
    vbUIntPtr = 38                          '   38  &H26&               T      The specified type MUST be either a 4 byte or an 8 byte unsigned integer. The size of the integer is platform specific and determines the system pointer size value, as specified in section 2.2.21.
                                            '
    vbFileTime = 64                         '   64  &H40&   VT_FILETIME T
    vbBlob = 65                             '   65  &H41&   VT_BLOB     T
    vbStream = 66                           '   66  &H42&   VT_STREAM
    vbStorage = 67                          '   67  &H43&   VT_STORAGE
    vbStreamedObject = 68                   '   68  &H44&   VT_STREAMED_OBJECT
    vbStoredObject = 69                     '   69  &H45&   VT_STORED_OBJECT
    vbBlobObject = 70                       '   70  &H46&   VT_BLOB_OBJECT
    vbCF = 71                               '   71  &H47&   VT_CF
    vbCLSID = 72                            '   72  &H48&   VT_CLSID
                                            '
    vbTypeMask = &HFFF&                     ' 4095  &HFFF&  VT_TYPEMASK
    vbIllegalMask = &HFFF&                  ' 4095  &HFFF&  VT_ILLEGALMASKED
    vbVector = &H1000&                      ' 4096  &H1000& VT_VECTOR
    vbArray = VbVarType.vbArray             ' 8192  &H2000& VT_ARRAY    V,S    The type of the element or contained field MUST be a SAFEARRAY (see section 2.2.30.10).
    vbByRef = &H4000&                       '16384  &H4000& VT_BYREF    V,S    The type of the element or contained field MUST be a pointer to one of the types listed in the previous rows of this table. If present, this bit flag MUST appear in a VARIANT discriminant (see section 2.2.28) with one of the previous flags.
    vbReserved = &H8000&                    '32768  &H8000& VT_RESERVED
    vbIllegal = &HFFFF&                     '65535  &HFFFF& VT_ILLEGAL
End Enum
#If False Then
Dim vbSByte: Dim vbUInteger: Dim vbULong: Dim vbLongLong: Dim vbULongLong: Dim vbInt: Dim vbUInt: Dim vbVoid: Dim vbHResult: Dim vbPtr: Dim vbSafeArray: Dim vbCArray: Dim vbUserdefined: Dim vbLPStr: Dim vbLPWStr: Dim vbWChar: Dim vbIntPtr
Dim vbUIntPtr: Dim vbFileTime: Dim vbBlob: Dim vbStream: Dim vbStorage: Dim vbStreamedObject: Dim vbStoredObject: Dim vbBlobObject: Dim vbCF: Dim vbCLSID: Dim vbTypeMask: Dim vbIllegalMask: Dim vbVector: Dim vbByRef: Dim vbIllegal
#End If

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



Public Function VVarType(ByRef aVar As Variant) As EVbVarType
    'Ersatz für die Funktion VarType
    RtlMoveMemory VVarType, ByVal VarPtr(aVar), 2
End Function

Public Function EVarType_ToStr(ByVal vt As EVbVarType) As String
    Dim s As String
    'If vt > &H8000& Then vt = vt - &H8000&
    Select Case vt
    Case vbEmpty:           s = "Empty"
    Case vbNull:            s = "Null"
    Case vbInteger:         s = "Integer"
    Case vbLong:            s = "Long"
    Case vbSingle:          s = "Single"
    Case vbDouble:          s = "Double"
    Case vbCurrency:        s = "Currency"
    Case vbDate:            s = "Date"
    Case vbString:          s = "String"
    Case vbObject:          s = "Object"
    Case vbError:           s = "Error"
    Case vbBoolean:         s = "Boolean"
    Case vbVariant:         s = "Variant"
    Case vbDataObject:      s = "DataObject"
    Case vbDecimal:         s = "Decimal"
                                                 
    Case vbSByte:           s = "SByte"
    Case vbByte:            s = "Byte"
    Case vbUInteger:        s = "UInteger"
    Case vbULong:           s = "ULong"
    Case vbLongLong:        s = "LongLong"
    Case vbULongLong:       s = "ULongLong"
    Case vbInt:             s = "Int "
    Case vbUInt:            s = "UInt "
    Case vbVoid:            s = "Void"
    Case vbHResult:         s = "HResult"
    Case vbPtr:             s = "Ptr"
    Case vbSafeArray:       s = "SafeArray"
    Case vbCArray:          s = "CArray"
    Case vbUserdefined:     s = "Userdefined"
    Case vbLPStr:           s = "LPStr"
    Case vbLPWStr:          s = "LPWStr"
    
    Case vbWChar:           s = "WChar"
    
    Case vbUserDefinedType: s = "UserDefinedType"
    Case vbIntPtr:          s = "IntPtr"
    Case vbUIntPtr:         s = "UIntPtr"
    Case vbFileTime:        s = "FileTime"
    Case vbBlob:            s = "Blob"
    Case vbStream:          s = "Stream"
    Case vbStorage:         s = "Storage"
    Case vbStreamedObject:  s = "StreamedObject"
    Case vbStoredObject:    s = "StoredObject"
    Case vbBlobObject:      s = "BlobObject"
    Case vbCF:              s = "CF"
    Case vbCLSID:           s = "CLSID"
    Case vbTypeMask:        s = "TypeMask"
    Case vbIllegalMask:     s = "IllegalMask"
    Case vbVector:          s = "Vector"
    Case vbArray:           s = "Array"
    Case vbByRef:           s = "ByRef"
    Case vbReserved:        s = "vbReserved"
    Case vbIllegal:         s = "Illegal"
    Case Else:
        If vt And vbVector Then
            If Len(s) Then s = s & " "
            s = s & "Vector"
        End If
        If vt And vbArray Then
            If Len(s) Then s = s & " "
            s = s & "Array"
        End If
        If vt And vbByRef Then
            If Len(s) Then s = s & " "
            s = s & "ByRef"
        End If
    End Select
    EVarType_ToStr = s
End Function

Sub VarTypes_ToCombo(aCmb As ComboBox, ParamArray exclude())
    Dim exc(): exc = exclude
    With aCmb
        '.Clear
        Dim vt As EVbVarType, s As String, c As Long
        For vt = vbEmpty To vbUIntPtr
            s = IIf(ArrayContains(exc, vt), "", EVarType_ToStr(vt))
            If Len(s) Then
                .AddItem s
                .ItemData(c) = vt
                c = c + 1
            End If
        Next
        vt = vbArray
        s = IIf(ArrayContains(exc, vt), "", EVarType_ToStr(vt))
        If Len(s) Then
            .AddItem s:
            .ItemData(c) = vt
        End If
    End With
End Sub

Function ArrayContains(Arr(), var) As Boolean
    Dim v
    For Each v In Arr
        If Not IsEmpty(v) And Not IsMissing(v) Then
            If v = var Then ArrayContains = True: Exit Function
        End If
    Next
End Function

Public Function TUdtVar(var) As TUdtVar
    RtlMoveMemory TUdtVar, var, 16
End Function
Public Sub TUdtVar_Zero(var As TUdtVar)
    RtlZeroMemory var, 16
End Sub

Public Function VarType2_ToStr(var) As String
    'On Error Resume Next
    Dim vt As VbVarType: vt = VVarType(var)
    Dim s As String
    If vt And vbByRef Then
        s = s & EVarType_ToStr(vt)
        vt = vt - vbByRef
    End If
    If vt And vbVector Then
        s = s & EVarType_ToStr(vt)
        vt = vt - vbVector
    End If
    If vt And vbArray Then
        s = s & EVarType_ToStr(vt)
        vt = vt - vbArray
        s = s & " As " & EVarType_ToStr(vt)
        If vt = vbVariant Then
            s = s & "(" & VarType2_ToStr(var(LBound(var))) & ")"
        'ElseIf vt = vbObject Then
        '    s = s & "(As " & TypeName(var(LBound(var))) & ")"
        End If
    Else
        s = s & EVarType_ToStr(vt)
        If vt = vbObject Or vt = vbUserDefinedType Or vt = vbDataObject Then
            s = s & "(As " & TypeName(var) & ")"
        End If
    End If
    VarType2_ToStr = s
    'On Error GoTo 0
End Function

Public Function EVbVarType_Parse(ByVal s As String) As EVbVarType
    s = Trim(s)
    If InStr(1, s, "=") Then s = Trim(Left(s, InStr(1, s, "=")))
    If Left(s, 2) <> "vb" Then s = "vb" & s
    Dim vt As EVbVarType
    Select Case LCase(s)
    Case "vbempty":           vt = EVbVarType.vbEmpty           ' vbEmpty      = 0"
    Case "vbnull":            vt = EVbVarType.vbNull            ' vbNull       = 1"
    Case "vbinteger":         vt = EVbVarType.vbInteger         ' vbInteger    = 2"
    Case "vblong":            vt = EVbVarType.vbLong            ' vbLong       = 3"
    Case "vbsingle":          vt = EVbVarType.vbSingle          ' vbSingle     = 4"
    Case "vbdouble":          vt = EVbVarType.vbDouble          ' vbDouble     = 5"
    Case "vbcurrency":        vt = EVbVarType.vbCurrency        ' vbCurrency   = 6"
    Case "vbdate":            vt = EVbVarType.vbDate            ' vbDate       = 7"
    Case "vbstring":          vt = EVbVarType.vbString          ' vbString     = 8"
    Case "vbobject":          vt = EVbVarType.vbObject          ' vbObject     = 9"
    Case "vberror":           vt = EVbVarType.vbError           ' vbError      = 10"
    Case "vbboolean":         vt = EVbVarType.vbBoolean         ' vbBoolean    = 11"
    Case "vbvariant":         vt = EVbVarType.vbVariant         ' vbVariant    = 12"
    Case "vbdataobject":      vt = EVbVarType.vbDataObject      ' vbDataObject = 13"
    Case "vbdecimal":         vt = EVbVarType.vbDecimal         ' vbDecimal    = 14"
                                                                '  15?
    Case "vbsbyte":           vt = EVbVarType.vbSByte           '  16   ' &H10& '   VT_I1          = 0x0010,
    Case "vbbyte":            vt = EVbVarType.vbByte            '  17   ' &H11& '   VT_UI1         = 0x0011,"
    Case "vbuinteger":        vt = EVbVarType.vbUInteger        '  18   ' &H12& '   VT_UI2         = 0x0012,"
    Case "vbulong":           vt = EVbVarType.vbULong           '  19   ' &H13& '   VT_UI4         = 0x0013,"
    Case "vblonglong":        vt = EVbVarType.vbLongLong        '  20   ' &H14& '   VT_I8          = 0x0014,"
    Case "vbulonglong":       vt = EVbVarType.vbULongLong       '  21   ' &H15& '   VT_UI8         = 0x0015,"
    Case "vbint":             vt = EVbVarType.vbInt             '  22   ' &H16& '   VT_INT         = 0x0016,"
    Case "vbuint":            vt = EVbVarType.vbUInt            '  23   ' &H17& '   VT_UINT        = 0x0017,"
    Case "vbvoid":            vt = EVbVarType.vbVoid            '  24   ' &H18& '   VT_VOID        = 0x0018,"
    Case "vbhresult":         vt = EVbVarType.vbHResult         '  25   ' &H19& '   VT_HRESULT     = 0x0019,"
    Case "vbptr":             vt = EVbVarType.vbPtr             '  26   ' &H1A& '   VT_PTR         = 0x001A,"
    Case "vbsafearray":       vt = EVbVarType.vbSafeArray       '  27   ' &H1B& '   VT_SAFEARRAY   = 0x001B,"
    Case "vbcArray":          vt = EVbVarType.vbCArray          '  28   ' &H1C& '   VT_CARRAY      = 0x001C,"
    Case "vbuserdefined":     vt = EVbVarType.vbUserdefined     '  29   ' &H1D& '   VT_USERDEFINED = 0x001D,"
    Case "vblpstr":           vt = EVbVarType.vbLPStr           '  30   ' &H1E& '   VT_LPSTR       = 0x001E,"
    Case "vblpwstr":          vt = EVbVarType.vbLPWStr          '  31   ' &H1F& '   VT_LPWSTR      = 0x001F,"
    Case "vbwchar":           vt = EVbVarType.vbWChar           '  32   ' &H20&                               'selbst definiert, nicht offiziell"
                                                                '  33   ' &H21& '
                                                                '  34   ' &H22& '
                                                                '  35   ' &H23& '
    Case "vbuserdefinedtype": vt = EVbVarType.vbUserDefinedType '  36   ' &H24& '   VT_RECORD      = 0x0024,"
    Case "vbintptr":          vt = EVbVarType.vbIntPtr          '  37   ' &H25& '   VT_INT_PTR     = 0x0025,"
    Case "vbuintptr":         vt = EVbVarType.vbUIntPtr         '  38   ' &H26& '   VT_UINT_PTR    = 0x0026,"
    Case "vbarray":           vt = EVbVarType.vbArray           '8192   '&H2000&'   VT_ARRAY       = 0x2000,"
    Case "vbbyref":           vt = EVbVarType.vbByRef          '16384   '&H4000&'   VT_BYREF       = 0x4000
    'Case Else:
    End Select
    EVbVarType_Parse = vt
End Function

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
