Attribute VB_Name = "MEVbVarType"
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
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)

Public Function VVarType(ByRef aVar As Variant) As EVbVarType
    'Ersatz für die Funktion VarType
    RtlMoveMemory VVarType, ByVal VarPtr(aVar), 2
End Function

Public Function EVbVarType_ToStr(ByVal vt As EVbVarType) As String
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
    EVbVarType_ToStr = s
End Function

Sub EVbVarTypes_ToCombo(aCmb As ComboBox, ParamArray exclude())
    Dim exc(): exc = exclude
    With aCmb
        '.Clear
        Dim vt As EVbVarType, s As String, c As Long
        For vt = vbEmpty To vbUIntPtr
            s = IIf(ArrayContains(exc, vt), "", EVbVarType_ToStr(vt))
            If Len(s) Then
                .AddItem s
                .ItemData(c) = vt
                c = c + 1
            End If
        Next
        vt = vbArray
        s = IIf(ArrayContains(exc, vt), "", EVbVarType_ToStr(vt))
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


