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
    vbObject = VbVarType.vbObject           '   9   &H9&    VT_DISPATCH V,S,T ' * be a pointer to IDispatch (see section 3.1.4).  'IUnknown+IDispatch
    vbError = VbVarType.vbError             '   10  &HA&    VT_ERROR    V,S,T ' * be HRESULT.                                        'andalso IsMissing (Nicht vorhanden)
    vbBoolean = VbVarType.vbBoolean         '   11  &HB&    VT_BOOL     V,S,T ' * be VARIANT_BOOL (see section 2.2.27).
    vbVariant = VbVarType.vbVariant         '   12  &HC&    VT_VARIANT  V,S,T ' * be VARIANT (see section 2.2.29). It MUST appear with the bit flag VT_BYREF.
    vbDataObject = VbVarType.vbDataObject   '   13  &HD&    VT_UNKNOWN  V,S,T ' * be a pointer to IUnknown.
    vbDecimal = VbVarType.vbDecimal         '   14  &HE&    VT_DECIMAL  V,S,T ' * be DECIMAL (see section 2.2.26).
                                            '   15  &HF&                      '
    vbSByte = 16                            '   16  &H10&   VT_I1       V,S,T ' * be a 1-byte integer.
    vbByte = VbVarType.vbByte               '   17  &H11&   VT_UI1      V,S,T ' * be a 1-byte unsigned integer.
    vbUInteger = 18                         '   18  &H12&   VT_UI2      V,S,T ' * be a 2-byte unsigned integer.
    vbULong = 19                            '   19  &H13&   VT_UI4      V,S,T ' * be a 4-byte unsigned integer.
    'vbLongPtr = 20 'in VBA7+Win64
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
    vbIntPtr = 37                           '   37  &H25&               T      The specified type MUST be either a 4-byte or an 8-byte   signed integer. The size of the integer is platform specific and determines the system pointer size value, as specified in section 2.2.21.
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

#If win64 Then
    Public Const SizeOf_Variant As Long = 24
#Else
    Public Const SizeOf_Variant As Long = 16
#End If

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
    Case vbInt:             s = "Int"
    Case vbUInt:            s = "UInt"
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
            vt = vt Xor EVbVarType.vbVector
            s = s & " " & EVbVarType_ToStr(vt)
        End If
        If vt And vbArray Then
            If Len(s) Then s = s & " "
            s = s & "Array"
            vt = vt Xor EVbVarType.vbArray
            s = s & " " & EVbVarType_ToStr(vt)
        End If
        If vt And vbByRef Then
            If Len(s) Then s = s & " "
            s = s & "ByRef"
            vt = vt Xor EVbVarType.vbByRef
            s = s & " " & EVbVarType_ToStr(vt)
        End If
        'hmm should we not remove array, vector or byref and go again?
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

Function ArrayContains(Arr(), Var) As Boolean
    Dim v
    For Each v In Arr
        If Not IsEmpty(v) And Not IsMissing(v) Then
            If v = Var Then ArrayContains = True: Exit Function
        End If
    Next
End Function

Public Function EVbVarType_Parse(ByVal s As String) As EVbVarType
    s = LCase(Trim(s))
    If InStr(1, s, "=") Then s = Trim(Left(s, InStr(1, s, "=")))
    'If Left(s, 2) <> "vb" And Left(s, 3) <> "vt_" Then
    '    s = "vb" & s
    'End If
    Dim vt As EVbVarType
    Select Case s
    Case "empty", "vbempty", "vt_empty":                              vt = EVbVarType.vbEmpty                       ' vbEmpty      = 0"
    Case "any", "null", "vbnull", "vt_null":                          vt = EVbVarType.vbNull            ' vbNull       = 1"
    Case "int16", "integer", "vbinteger", "vt_i2":                    vt = EVbVarType.vbInteger         ' vbInteger    = 2"
    Case "int32", "long", "vblong", "vt_i4":                          vt = EVbVarType.vbLong            ' vbLong       = 3"
    Case "float32", "flt32", "float", "single", "vbsingle", "vt_r4":  vt = EVbVarType.vbSingle          ' vbSingle     = 4"
    Case "float64", "flt64", "double", "vbdouble", "vt_r8":           vt = EVbVarType.vbDouble          ' vbDouble     = 5"
    Case "currency", "vbcurrency", "vt_cy":                           vt = EVbVarType.vbCurrency        ' vbCurrency   = 6"
    Case "date", "vbdate", "vt_date":                                 vt = EVbVarType.vbDate            ' vbDate       = 7"
    Case "datetime", "vbdatetime", "vt_datetime":                     vt = EVbVarType.vbDate            ' vbDate       = 7"
    Case "string", "vbstring", "vt_bstr":                             vt = EVbVarType.vbString          ' vbString     = 8"
    Case "object", "vbobject", "vt_dispatch":                         vt = EVbVarType.vbObject          ' vbObject     = 9"
    Case "error", "vberror", "vt_error":                              vt = EVbVarType.vbError           ' vbError      = 10"
    Case "bool", "boolean", "vbboolean", "vt_bool":                   vt = EVbVarType.vbBoolean         ' vbBoolean    = 11"
    Case "variant", "vbvariant", "vt_variant":                        vt = EVbVarType.vbVariant         ' vbVariant    = 12"
    Case "dataobject", "vbdataobject":                                vt = EVbVarType.vbDataObject      ' vbDataObject = 13"
    Case "decimal", "vbdecimal", "vt_decimal":                        vt = EVbVarType.vbDecimal         ' vbDecimal    = 14"
                                                                '  15?
    Case "sint8", "sbyte", "vbsbyte", "vt_i1":                        vt = EVbVarType.vbSByte           '  16   ' &H10& '   VT_I1          = 0x0010,
    Case "int8", "byte", "vbbyte", "vt_ui1":                          vt = EVbVarType.vbByte            '  17   ' &H11& '   VT_UI1         = 0x0011,"
    Case "uint16", "uinteger", "vbuinteger", "vt_ui2":                vt = EVbVarType.vbUInteger        '  18   ' &H12& '   VT_UI2         = 0x0012,"
    Case "uint32", "ulong", "vbulong", "vt_ui4":                      vt = EVbVarType.vbULong           '  19   ' &H13& '   VT_UI4         = 0x0013,"
    Case "int64", "longlong", "vblonglong", "vt_i8":                  vt = EVbVarType.vbLongLong        '  20   ' &H14& '   VT_I8          = 0x0014,"
    Case "uint64", "ulonglong", "vbulonglong", "vt_ui8":              vt = EVbVarType.vbULongLong       '  21   ' &H15& '   VT_UI8         = 0x0015,"
    Case "int", "vbint", "vt_int":                                    vt = EVbVarType.vbInt             '  22   ' &H16& '   VT_INT         = 0x0016,"
    Case "uint", "vbuint", "vt_uint":                                 vt = EVbVarType.vbUInt            '  23   ' &H17& '   VT_UINT        = 0x0017,"
    Case "void", "vbvoid", "vt_void":                                 vt = EVbVarType.vbVoid            '  24   ' &H18& '   VT_VOID        = 0x0018,"
    Case "hresult", "vbhresult", "vt_hresult":                        vt = EVbVarType.vbHResult         '  25   ' &H19& '   VT_HRESULT     = 0x0019,"
    Case "ptr", "vbptr", "vt_ptr":                                    vt = EVbVarType.vbPtr             '  26   ' &H1A& '   VT_PTR         = 0x001A,"
    Case "safearray", "vbsafearray", "vt_safearray":                  vt = EVbVarType.vbSafeArray       '  27   ' &H1B& '   VT_SAFEARRAY   = 0x001B,"
    Case "carray", "vbcArray", "vt_carray":                           vt = EVbVarType.vbCArray          '  28   ' &H1C& '   VT_CARRAY      = 0x001C,"
    Case "userdefined", "vbuserdefined", "vt_userdefined":            vt = EVbVarType.vbUserdefined                '  29   ' &H1D& '   VT_USERDEFINED = 0x001D,"
    Case "lpstr", "vblpstr", "vt_lpstr":                              vt = EVbVarType.vbLPStr           '  30   ' &H1E& '   VT_LPSTR       = 0x001E,"
    Case "lpwstr", "vblpwstr", "vt_lpwstr":                           vt = EVbVarType.vbLPWStr          '  31   ' &H1F& '   VT_LPWSTR      = 0x001F,"
    Case "wchar", "vbwchar", "vt_wchar":                              vt = EVbVarType.vbWChar           '  32   ' &H20&                               'self defined inofficial"
                                                                '  33   ' &H21& '
                                                                '  34   ' &H22& '
                                                                '  35   ' &H23& '
    Case "userdefinedtype", "vbuserdefinedtype":                      vt = EVbVarType.vbUserDefinedType '  36   ' &H24& '   VT_RECORD      = 0x0024,"
    Case "intptr", "vbintptr", "vt_int_ptr":                          vt = EVbVarType.vbIntPtr          '  37   ' &H25& '   VT_INT_PTR     = 0x0025,"
    Case "uintptr", "vbuintptr", "vt_uint_ptr":                       vt = EVbVarType.vbUIntPtr         '  38   ' &H26& '   VT_UINT_PTR    = 0x0026,"
    
    
                                            '
    Case "filetime", "vbfiletime", "vt_filetime":                     vt = EVbVarType.vbFileTime        '  64   ' &H40& '   VT_FILETIME T
    Case "blob", "vbblob", "vt_blob":                                 vt = EVbVarType.vbBlob            '  65   ' &H41& '   VT_BLOB     T
    Case "stream", "vbstream", "vt_stream":                           vt = EVbVarType.vbStream          '  66   ' &H42& '   VT_STREAM
    Case "storage", "vbStorage", "vt_storage":                        vt = EVbVarType.vbStorage '= 67                          '   67  &H43&   VT_STORAGE
    Case "streamedobject", "vbstreamedobject", "vt_streamed_object":  vt = EVbVarType.vbStreamedObject '= 68                   '   68  &H44&   VT_STREAMED_OBJECT
    Case "storedobject", "vbstoredobject", "vt_stored_object":        vt = EVbVarType.vbStoredObject '= 69                     '   69  &H45&   VT_STORED_OBJECT
    Case "blobobject", "vbblobobject", "vt_blob_object":              vt = EVbVarType.vbBlobObject '= 70                       '   70  &H46&   VT_BLOB_OBJECT
    Case "clipboard", "vbcf", "vt_cf":                                vt = vbCF
    Case "guid", "clsid", "vbclsid", "vt_clsid":                      vt = vbCLSID
    Case "vector", "vbvector", "vt_vector":                           vt = EVbVarType.vbVector          '4096   '&H1000&'   VT_VECTOR      = 0x1000,"
    Case "multivalue", "array", "vbarray", "vt_array":                vt = EVbVarType.vbArray           '8192   '&H2000&'   VT_ARRAY       = 0x2000,"
    Case "byref", "vbbyref", "vt_byref":                              vt = EVbVarType.vbByRef          '16384   '&H4000&'   VT_BYREF       = 0x4000
    Case "buffer":                                                    vt = EVbVarType.vbArray Or EVbVarType.vbByte
    Case Else
        Dim sa() As String: sa = Split(s, " ")
        Dim i As Long, u As Long: u = UBound(sa)
        If i = u Then
            If s = sa(i) Then Exit Function
            vt = vt Or EVbVarType_Parse(sa(i))
        Else
            If u >= 0 Then
                For i = 0 To u
                    If Len(sa(i)) Then
                        vt = vt Or EVbVarType_Parse(sa(i))
                    End If
                Next
            End If
        End If
    End Select
    EVbVarType_Parse = vt
    
'https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum
'VARENUM
'  VT_EMPTY = 0,
'  VT_NULL = 1,
'  VT_I2 = 2,
'  VT_I4 = 3,
'  VT_R4 = 4,
'  VT_R8 = 5,
'  VT_CY = 6,
'  VT_DATE = 7,
'  VT_BSTR = 8,
'  VT_DISPATCH = 9,
'  VT_ERROR = 10,
'  VT_BOOL = 11,
'  VT_VARIANT = 12,
'  VT_UNKNOWN = 13,
'  VT_DECIMAL = 14,
'  VT_I1 = 16,
'  VT_UI1 = 17,
'  VT_UI2 = 18,
'  VT_UI4 = 19,
'  VT_I8 = 20,
'  VT_UI8 = 21,
'  VT_INT = 22,
'  VT_UINT = 23,
'  VT_VOID = 24,
'  VT_HRESULT = 25,
'  VT_PTR = 26,
'  VT_SAFEARRAY = 27,
'  VT_CARRAY = 28,
'  VT_USERDEFINED = 29,
'  VT_LPSTR = 30,
'  VT_LPWSTR = 31,
'  VT_RECORD = 36,
'  VT_INT_PTR = 37,
'  VT_UINT_PTR = 38,
'  VT_FILETIME = 64,
'  VT_BLOB = 65,
'  VT_STREAM = 66,
'  VT_STORAGE = 67,
'  VT_STREAMED_OBJECT = 68,
'  VT_STORED_OBJECT = 69,
'  VT_BLOB_OBJECT = 70,
'  VT_CF = 71,
'  VT_CLSID = 72,
'  VT_VERSIONED_STREAM = 73,
'  VT_BSTR_BLOB = 0xfff,
'  VT_VECTOR = 0x1000,
'  VT_ARRAY = 0x2000,
'  VT_BYREF = 0x4000,
'  VT_RESERVED = 0x8000,
'  VT_ILLEGAL = 0xffff,
'  VT_ILLEGALMASKED = 0xfff,
'  VT_TYPEMASK = 0xfff
End Function

Public Function VarType2_ToStr(Var) As String
    'On Error Resume Next
    Dim vt As VbVarType: vt = VVarType(Var)
    Dim s As String
    If vt And vbByRef Then
        s = s & EVbVarType_ToStr(vt)
        vt = vt - vbByRef
    End If
    If vt And vbVector Then
        s = s & EVbVarType_ToStr(vt)
        vt = vt - vbVector
    End If
    If vt And vbArray Then
        s = s & EVbVarType_ToStr(vt)
        vt = vt - vbArray
        s = s & " As " & EVbVarType_ToStr(vt)
        If vt = vbVariant Then
            s = s & "(" & VarType2_ToStr(Var(LBound(Var))) & ")"
        'ElseIf vt = vbObject Then
        '    s = s & "(As " & TypeName(var(LBound(var))) & ")"
        End If
    Else
        s = s & EVbVarType_ToStr(vt)
        If vt = vbObject Or vt = vbUserDefinedType Or vt = vbDataObject Then
            s = s & "(As " & TypeName(Var) & ")"
        End If
    End If
    VarType2_ToStr = s
    'On Error GoTo 0
End Function

