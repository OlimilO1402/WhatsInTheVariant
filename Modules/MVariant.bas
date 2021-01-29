Attribute VB_Name = "MVariant"
Option Explicit
'all Variant functions of Oleaut32.dll
'https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-varabs
'HRESULT VarAbs(LPVARIANT pvarIn, LPVARIANT pvarResult);
Public Declare Function VarAbs Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long

'HRESULT VarAdd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarAdd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarAnd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarAnd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarDiv(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarDiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarEqv(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarEqv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarFix(LPVARIANT pvarIn, LPVARIANT pvarResult);
Public Declare Function VarFix Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long

'HRESULT VarFormat(LPVARIANT pvarIn, LPOLESTR pstrFormat, int iFirstDay, int iFirstWeek, ULONG dwFlags, BSTR *pbstrOut);
'Public Declare Function VarFormat Lib "oleaut32" (ByRef pvarIn As Any, ByVal pstrFormat As String, ByVal iFirstDay As Long, ByVal iFirstWeek As Long, ByVal dwFlags As Long, ByRef pbstrOut As String) As Long

'INT VariantTimeToDosDateTime(DOUBLE vtime, USHORT *pwDosDate, USHORT *pwDosTime);
'INT  VariantTimeToSystemTime(DOUBLE vtime, lpSystemTime lpSystemTime);

'HRESULT VarIdiv(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarIdiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarImp (LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarImp Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarInt(LPVARIANT pvarIn, LPVARIANT pvarResult);
Public Declare Function VarInt Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long

'HRESULT VarMod(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarMod Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarMul(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarMul Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarNeg(LPVARIANT pvarIn, LPVARIANT pvarResult);
Public Declare Function VarNeg Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long

'HRESULT VarNot(LPVARIANT pvarIn, LPVARIANT pvarResult);
Public Declare Function VarNot Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long

'HRESULT VarOr(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarOr Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarPow(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarPow Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarSub(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarSub Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'HRESULT VarXor(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
Public Declare Function VarXor Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long

'WINOLECTLAPI OleLoadPictureFile(VARIANT varFileName, LPDISPATCH *lplpdispPicture);
Public Declare Function OleLoadPictureFile Lib "oleaut32" (ByVal varFileName As Variant, ByRef lplpdispPicture As IPictureDisp) As Long
'Public Declare Function OleLoadPictureFile Lib "oleaut32" (ByVal varFileName As Variant, ByRef lplpdispPicture As Long) As Long

Public Declare Function VariantChangeTypeEx Lib "oleaut32" (ByRef Dst As Any, ByRef Src As Any, ByVal Lcid As Long, ByVal wflags As Integer, ByVal vt As Integer) As Long

Public Function VAbs(v):       Dim hr As Long: hr = VarAbs(v, VAbs):        End Function

Public Function VAdd(v1, v2):  Dim hr As Long: hr = VarAdd(v1, v2, VAdd):   End Function

Public Function VAnd(v1, v2):  Dim hr As Long: hr = VarAnd(v1, v2, VAnd):   End Function

Public Function VDiv(v1, v2):  Dim hr As Long: hr = VarDiv(v1, v2, VDiv):   End Function

Public Function VEqv(v1, v2):  Dim hr As Long: hr = VarEqv(v1, v2, VEqv):   End Function

Public Function VFix(v):       Dim hr As Long: hr = VarFix(v, VFix):        End Function

Public Function VIdiv(v1, v2): Dim hr As Long: hr = VarIdiv(v1, v2, VIdiv): End Function

Public Function VImp(v1, v2):  Dim hr As Long: hr = VarImp(v1, v2, VImp):   End Function

Public Function VInt(v):       Dim hr As Long: hr = VarInt(v, VInt):        End Function

Public Function VMod(v1, v2):  Dim hr As Long: hr = VarMod(v1, v2, VMod):   End Function

Public Function VMul(v1, v2):  Dim hr As Long: hr = VarMul(v1, v2, VMul):   End Function

Public Function VNeg(v):       Dim hr As Long: hr = VarNeg(v, VNeg):        End Function

Public Function VNot(v):       Dim hr As Long: hr = VarNot(v, VNot):        End Function

Public Function VOr(v1, v2):   Dim hr As Long: hr = VarOr(v1, v2, VOr):     End Function

Public Function VPow(v1, v2):  Dim hr As Long: hr = VarPow(v1, v2, VPow):   End Function

Public Function VSub(v1, v2):  Dim hr As Long: hr = VarSub(v1, v2, VSub):   End Function

Public Function VXOr(v1, v2):  Dim hr As Long: hr = VarXor(v1, v2, VXOr):   End Function

'Function Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek = vbSunday], [FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1]) As String
'Public Function VFormat(Expression, sFormat As String, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, Optional FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
'    Dim hr As Long: hr = VarFormat(Expression, sFormat & vbNullChar, FirstDayOfWeek, FirstWeekOfYear, 1, VFormat)
'End Function
Public Function LoadPic(FileName As String) As StdPicture
    Dim hr As Long: hr = OleLoadPictureFile(FileName, LoadPic)
    'Debug.Print Hex$(hr) '0x800A01E1 = CTL_E_INVALIDPICTURE
    'Debug.Print Hex$(Err.LastDllError)
End Function
Public Sub TestVariantOperators()

    Dim v1: v1 = -123456
    Dim v2: v2 = 123456
    Dim v
    
    v = VAbs(v1)
    Debug.Print v
    
    v = VAdd(v1, v2)
    Debug.Print v
    
    v = VAnd(v1, v2)
    Debug.Print v
    
    v = VDiv(v1, v2)
    Debug.Print v
    
    v1 = -123456.789456
    v = VFix(v1)
    Debug.Print v
    
    v = Now
    
    'Debug.Print VFormat(v, "dd.mmm.yyyy") 'hmmm?
    
    Debug.Print Format(v, "dd.mmm.yyyy")
    
    Dim d As Double: d = v
    
    Debug.Print Format(d, "dd.mmm.yyyy")

End Sub

Public Function IsInIDE() As Boolean
Try: On Error GoTo Catch
    Debug.Print 1 / 0
    Exit Function
Catch: IsInIDE = True
End Function
