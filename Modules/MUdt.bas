Attribute VB_Name = "MUdt"
Option Explicit
'herein just 3 helper-functions for the udt for convenient-reasons

Function Udt_ToStr(var) As String
    'works for every UDType from a typelibrary
    Dim tudt As TUdtVar: tudt = TUdtVar(var)
    Dim pUdt As Long: pUdt = tudt.pvData
    With tudt.RecInfo
        Dim s As String: s = .GetName() & "{"
        Dim c As Long: .GetFieldNames c, 0
        ReDim snames(0 To c - 1) As String
        .GetFieldNames c, ByVal VarPtr(snames(0))
        Dim i As Long, v
        .GetField ByVal pUdt, snames(i), v
        s = s & snames(i) & ":=" & v
        For i = 1 To c - 1
            .GetField ByVal pUdt, snames(i), v
            s = s & "; " & snames(i) & ":=" & v
        Next
        s = s & "}"
    End With
    TUdtVar_Zero tudt
    Udt_ToStr = s
End Function

Function Udt_Parse(s As String)
    Dim var
    Select Case True
    Case BeginsWith(s, "GridsettingsType")
        'here comes the point where all the magic happens
        'if you assign an udt to a Variant, VB does the following:
        '* a new udt will be created allocated on the heap,
        '* the pointer to the heap-mem will be stored in the Variant at pos 9-12
        '* an IRecordInfo-obj will be created and
        '* will be stored next to the pointer in the Variant at pos 13-16
        Dim gst As GridSettingsType: var = gst
    Case Else
        'other ud-types here
    End Select
    Dim tudt As TUdtVar: tudt = TUdtVar(var)
    Dim sa() As String: sa = Split(s, "{")
    Dim sv   As String: sv = sa(1)
    If Right(sv, 1) = "}" Then sv = Mid(sv, 1, Len(sv) - 1)
    Dim svs() As String: svs = Split(sv, ";")
    Dim pUdt As Long: pUdt = tudt.pvData
    With tudt.RecInfo
        Dim i As Long, u As Long: u = UBound(svs)
        Dim nvs() As String
        Dim n As String, v
        For i = 0 To u
            nvs = Split(svs(i), ":=")
            n = Trim$(nvs(0))
            'GetField liefert den Wert und den *Typ* des Feldes, den wir hier brauchen
            .GetField ByVal pUdt, n, v
            .PutField VbCallType.VbLet, ByVal pUdt, n, CXVar(nvs(1), VVarType(v))
        Next
    End With
    'do not forget!!!
    TUdtVar_Zero tudt
    Udt_Parse = var
    Debug.Print Udt_ToStr(var)
End Function

Function BeginsWith(s As String, sval As String) As Boolean
    If Len(s) < Len(sval) Then Exit Function
    Dim sn As String: sn = Left(s, Len(sval))
    BeginsWith = (StrComp(sn, sval, vbTextCompare) = 0)
End Function

