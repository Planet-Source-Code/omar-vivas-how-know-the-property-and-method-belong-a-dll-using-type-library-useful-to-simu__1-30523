Attribute VB_Name = "Module1"
Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes
  If SearchData And &H80000000 Then
    GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
  Else
    GetSearchType = SearchData \ &H1000000 And &HFF&
  End If
End Function

Function GetTypeInfoNumber(ByVal SearchData As Long) As Integer
  GetTypeInfoNumber = SearchData And &HFFF&
End Function

Function GetLibNum(ByVal SearchData As Long) As Integer
  SearchData = SearchData And &H7FFFFFFF
  GetLibNum = ((SearchData \ &H2000& And &H7) * &H100&) Or _
               (SearchData \ &H10000 And &HFF&)
End Function

Function GetHidden(ByVal SearchData As Long) As Boolean
    If SearchData And &H1000& Then GetHidden = True
End Function

Function BuildSearchData( _
   ByVal TypeInfoNumber As Integer, _
   ByVal SearchTypes As TliSearchTypes, _
   Optional ByVal LibNum As Integer, _
   Optional ByVal Hidden As Boolean = False) As Long
  If SearchTypes And &H80 Then
    BuildSearchData = _
      (TypeInfoNumber And &H1FFF&) Or _
      ((SearchTypes And &H7F) * &H1000000) Or &H80000000
  Else
    BuildSearchData = _
      (TypeInfoNumber And &H1FFF&) Or _
      (SearchTypes * &H1000000)
  End If

  If LibNum Then
    BuildSearchData = BuildSearchData Or _
      ((LibNum And &HFF) * &H10000) Or _
      ((LibNum And &H700) * &H20&)
  End If
  If Hidden Then
    BuildSearchData = BuildSearchData Or &H1000&
  End If
End Function

Function PrototypeMember( _
  TLInf As TypeLibInfo, _
  ByVal SearchData As Long, _
  ByVal InvokeKinds As InvokeKinds, _
  Optional ByVal MemberId As Long = -1, _
  Optional ByVal MemberName As String) As String
Dim pi As ParameterInfo
Dim fFirstParameter As Boolean
Dim fIsConstant As Boolean
Dim fByVal As Boolean
Dim retVal As String
Dim ConstVal As Variant
Dim strTypeName As String
Dim VarTypeCur As Integer
Dim fDefault As Boolean, fOptional As Boolean, fParamArray As Boolean
Dim TIType As TypeInfo
Dim TIResolved As TypeInfo
Dim TKind As TypeKinds
  With TLInf
    fIsConstant = GetSearchType(SearchData) And tliStConstants
    With .GetMemberInfo(SearchData, InvokeKinds, MemberId, MemberName)
      If fIsConstant Then
        retVal = "Const "
      ElseIf InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then
        Select Case .ReturnType.VarType
          Case VT_VOID, VT_HRESULT
            retVal = "Sub "
          Case Else
            retVal = "Function "
        End Select
      Else
        retVal = "Property "
      End If
      retVal = retVal & .Name
      With .Parameters
        If .Count Then
          retVal = retVal & "("
          fFirstParameter = True
          fParamArray = .OptionalCount = -1
          For Each pi In .Me
            If Not fFirstParameter Then
              retVal = retVal & ", "
            End If
            fFirstParameter = False
            fDefault = pi.Default
            fOptional = fDefault Or pi.Optional
            If fOptional Then
              If fParamArray Then
                'This will be the only optional parameter
                retVal = retVal & "[ParamArray "
              Else
                retVal = retVal & "["
              End If
            End If
            With pi.VarTypeInfo
              Set TIType = Nothing
              Set TIResolved = Nothing
              TKind = TKIND_MAX
              VarTypeCur = .VarType
              If (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
              'If Not .TypeInfoNumber Then 'This may error, don't use here
                On Error Resume Next
                Set TIType = .TypeInfo
                If Not TIType Is Nothing Then
                  Set TIResolved = TIType
                  TKind = TIResolved.TypeKind
                  Do While TKind = TKIND_ALIAS
                    TKind = TKIND_MAX
                    Set TIResolved = TIResolved.ResolvedType
                    If Err Then
                      Err.Clear
                    Else
                      TKind = TIResolved.TypeKind
                    End If
                  Loop
                End If
                Select Case TKind
                  Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
                    fByVal = .PointerLevel = 1
                  Case TKIND_RECORD
                    'Records not passed ByVal in VB
                    fByVal = False
                  Case Else
                    fByVal = .PointerLevel = 0
                End Select
                If fByVal Then retVal = retVal & "ByVal "
                retVal = retVal & pi.Name
                If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then retVal = retVal & "()"
                If TIType Is Nothing Then 'Error
                  retVal = retVal & " As ?"
                Else
                  If .IsExternalType Then
                    retVal = retVal & " As " & _
                             .TypeLibInfoExternal.Name & "." & TIType.Name
                  Else
                    retVal = retVal & " As " & TIType.Name
                  End If
                End If
                On Error GoTo 0
              Else
                If .PointerLevel = 0 Then retVal = retVal & "ByVal "
                retVal = retVal & pi.Name
                If VarTypeCur <> vbVariant Then
                  strTypeName = TypeName(.TypedVariant)
                  If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                    retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                  Else
                    retVal = retVal & " As " & strTypeName
                  End If
                End If
              End If
              If fOptional Then
                If fDefault Then
                  retVal = retVal & ProduceDefaultValue(pi.DefaultValue, TIResolved)
                End If
                retVal = retVal & "]"
              End If
            End With
          Next
          retVal = retVal & ")"
        End If
      End With
      If fIsConstant Then
        ConstVal = .Value
        retVal = retVal & " = " & ConstVal
        Select Case VarType(ConstVal)
          Case vbInteger, vbLong
            If ConstVal < 0 Or ConstVal > 15 Then
              retVal = retVal & " (&H" & Hex$(ConstVal) & ")"
            End If
        End Select
      Else
        With .ReturnType
          VarTypeCur = .VarType
          If VarTypeCur = 0 Or (VarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
          'If Not .TypeInfoNumber Then 'This may error, don't use here
            On Error Resume Next
            If Not .TypeInfo Is Nothing Then
              If Err Then 'Information not available
                retVal = retVal & " As ?"
              Else
                If .IsExternalType Then
                  retVal = retVal & " As " & _
                           .TypeLibInfoExternal.Name & "." & .TypeInfo.Name
                Else
                  retVal = retVal & " As " & .TypeInfo.Name
                End If
              End If
            End If
            If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then retVal = retVal & "()"
            On Error GoTo 0
          Else
            Select Case VarTypeCur
              Case VT_VARIANT, VT_VOID, VT_HRESULT
              Case Else
                strTypeName = TypeName(.TypedVariant)
                If VarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                  retVal = retVal & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                Else
                  retVal = retVal & " As " & strTypeName
                End If
            End Select
          End If
        End With
      End If
      PrototypeMember = retVal & vbCrLf & "  " & _
                        "Member of " & TLInf.Name & "." & _
                        TLInf.GetTypeInfo(SearchData And &HFFFF&).Name & _
                        vbCrLf & "  " & .HelpString
    End With
  End With
End Function




Private Function ProduceDefaultValue(DefVal As Variant, ByVal TI As TypeInfo) As String
Dim lTrackVal As Long
Dim MI As MemberInfo
Dim TKind As TypeKinds
    If TI Is Nothing Then
        Select Case VarType(DefVal)
            Case vbString
                If Len(DefVal) Then
                    ProduceDefaultValue = """" & DefVal & """"
                End If
            Case vbBoolean 'Always show for Boolean
                ProduceDefaultValue = DefVal
            Case vbDate
                If DefVal Then
                    ProduceDefaultValue = "#" & DefVal & "#"
                End If
            Case Else 'Numeric Values
                If DefVal <> 0 Then
                    ProduceDefaultValue = DefVal
                End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = TI.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
            Set TI = TI.ResolvedType
            If Err = 0 Then TKind = TI.TypeKind
            On Error GoTo 0
        Loop
        If TI.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each MI In TI.Members
                If MI.Value = lTrackVal Then
                    ProduceDefaultValue = MI.Name
                    Exit For
                End If
            Next
        End If
    End If
End Function
