Attribute VB_Name = "ArrayPlus"
'#################################################
'##############  ArrayPlus by Arfu  ##############
'##############        v1.5.1       ##############
'#################################################

Option Explicit

Public Function LengthOf(list As Variant) As Long
    If IsArray(list) Then LengthOf = UBound(list) + 1
End Function

Public Function MinOf(list As Variant) As Double
    Dim i%
    If Not IsArray(list) Then Exit Function
    For i = LBound(list) To UBound(list)
        If IsNumeric(list(i)) Then
            If i = 0 Or list(i) < MinOf Then MinOf = list(i)
        End If
    Next
End Function

Public Function MaxOf(list As Variant) As Double
    Dim i%
    If Not IsArray(list) Then Exit Function
    For i = LBound(list) To UBound(list)
        If IsNumeric(list(i)) Then
            If i = 0 Or list(i) > MaxOf Then MaxOf = list(i)
        End If
    Next
End Function

Public Function AT(list As Variant, Optional ByVal Index As Long = 0, Optional ByVal ReturnIndex As Boolean = False)
    If Not IsArray(list) Then Exit Function
    If UBound(list) + 1 > 0 Then
        If Index >= 0 And Index <= UBound(list) Then
            AT = list(Index)
        ElseIf Index > UBound(list) Then
            AT = list(UBound(list))
        ElseIf Index < 0 And Abs(Index) <= UBound(list) + 1 Then
            AT = list((UBound(list) + 1) + Index)
        Else
            AT = list(0)
        End If
    End If
    If ReturnIndex Then AT = IndexOf(list, AT)
End Function

Public Sub Insert(list As Variant, Item As Variant, Optional ByVal Index As Long)
    Dim i As Long
    If Not IsArray(list) Then Exit Sub
    Index = AT(list, Index, True)
    If UBound(list) < Index Or Index = 0 Then
        ReDim Preserve list(UBound(list) + 1)
        If IsObject(Item) Then
            Set list(UBound(list)) = Item
        Else
            list(UBound(list)) = Item
        End If
    ElseIf UBound(list) >= Index Then
        ReDim Preserve list(UBound(list) + 1)
        For i = UBound(list) - 1 To Index Step -1
            If IsObject(list(i)) Then
                Set list(i + 1) = list(i)
            Else
                list(i + 1) = list(i)
            End If
        Next
        If IsObject(Item) Then
            Set list(Index) = Item
        Else
            list(Index) = Item
        End If
    End If
End Sub

Public Sub Remove(list As Variant, ByVal Value)
    Dim i As Long, Index As Long
    If IsArray(list) And IncludesOf(list, Value) Then
        Index = IndexOf(list, Value)
        If UBound(list) = Index And UBound(list) + 1 > 1 Then
            ReDim Preserve list(UBound(list) - 1)
        ElseIf UBound(list) > Index Then
            For i = Index To UBound(list)
                If i <> UBound(list) Then
                    If IsObject(list(i + 1)) Then
                        Set list(i) = list(i + 1)
                    Else
                        list(i) = list(i + 1)
                    End If
                End If
            Next
            ReDim Preserve list(UBound(list) - 1)
        ElseIf UBound(list) >= Index And UBound(list) + 1 = 1 Then
            ReDim list(UBound(list) - UBound(list))
        Else
            Exit Sub
        End If
    End If
End Sub

Public Sub Pop(list As Variant, Optional ByVal Index As Long)
    Dim i As Long
    If Not IsArray(list) Then Exit Sub
    If IsMissing(Index) = False Then Index = AT(list, Index, True)
    If IsMissing(Index) Or Index > UBound(list) Or UBound(list) = Index And UBound(list) + 1 > 1 Then
        ReDim Preserve list(UBound(list) - 1)
    ElseIf UBound(list) > Index Then
        For i = Index To UBound(list)
            If i <> UBound(list) Then
                If IsObject(list(i + 1)) Then
                    Set list(i) = list(i + 1)
                Else
                    list(i) = list(i + 1)
                End If
            End If
        Next
        ReDim Preserve list(UBound(list) - 1)
    ElseIf UBound(list) >= Index And UBound(list) + 1 = 1 Then
        ReDim list(UBound(list) - UBound(list))
    Else
        Exit Sub
    End If
End Sub

Public Function IncludesOf(list As Variant, Item As Variant) As Boolean
    Dim i As Long
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If IsObject(Item) And IsObject(list(i)) Then
            If list(i) Is Item Then IncludesOf = True
        ElseIf Not IsObject(Item) And Not IsObject(list(i)) Then
            If list(i) = Item Then IncludesOf = True
        End If
    Next
End Function

Public Function IndexOf(list As Variant, Item As Variant) As Long
    Dim i As Long
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If IsObject(Item) And IsObject(list(i)) Then
            If list(i) Is Item Then
                IndexOf = i
                Exit Function
            Else
                IndexOf = -1
            End If
        ElseIf Not IsObject(Item) And Not IsObject(list(i)) Then
            If list(i) = Item Then
                IndexOf = i
                Exit Function
            Else
                IndexOf = -1
            End If
        End If
    Next
End Function

Public Function CountOf(list As Variant, Item As Variant) As Long
    Dim i As Long
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If IsObject(Item) And IsObject(list(i)) Then
            If list(i) Is Item Then CountOf = CountOf + 1
        ElseIf Not IsObject(Item) And Not IsObject(list(i)) Then
            If list(i) = Item Then CountOf = CountOf + 1
        End If
    Next
End Function

Public Sub Reverse(list As Variant)
    Dim handlerlist As Variant, i As Long
    handlerlist = Array()
    If Not IsArray(list) Then Exit Sub
    For i = UBound(list) To 0 Step -1
        Insert handlerlist, list(i)
    Next
    list = handlerlist
End Sub


Public Sub ConcatOf(List1 As Variant, List2 As Variant)
    Dim i As Long
    If IsArray(List1) And IsArray(List2) Then
        For i = 0 To UBound(List2)
            Insert List1, List2(i)
        Next
    End If
End Sub

Public Sub Shuffle(list)
    Dim handler As Variant, randarr As Variant, i As Long
    If Not IsArray(list) Then Exit Sub
    handler = Array()
    For i = 0 To UBound(list)
        randarr = RandomArray(list)
        Insert handler, randarr
        Remove list, randarr
    Next
    list = handler
End Sub

Public Sub Clear(list As Variant)
    If IsArray(list) Then list = Array(Empty)
End Sub

Public Function RandomArray(list As Variant)
    Randomize
    If IsArray(list) Then RandomArray = list(Int((UBound(list) + 1) * Rnd + 0))
End Function

Public Sub Reduce(list As Variant, ByVal Weight As Long, Optional ByVal right As Boolean = False)
    Dim i As Long
    If Not IsArray(list) Then Exit Sub
    If right Then
        If Weight - 1 >= UBound(list) Then
            ReDim list(UBound(list) - UBound(list))
        Else
            ReDim Preserve list(UBound(list) - Weight)
        End If
    Else
        If Weight - 1 >= UBound(list) Then
            ReDim list(UBound(list) - UBound(list))
        Else
            For i = 0 To Weight - 1
                Pop list, i
            Next
        End If
    End If
End Sub

Public Function Swap(list As Variant, ByVal Index1 As Long, ByVal Index2 As Long)
    Dim tmp As Variant
    If Not IsArray(list) Then Exit Function
    Index1 = AT(list, Index1, True)
    Index2 = AT(list, Index2, True)
    tmp = list(Index1)
    list(Index1) = list(Index2)
    list(Index2) = tmp
End Function

Public Function Slice(list As Variant, ByVal StartPos As Long, ByVal EndPos As Long)
    Dim sliced As Variant, i As Long
    sliced = Array()
    If Not IsArray(list) Then Exit Function
    StartPos = AT(list, StartPos, True)
    EndPos = AT(list, EndPos, True)
    For i = StartPos To EndPos
        Insert sliced, list(i)
    Next
    Slice = sliced
End Function

Public Function Map(list As Variant, ByVal Func As String)
    Dim i As Long, maparray As Variant
    maparray = Array()
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        Insert maparray, Application.Run(Func, list(i), i)
    Next
    Map = maparray
End Function

Public Function Find(list As Variant, ByVal Func As String)
    Dim i As Long
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If Application.Run(Func, list(i), i) = True Then
            Find = list(i)
            Exit Function
        End If
    Next
End Function

Public Function FindIndex(list As Variant, ByVal Func As String) As Long
    Dim i As Long
    FindIndex = -1
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If Application.Run(Func, list(i), i) = True Then
            FindIndex = i
            Exit Function
        End If
    Next
End Function

Public Function Filter(list As Variant, ByVal Func As String)
    Dim i As Long, filtered As Variant
    filtered = Array()
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If Application.Run(Func, list(i), i) = True Then Insert filtered, list(i)
    Next
    Filter = filtered
End Function

Public Function Every(list As Variant, ByVal Func As String) As Boolean
    Dim i As Long
    Every = False
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If Application.Run(Func, list(i)) = False Then Exit Function
    Next
    Every = True
End Function

Public Function Some(list As Variant, ByVal Func As String) As Boolean
    Dim i As Long
    Some = True
    If Not IsArray(list) Then Exit Function
    For i = 0 To UBound(list)
        If Application.Run(Func, list(i), i) = True Then Exit Function
    Next
    Some = False
End Function

Public Sub Quicksort(vArray As Variant, Optional arrLbound As Long = 0, Optional arrUbound As Long)
    Dim pivotVal As Variant
    Dim vSwap    As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long
    
    If Not IsArray(vArray) Then Exit Sub
    If arrUbound <= -1 Then arrUbound = UBound(vArray)
    
    tmpLow = arrLbound
    tmpHi = arrUbound
    pivotVal = vArray((arrLbound + arrUbound) \ 2)

    While (tmpLow <= tmpHi)
       While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
          tmpLow = tmpLow + 1
       Wend

       While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
          tmpHi = tmpHi - 1
       Wend

       If (tmpLow <= tmpHi) Then
          vSwap = vArray(tmpLow)
          vArray(tmpLow) = vArray(tmpHi)
          vArray(tmpHi) = vSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend

    If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi
    If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound
End Sub


