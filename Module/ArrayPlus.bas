Attribute VB_Name = "ArrayPlus"
'#################################################
'##############  ArrayPlus by Arfu  ##############
'##############        v1.4         ##############
'#################################################

Option Explicit

Public Function Length(list As Variant) As Long
    If IsArray(list) Then Length = UBound(list) + 1
End Function

Public Function MinOf(list As Variant) As Single
    Dim i%
    If IsArray(list) Then
        For i = LBound(list) To UBound(list)
            If IsNumeric(list(i)) Then
                If i = 0 Or list(i) < MinOf Then MinOf = list(i)
            End If
        Next
    End If
End Function

Public Function MaxOf(list As Variant) As Single
    Dim i%
    If IsArray(list) Then
        For i = LBound(list) To UBound(list)
            If IsNumeric(list(i)) Then
                If i = 0 Or list(i) > MaxOf Then MaxOf = list(i)
            End If
        Next
    End If
End Function

Public Function AT(list As Variant, Optional ByVal Index As Single = 0, Optional ByVal ReturnIndex As Boolean = False)
    If IsArray(list) Then
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
    End If
End Function

Public Function Insert(list As Variant, Item As Variant, Optional ByVal Index As Long)
    Dim i As Single
    If IsArray(list) Then
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
    End If
End Function

Public Function Remove(list As Variant, ByVal Value)
    Dim i As Single, Index As Single
    If IsArray(list) And Includes(list, Value) Then
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
            Exit Function
        End If
    End If
End Function

Public Function Pop(list As Variant, Optional ByVal Index As Single)
    Dim i As Single
    If IsArray(list) Then
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
            Exit Function
        End If
    End If
End Function

Public Function Includes(list As Variant, Item As Variant) As Boolean
    Dim i As Single
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If IsObject(Item) And IsObject(list(i)) Then
                If list(i) Is Item Then Includes = True
            ElseIf Not IsObject(Item) And Not IsObject(list(i)) Then
                If list(i) = Item Then Includes = True
            End If
        Next
    End If
End Function

Public Function IndexOf(list As Variant, Item As Variant) As Long
    Dim i As Single
    If IsArray(list) Then
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
    End If
End Function

Public Function CountOf(list As Variant, Item As Variant) As Long
    Dim i As Single
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If IsObject(Item) And IsObject(list(i)) Then
                If list(i) Is Item Then CountOf = CountOf + 1
            ElseIf Not IsObject(Item) And Not IsObject(list(i)) Then
                If list(i) = Item Then CountOf = CountOf + 1
            End If
        Next
    End If
End Function

Public Function Reverse(list As Variant)
    Dim handlerlist As Variant, i As Single
    handlerlist = Array()
    If IsArray(list) Then
        For i = UBound(list) To 0 Step -1
            Insert handlerlist, list(i)
        Next
        list = handlerlist
    End If
End Function


Public Function ConcatOf(List1 As Variant, List2 As Variant)
    Dim i As Single
    If IsArray(List1) And IsArray(List2) Then
        For i = 0 To UBound(List2)
            Insert List1, List2(i)
        Next
    End If
End Function

Public Function Shuffle(list)
    Dim handler As Variant, randarr As Variant, i As Single
    If IsArray(list) Then
        handler = Array()
        For i = 0 To UBound(list)
            randarr = RandArray(list)
            Insert handler, randarr
            Remove list, randarr
        Next
        list = handler
    End If
End Function

Public Function Clear(list As Variant)
    If IsArray(list) Then list = Array(Empty)
End Function

Public Function RandArray(list As Variant)
    Randomize
    If IsArray(list) Then RandArray = list(Int((UBound(list) + 1) * Rnd + 0))
End Function

Public Function Reduce(list As Variant, ByVal Weight As Single, Optional ByVal Right As Boolean = False)
    Dim i As Single
    If IsArray(list) Then
        If Right Then
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
    End If
End Function

Public Function Swap(list As Variant, ByVal Index1 As Integer, ByVal Index2 As Integer)
    Dim tmp As Variant
    If IsArray(list) Then
        Index1 = AT(list, Index1, True)
        Index2 = AT(list, Index2, True)
        tmp = list(Index1)
        list(Index1) = list(Index2)
        list(Index2) = tmp
    End If
End Function

Public Function Slice(list As Variant, ByVal StartPos As Integer, ByVal EndPos As Integer)
    Dim sliced As Variant, i As Single
    sliced = Array()
    If IsArray(list) Then
        StartPos = AT(list, StartPos, True)
        EndPos = AT(list, EndPos, True)
        For i = StartPos To EndPos
            Insert sliced, list(i)
        Next
        Slice = sliced
    End If
End Function

Public Function Map(list As Variant, ByVal Func As String)
    Dim i As Single, maparray As Variant
    maparray = Array()
    If IsArray(list) Then
        For i = 0 To UBound(list)
            Insert maparray, Application.Run(Func, list(i), i)
        Next
        Map = maparray
    End If
End Function

Public Function Find(list As Variant, ByVal Func As String)
    Dim i As Single
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If Application.Run(Func, list(i), i) = True Then
                Find = list(i)
                Exit Function
            End If
        Next
    End If
End Function

Public Function FindIndex(list As Variant, ByVal Func As String) As Long
    Dim i As Single
    FindIndex = -1
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If Application.Run(Func, list(i), i) = True Then
                FindIndex = i
                Exit Function
            End If
        Next
    End If
End Function

Public Function Filter(list As Variant, ByVal Func As String)
    Dim i As Single, filtered As Variant
    filtered = Array()
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If Application.Run(Func, list(i), i) = True Then Insert filtered, list(i)
        Next
        Filter = filtered
    End If
End Function

Public Function Every(list As Variant, ByVal Func As String) As Boolean
    Dim i As Single
    Every = False
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If Application.Run(Func, list(i)) = False Then Exit Function
        Next
        Every = True
    End If
End Function

Public Function Some(list As Variant, ByVal Func As String) As Boolean
    Dim i As Single
    Some = True
    If IsArray(list) Then
        For i = 0 To UBound(list)
            If Application.Run(Func, list(i), i) = True Then Exit Function
        Next
        Some = False
    End If
End Function

'Function by Victor Gabriel
Public Sub QuickSort(vArray As Variant, arrLbound As Long, arrUbound As Long)
    
    Dim pivotVal As Variant
    Dim vSwap    As Variant
    Dim tmpLow   As Long
    Dim tmpHi    As Long
    
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
    
    If (arrLbound < tmpHi) Then QuickSort vArray, arrLbound, tmpHi
    If (tmpLow < arrUbound) Then QuickSort vArray, tmpLow, arrUbound
    
End Sub

