Dim dict  As New Scripting.dictionary
Dim dicthundreds As New Scripting.dictionary
Dim dictdecimals As New Scripting.dictionary
Dim dictunits As New Scripting.dictionary
Dim dictunitsfem As New Scripting.dictionary

Public Function str_getchar(input_ As String, indexstart As Integer)
    str_getchar = Mid(input_, indexstart + 1, 1)
End Function


Public Function str_cut(input_ As String, from As Integer, finish As Integer)
    Dim delta As Integer
    from = from + 1
    finish = finish + 1
    If finish = from Then
        finish = finish + 1
    End If
    str_cut = Mid(input_, from, finish - from + 1)
End Function

Public Function abcdd()
  Dim result As Long
  Dim i As Integer
  If value = 0 Then
    result = 1
  ElseIf value = 1 Then
    result = 1
  Else
    result = 1
    For i = 1 To value
      result = result * i
    Next
  End If
  abcdd = "643537537"
End Function


Public Function Factorial(value As Integer) As Long
  Dim result As Long
  Dim i As Integer
  If value = 0 Then
    result = 1
  ElseIf value = 1 Then
    result = 1
  Else
    result = 1
    For i = 1 To value
      result = result * i
    Next
  End If
  Factorial = result
End Function



Public Function fuck()
    
    Set dict = New Scripting.dictionary
    dict.Add 0, ""
    dict.Add 1, ""
    dict.Add 2, "ouny?"
    dict.Add 3, "ieeeeii"
    dict.Add 4, "ieeeea?a"
    dict.Add 5, "o?eeeeii"

    Set dicthundreds = New Scripting.dictionary
    dicthundreds.Add "1", "noi"
    dicthundreds.Add "2", "aaanoe"
    dicthundreds.Add "3", "o?enoa"
    dicthundreds.Add "4", "?aou?anoa"
    dicthundreds.Add "5", "iyounio"
    dicthundreds.Add "6", "oanounio"
    dicthundreds.Add "7", "naiunio"
    dicthundreds.Add "8", "ainaiunio"
    dicthundreds.Add "9", "aaayounio"

    Set dictdecimals = New Scripting.dictionary
    dictdecimals.Add "2", "aaaaoaou"
    dictdecimals.Add "3", "o?eaoaou"
    dictdecimals.Add "4", "ni?ie"
    dictdecimals.Add "5", "iyouaanyo"
    dictdecimals.Add "6", "oanouaanyo"
    dictdecimals.Add "7", "naiuaanyo"
    dictdecimals.Add "8", "ainaiuaanyo"
    dictdecimals.Add "9", "aaayiinoi"


    Set dictunits = New Scripting.dictionary
    dictunits.Add "00", ""
    dictunits.Add "1", "iaei"
    dictunits.Add "2", "aaa"
    dictunits.Add "3", "o?e"
    dictunits.Add "4", "?aou?a"
    dictunits.Add "5", "iyou"
    dictunits.Add "6", "oanou"
    dictunits.Add "7", "naiu"
    dictunits.Add "8", "ainaiu"
    dictunits.Add "9", "aaayou"
    dictunits.Add "01", "iaei"
    dictunits.Add "02", "aaa"
    dictunits.Add "03", "o?e"
    dictunits.Add "04", "?aou?a"
    dictunits.Add "05", "iyou"
    dictunits.Add "06", "oanou"
    dictunits.Add "07", "naiu"
    dictunits.Add "08", "ainaiu"
    dictunits.Add "09", "aaayou"
    dictunits.Add "10", "aanyou"
    dictunits.Add "11", "iaeiiaaoaou"
    dictunits.Add "12", "aaaiaaoaou"
    dictunits.Add "13", "o?eiaaoaou"
    dictunits.Add "14", "?aou?iaaoaou"
    dictunits.Add "15", "iyouiaaoaou"
    dictunits.Add "16", "oanouiaaoaou"
    dictunits.Add "17", "naiiaaoaou"
    dictunits.Add "18", "ainaiiaaoaou"
    dictunits.Add "19", "aaayoiaaoaou"
    
    Set dictunitsfem = New Scripting.dictionary
    dictunitsfem.Add "00", ""
    dictunitsfem.Add "1", "iaia"
    dictunitsfem.Add "2", "aaa"
    dictunitsfem.Add "3", "o?e"
    dictunitsfem.Add "4", "?aou?a"
    dictunitsfem.Add "5", "iyou"
    dictunitsfem.Add "6", "oanou"
    dictunitsfem.Add "7", "naiu"
    dictunitsfem.Add "8", "ainaiu"
    dictunitsfem.Add "9", "aaayou"
    dictunitsfem.Add "01", "iaia"
    dictunitsfem.Add "02", "aaa"
    dictunitsfem.Add "03", "o?e"
    dictunitsfem.Add "04", "?aou?a"
    dictunitsfem.Add "05", "iyou"
    dictunitsfem.Add "06", "oanou"
    dictunitsfem.Add "07", "naiu"
    dictunitsfem.Add "08", "ainaiu"
    dictunitsfem.Add "09", "aaayou"
    dictunitsfem.Add "10", "aanyou"
    dictunitsfem.Add "11", "iaeiiaaoaou"
    dictunitsfem.Add "12", "aaaiaaoaou"
    dictunitsfem.Add "13", "o?eiaaoaou"
    dictunitsfem.Add "14", "?aou?iaaoaou"
    dictunitsfem.Add "15", "iyouiaaoaou"
    dictunitsfem.Add "16", "oanouiaaoaou"
    dictunitsfem.Add "17", "naiiaaoaou"
    dictunitsfem.Add "18", "ainaiiaaoaou"
    dictunitsfem.Add "19", "aaayoiaaoaou"
End Function
       

Public Function cents(input_ As String)
    input__ = Replace(input_, ".", "0")
    If StrComp(input_, "000", vbTextCompare) = 0 Then
        cents = "00 eiiaae"
        Exit Function
    End If
    
    If Not StrComp(Mid(input__, 2, 1), "1", vbTextCompare) = 0 Then
        If StrComp(Mid(input__, 3, 1), "1", vbTextCompare) = 0 Then
            cents = subnumbfem(input__) + " " + "eiiaeea"
            Exit Function
        If StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0 Or StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0 Or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0 Then
            cents = subnumbfem(input__) + " " + "eiiaeee"
            Exit Function
    cents = subnumbfem(input__) + " " + "eiiaae"
    Exit Function
End Function

Public Function subnumbfem(input_ As String)
    Dim result As String
    Dim suffix As String
    result = ""
    suffix = ""
    If StrComp(input_, "000", vbTextCompare) = 0 Then
        subnumbfem = result
        Exit Function
    End If
    If Not StrComp(Mid(input_, 1, 1), "0", vbTextCompare) = 0 Then
        result = result & dicthundreds(Mid(input_, 1, 1))
    End If
    If Not StrComp(Mid(input_, 2, 1), "0", vbTextCompare) = 0 And Not StrComp(Mid(input_, 2, 1), "1", vbTextCompare) = 0 Then
        result = result & dictdecimals(Mid(input_, 2, 1))
        result = result & dictunitsfem(Mid(input_, 3, 1))
        subnumbfem = result
        Exit Function
    End If
    suffix = Mid(input_, 2, 1) & Mid(input_, 3, 1)
    result = result & " " + dictunitsfem(suffix)
    subnumbfem = result
End Function

Public Function conjuntion(prefix As String, groups As Integer)
    If groups = 2 Then
        If Not StrComp(Mid(prefix, 2, 1), "1", vbTextCompare) = 0 Then
        If StrComp(Mid(prefix, 3, 1), "1", vbTextCompare) = 0 Then
            conjuntion = subnumbfem(prefix) & " " & dict(groups) & "a"
            Exit Function
        End If
        If StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 Or StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 Or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then
            conjuntion = subnumbfem(prefix) & " " & dict(groups) & "e"
            Exit Function
        End If
        conjuntion = subnumbfem(prefix) & " " & dict(groups)
        Exit Function
    End If
    
    If groups = 1 Then
        conjuntion = subnumbfem(prefix) & " " & dict(groups) & "a"
        Exit Function
    End If
        
    If groups = 3 Or groups = 4 Or groups = 5 Then
        If Not StrComp(Mid(prefix, 2, 1), "1", vbTextCompare) = 0 Then
            If StrComp(Mid(prefix, 3, 1), "1", vbTextCompare) = 0 Then
                conjuntion = subnumbfem(prefix) & " " & dict(groups)
                Exit Function
            End If
        
            If StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 Or StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 Or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then
                conjuntion = subnumbfem(prefix) & " " & dict(groups) & "a"
                Exit Function
            End If
        conjuntion = subnumbfem(prefix) & " " & dict(groups) & "ia"
        Exit Function
        End If
    End If
    
End Function


Public Function numbers(input_ As String)
    Dim qinput As String
    Dim groups As Integer
    Dim result As String
    Dim iter As Integer
    qinput = prepared(input_)
    
    groups = Len(qinput) \ 3
    result = ""
    iter = 0
    MsgBox qinput
    MsgBox groups
    Dim prefix As String
    prefix = ""
    Do While (groups <> 0):
        prefix = Mid(prefix, iter + 1, iter + 3 + 1)
        iter = iter + 3
        If StrComp(subnumb(prefix), "", vbTextCompare) = 0 Then
            GoTo a
        End If
        result = result & conjuntion(prefix, groups)
        groups = groups - 1
a:
    Loop
    input__ = Mid(qinput, Len(qinput) - 3, 3)
    MsgBox input__
    If Not StrComp(Mid(input__, 2, 1), "1", vbTextCompare) = 0 Then
        If StrComp(Mid(input__, 3, 1), "1", vbTextCompare) = 0 Then
            result = result & " ?oaeu"
            numbers = result
            Exit Function
        End If
        If StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0 Or StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0 Or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0 Then
            result = result & " ?oaey"
            numbers = result
            Exit Function
        End If
        result = result & " ?oaeae"
        numbers = result
        Exit Function
    Else
        result = result & " ?oaeae"
        numbers = result
        Exit Function
    MsgBox result
    result = result & " ?oaeae"
    numbers = result
    Exit Function
    End If
End Function

Public Function subnumb(input_ As String)
    fuck
    Dim result As String
    Dim suffix As String
    result = ""
    suffix = ""
    If StrComp(input_, "000", vbTextCompare) = 0 Then
        subnumb = result
        Exit Function
    End If
    If Not StrComp(Mid(input_, 1, 1), "0", vbTextCompare) = 0 Then
        result = result & dicthundreds(Mid(input_, 1, 1))
    End If
    If Not StrComp(Mid(input_, 2, 1), "0", vbTextCompare) = 0 And Not StrComp(Mid(input_, 2, 1), "1", vbTextCompare) = 0 Then
        result = result & dictdecimals(Mid(input_, 2, 1))
        result = result & dictunits(Mid(input_, 3, 1))
        subnumb = result
        Exit Function
    End If
    suffix = Mid(input_, 2, 1) & Mid(input_, 3, 1)
    result = result & " " + dictunits(suffix)
    MsgBox "suffix" & suffix
    subnumb = result
End Function


Public Function numtoRussian(input_ As String)
   fuck
   Dim delim As Integer
   delim = -2
   delim = InStr(0, input_, ".", vbTextCompare)
   If delim = -1 Then
       numtoRussian = numbers(input_) & cents(".00")
       Exit Function
   Else
       numtoRussian = numbers(Mid(input_, 1, delim)) & cents(Mid(input_, delim, delim + 3))
       Exit Function
   End If
End Function

Public Function prepared(input_ As String)
    Dim rest As Integer
    Dim result As String
    rest = 3 - (Len(input_) Mod 3)
    If (rest = 3) Then
        prepared = input_
    End If
    result = ""
    Do While Not rest = 0
        rest = rest - 1
        result = result & "0"
    Loop
    result = result & input_
    prepared = result
End Function
