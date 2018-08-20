Attribute VB_Name = "Module1"
Dim dict  As New Scripting.dictionary
Dim dicthundreds As New Scripting.dictionary
Dim dictdecimals As New Scripting.dictionary
Dim dictunits As New Scripting.dictionary
Dim dictunitsfem As New Scripting.dictionary

Public Function str_getchar(input_ As String, indexstart As Integer)
    str_getchar = Mid(input_, indexstart + 1, 1)
End Function


Public Function test(input_ As String)
    Dim delim As Integer
    delim = index(input_, ",")
    If delim = -1 Then
        delim = index(input_, ".")
    End If
    test = delim
   
    Exit Function
    
    delim = 0
    MsgBox index(input_, ",")
    Dim str1 As String
    Dim str2 As String
    str1 = str_cut(input_, delim, delim + 2)
    str2 = str_cut(input_, delim, delim + 2)
    
    MsgBox equal(str1, str2)
    MsgBox str1
    MsgBox str2


End Function

Public Function str_cut(input_ As String, from As Integer, finish As Integer)
    Dim delta As Integer
    Dim from_ As Integer
    Dim finish_ As Integer
    from_ = from + 1
    
    str_cut = Mid(input_, from_, finish - from_ + 1)
End Function

Public Function equal(input1 As String, input2 As String)
    If StrComp(input1, input2, vbTextCompare) = 0 Then
        equal = True
            Exit Function
    End If
    equal = False
    Exit Function
End Function

Public Function index(input_ As String, found As String)
    index = InStr(1, input_, found) - 1
End Function


Public Function preps()
    
    Set dict = New Scripting.dictionary
    dict.Add 0, ""
    dict.Add 1, ""
    dict.Add 2, "тысяч"
    dict.Add 3, "миллион"
    dict.Add 4, "миллиард"
    dict.Add 5, "триллион"

    Set dicthundreds = New Scripting.dictionary
    dicthundreds.Add "1", "сто"
    dicthundreds.Add "2", "двести"
    dicthundreds.Add "3", "триста"
    dicthundreds.Add "4", "четыреста"
    dicthundreds.Add "5", "пятьсот"
    dicthundreds.Add "6", "шестьсот"
    dicthundreds.Add "7", "семьсот"
    dicthundreds.Add "8", "восемьсот"
    dicthundreds.Add "9", "девятьсот"

    Set dictdecimals = New Scripting.dictionary
    dictdecimals.Add "2", "двадцать"
    dictdecimals.Add "3", "тридцать"
    dictdecimals.Add "4", "сорок"
    dictdecimals.Add "5", "пятьдесят"
    dictdecimals.Add "6", "шестьдесят"
    dictdecimals.Add "7", "семьдесят"
    dictdecimals.Add "8", "восемьдесят"
    dictdecimals.Add "9", "девяносто"


    Set dictunits = New Scripting.dictionary
    dictunits.Add "00", ""
    dictunits.Add "1", "один"
    dictunits.Add "2", "два"
    dictunits.Add "3", "три"
    dictunits.Add "4", "четыре"
    dictunits.Add "5", "пять"
    dictunits.Add "6", "шесть"
    dictunits.Add "7", "семь"
    dictunits.Add "8", "восемь"
    dictunits.Add "9", "девять"
    dictunits.Add "01", "один"
    dictunits.Add "02", "два"
    dictunits.Add "03", "три"
    dictunits.Add "04", "четыре"
    dictunits.Add "05", "пять"
    dictunits.Add "06", "шесть"
    dictunits.Add "07", "семь"
    dictunits.Add "08", "восемь"
    dictunits.Add "09", "девять"
    dictunits.Add "10", "десять"
    dictunits.Add "11", "одиннадцать"
    dictunits.Add "12", "двенадцать"
    dictunits.Add "13", "тринадцать"
    dictunits.Add "14", "четырнадцать"
    dictunits.Add "15", "пятьнадцать"
    dictunits.Add "16", "шестьнадцать"
    dictunits.Add "17", "семнадцать"
    dictunits.Add "18", "восемнадцать"
    dictunits.Add "19", "девятнадцать"
    
    Set dictunitsfem = New Scripting.dictionary
    dictunitsfem.Add "00", ""
    dictunitsfem.Add "1", "одна"
    dictunitsfem.Add "2", "две"
    dictunitsfem.Add "3", "три"
    dictunitsfem.Add "4", "четыре"
    dictunitsfem.Add "5", "пять"
    dictunitsfem.Add "6", "шесть"
    dictunitsfem.Add "7", "семь"
    dictunitsfem.Add "8", "восемь"
    dictunitsfem.Add "9", "девять"
    dictunitsfem.Add "01", "одна"
    dictunitsfem.Add "02", "две"
    dictunitsfem.Add "03", "три"
    dictunitsfem.Add "04", "четыре"
    dictunitsfem.Add "05", "пять"
    dictunitsfem.Add "06", "шесть"
    dictunitsfem.Add "07", "семь"
    dictunitsfem.Add "08", "восемь"
    dictunitsfem.Add "09", "девять"
    dictunitsfem.Add "10", "десять"
    dictunitsfem.Add "11", "одиннадцать"
    dictunitsfem.Add "12", "двенадцать"
    dictunitsfem.Add "13", "тринадцать"
    dictunitsfem.Add "14", "четырнадцать"
    dictunitsfem.Add "15", "пятьнадцать"
    dictunitsfem.Add "16", "шестьнадцать"
    dictunitsfem.Add "17", "семнадцать"
    dictunitsfem.Add "18", "восемнадцать"
    dictunitsfem.Add "19", "девятнадцать"
End Function
       

Public Function cents(input___ As String)
    Dim input_ As String
    input_ = Replace(input___, ",", "0")
    value = str_cut(input_, 1, 3)
    If equal(input_, "000") Then
        cents = " 00 копеек"
        Exit Function
    End If
    
    If Not equal(str_getchar(input_, 1), "1") Then
        If equal(str_getchar(input_, 2), "1") Then
            cents = " " & value & " " & "копейка"
            Exit Function
            End If
        If equal(str_getchar(input_, 2), "2") Or equal(str_getchar(input_, 2), "3") Or equal(str_getchar(input_, 2), "4") Then
            cents = " " & value & " " & "копейки"
            Exit Function
        End If
    cents = " " & value & " " & "копеек"
    Exit Function
    End If
End Function

Public Function subnumbfem(input_ As String)
    Dim result As String
    Dim suffix As String
    result = ""
    suffix = ""
    If equal("000", input_) Then
        subnumbfem = result
        Exit Function
    End If
    If Not equal(str_getchar(input_, 0), "0") Then
        result = result & " " & dicthundreds(str_getchar(input_, 0))
    End If
    If (Not equal(str_getchar(input_, 1), "0") And Not equal(str_getchar(input_, 1), "1")) Then
        result = result & " " & dictdecimals(str_getchar(input_, 1))
        result = result & " " & dictunitsfem(str_getchar(input_, 2))
        subnumbfem = result
        Exit Function
    End If
    suffix = str_getchar(input_, 1) & str_getchar(input_, 2)
    result = result & " " & dictunitsfem(suffix)
    subnumbfem = result
End Function

Public Function conjuntion(prefix As String, groups As Integer)
    If groups = 2 Then
    
        If Not equal(str_getchar(prefix, 1), "1") Then
            If equal(str_getchar(prefix, 2), "1") Then
                conjuntion = subnumbfem(prefix) & " " & dict(groups) & "а"
            Exit Function
            End If
            If equal(str_getchar(prefix, 2), "2") Or equal(str_getchar(prefix, 2), "3") Or equal(str_getchar(prefix, 2), "4") Then
                conjuntion = subnumbfem(prefix) & " " & dict(groups) & "а"
            Exit Function
            End If
        End If
        conjuntion = subnumbfem(prefix) & " " & dict(groups)
        Exit Function
    End If
    
    If groups = 1 Then
        conjuntion = subnumb(prefix)
        Exit Function
    End If
        
    If groups = 3 Or groups = 4 Or groups = 5 Then
        If Not equal(str_getchar(prefix, 1), "1") Then
            If equal(str_getchar(prefix, 2), "1") Then
                conjuntion = subnumb(prefix) & " " & dict(groups) & ""
                Exit Function
            End If
        
            If equal(str_getchar(prefix, 2), "2") Or equal(str_getchar(prefix, 2), "3") Or equal(str_getchar(prefix, 2), "4") Then
                
                conjuntion = subnumb(prefix) & " " & dict(groups) & "а"
                Exit Function
            End If
       
        conjuntion = subnumb(prefix) & " " & dict(groups) & "ов"
        
        Exit Function
        
        Else
        
        conjuntion = subnumb(prefix) & " " & dict(groups) & "ов"
        End If
    End If
    
End Function


Public Function numbers(input_ As String)
    preps
    Dim qinput As String
    Dim groups As Integer
    Dim result As String
    Dim iter As Integer
    Dim input__ As String
    qinput = prepared(input_)
    groups = Len(qinput) \ 3
    result = ""
    iter = 0
    Dim prefix As String
    Do While (groups <> 0):
        Do
            
            prefix = str_cut(qinput, iter, iter + 3)
            iter = iter + 3
            
            If equal(subnumb(prefix), "") Then
                groups = groups - 1
                Exit Do
            End If
            result = result & conjuntion(prefix, groups)
            groups = groups - 1
            Exit Do
        Loop
    Loop
    input__ = str_cut(qinput, Len(qinput) - 3, Len(qinput))
    If Not equal(str_getchar(input__, 1), "1") Then
        If equal(str_getchar(input__, 2), "1") Then
            result = result & " рубль"
            numbers = result
            Exit Function
        End If
        If equal(str_getchar(input__, 2), "2") Or equal(str_getchar(input__, 2), "3") Or equal(str_getchar(input__, 2), "4") Then
            result = result & " рубля"
            numbers = result
            Exit Function
        End If
        result = result & " рублей"
        numbers = result
        Exit Function
    Else
        result = result & " рублей"
        numbers = result
        Exit Function
    result = result & " рублей"
    numbers = result
    Exit Function
    End If
End Function

Public Function subnumb(input_ As String)
    Dim result As String
    Dim suffix As String
    result = ""
    suffix = ""
    If equal("000", input_) Then
        subnumb = result
        Exit Function
    End If
    If Not equal(str_getchar(input_, 0), "0") Then
        result = result & " " & dicthundreds(str_getchar(input_, 0))
    End If
    If (Not equal(str_getchar(input_, 1), "0") And Not equal(str_getchar(input_, 1), "1")) Then
        result = result & " " & dictdecimals(str_getchar(input_, 1))
        result = result & " " & dictunits(str_getchar(input_, 2))
        subnumb = result
        Exit Function
    End If
    suffix = str_getchar(input_, 1) & str_getchar(input_, 2)
    result = result & " " & dictunits(suffix)
    subnumb = result
End Function


Public Function numtoRussian(input_ As String)
   preps

   Dim delim As Integer
   delim = -2
   delim = index(input_, ",")
   If delim = -1 Then
     delim = index(input_, ".")
   End If
   If delim = -1 Then
       numtoRussian = numbers(input_) & cents(",00")
       Exit Function
   Else
       Dim numbers_ As String
       Dim cents_ As String
       numbers_ = str_cut(input_, 0, delim)
       cents_ = str_cut(input_, delim, delim + 3)
       numtoRussian = numbers(numbers_) & cents(cents_)
       Exit Function
   End If
End Function

Public Function prepared(input_ As String)
    Dim rest As Integer
    Dim result As String
    rest = 3 - (Len(input_) Mod 3)
    If (rest = 3) Then
        prepared = input_
        Exit Function
    End If
    result = ""
    Do While Not rest = 0
        rest = rest - 1
        result = result & "0"
    Loop
    result = result & input_
    prepared = result
End Function
