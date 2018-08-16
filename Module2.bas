Attribute VB_Name = "Module1"
    Dim dict  As New Scripting.dictionary
    Dim dicthundreds As New Scripting.dictionary
    Dim dictdecimals As New Scripting.dictionary
    Dim dictunits As New Scripting.dictionary
    Dim dictunitsfem As New Scripting.dictionary

Public Function fuck()
    
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
       

Public cents(input_ as String)
	input__ = replace(input_, ".", "0")
	If StrComp(input_, "000", vbTextCompare) = 0 Then
		cents="00 копеек"
		Exit Function
	End if
	
	if Not StrComp(Mid(input__, 2, 1), "1", vbTextCompare) = 0 Then       
        if StrComp(Mid(input__, 3, 1), "1", vbTextCompare) = 0  Then
            cents =  subnumbfem(input__) + " " +  "копейка"
			Exit Function
        if 	StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0   or	StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0   or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0
            cents =  subnumbfem(input__) + " " +  "копейки"	
			Exit Function
    cents= subnumbfem(input__) + " " + "копеек"
	Exit Function
end function

Public subnumbfem(input_ as String)
	Dim result as String 
	Dim suffix as String
	result = ""
	suffix = ""
	If StrComp(input_, "000", vbTextCompare) = 0 Then
		subnumbfem=result
		Exit Function
	End if
	if Not StrComp(Mid(input_, 1, 1), "0", vbTextCompare) = 0 Then    
		result = result & dicthundreds(Mid(input_, 1, 1))
	End if
    if Not StrComp(Mid(input_, 2, 1), "0", vbTextCompare) = 0 and    Not StrComp(Mid(input_, 2, 1), "1", vbTextCompare) = 0  Then  
		result = result & dictdecimals(Mid(input_, 2, 1))
		result = result & dictunitsfem(Mid(input_, 3, 1))
		subnumbfem=result
		Exit Function
	End if
	suffix = Mid(input_, 2, 1)  & Mid(input_, 3, 1)
	result = result & " " + dictunitsfem(suffix)
	subnumbfem=result
end function

Public function conjuntion(prefix as String , groups as Integer)
	if groups = 2 then
		if Not StrComp(Mid(prefix, 2, 1), "1", vbTextCompare) = 0 Then 
		if StrComp(Mid(prefix, 3, 1), "1", vbTextCompare) = 0 Then
			conjuntion=subnumbfem(prefix) & " " & dict(groups) & "а"
			Exit Function
		End if
		if  StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 or  StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then 
			conjuntion=subnumbfem(prefix) & " " & dict(groups) & "и"
			Exit Function
		end if
		conjuntion=subnumbfem(prefix) & " " & dict(groups) 
		Exit Function
	end if 
	
	if groups = 1 then
		conjuntion=subnumbfem(prefix) & " " & dict(groups) & "а"
		Exit Function
	end if
		
	if groups = 3 or 	groups = 4 or groups = 5 then
		if Not StrComp(Mid(prefix, 2, 1), "1", vbTextCompare) = 0 Then 
			if StrComp(Mid(prefix, 3, 1), "1", vbTextCompare) = 0 Then
				conjuntion=subnumbfem(prefix) & " " & dict(groups) 
				Exit Function
			End if		
		
			if  StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 or  StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then 
				conjuntion=subnumbfem(prefix) & " " & dict(groups) & "а"
				Exit Function
			End if
		conjuntion=subnumbfem(prefix) & " " & dict(groups)  & "ов"
		Exit Function
		end if 
	end if 
	
end fucntion


public function numbers(input_ as String)
    qinput  = prepared(input_)
    groups = len(qinput) \ 3
    result = ""
    iter=0
	dim prefix as String
	prefix = ""
    do while (groups<>0):
		prefix= Mid(prefix, iter+1, iter+3+1)      
        iter = iter + 3
        if StrComp(subnumb(prefix), "", vbTextCompare) = 0  then
            continue
		end if 
        result = result & conjuntion(prefix, groups)
        groups = groups-1
	Loop
    input__ =  Mid(qinput, len(qinput)-3, 3) 
    print(input__)
    if Not StrComp(Mid(input__, 2, 1), "1", vbTextCompare) = 0 Then 
        if StrComp(Mid(input__, 3, 1), "1", vbTextCompare) = 0 Then
            print("flow2")
            result = result & " рубль"
            numbers= result
			Exit Function
		end if 
        if StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0 or StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0 then
            print("flow3")
            result = result & " рубля"
            numbers= result
			Exit Function
        end if     
        result = result & " рублей"
        numbers= result
		Exit Function
    else        
        result = result & " рублей"
        numbers= result
		Exit Function
    print(result)
    result = result & рублей"
    numbers= result
	Exit Function
	end if
end function
Public subnumb(input_ as String)
	Dim result as String 
	Dim suffix as String
	result = ""
	suffix = ""
	If StrComp(input_, "000", vbTextCompare) = 0 Then
		subnumb=result
		Exit Function
	End if
	if Not StrComp(Mid(input_, 1, 1), "0", vbTextCompare) = 0 Then    
		result = result & dicthundreds(Mid(input_, 1, 1))
	End if
    if Not StrComp(Mid(input_, 2, 1), "0", vbTextCompare) = 0 and    Not StrComp(Mid(input_, 2, 1), "1", vbTextCompare) = 0  Then  
		result = result & dictdecimals(Mid(input_, 2, 1))
		result = result & dictunits(Mid(input_, 3, 1))
		subnumb=result
		Exit Function
	End if
	suffix = Mid(input_, 2, 1)  & Mid(input_, 3, 1)
	result = result & " " + dictunits(suffix)
	subnumb=result
end function


Public Function numtoRussian(input_ as String)
   fuck
   Dim delim As integer
   delim=-2
   delim = Instr(0, input_, '.', vbTextCompare) 
   if delim = -1 Then
       numtoRussian= numbers(input)&cents(".00")
   else
       numtoRussian= numbers(MID(input_; 1;delim)) &cents(MID(input_; delim;delim+3))
   End if
End Function

Public Function prepared(input_ As String)
    Dim rest As Integer
    Dim result As String
    rest = 3 - (Len(input_) Mod 3)
    If (rest = 3) Then
        prepared = input_
    result = ""
    Do While Not rest = 0
        rest = rest - 1
        result = result & "0"
    result = result & input_
    prepared = result
End Function




