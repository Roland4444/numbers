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
    dict.Add 2, "�����"
    dict.Add 3, "�������"
    dict.Add 4, "��������"
    dict.Add 5, "��������"

    Set dicthundreds = New Scripting.dictionary
    dicthundreds.Add "1", "���"
    dicthundreds.Add "2", "������"
    dicthundreds.Add "3", "������"
    dicthundreds.Add "4", "���������"
    dicthundreds.Add "5", "�������"
    dicthundreds.Add "6", "��������"
    dicthundreds.Add "7", "�������"
    dicthundreds.Add "8", "���������"
    dicthundreds.Add "9", "���������"

    Set dictdecimals = New Scripting.dictionary
    dictdecimals.Add "2", "��������"
    dictdecimals.Add "3", "��������"
    dictdecimals.Add "4", "�����"
    dictdecimals.Add "5", "���������"
    dictdecimals.Add "6", "����������"
    dictdecimals.Add "7", "���������"
    dictdecimals.Add "8", "�����������"
    dictdecimals.Add "9", "���������"


    Set dictunits = New Scripting.dictionary
    dictunits.Add "00", ""
    dictunits.Add "1", "����"
    dictunits.Add "2", "���"
    dictunits.Add "3", "���"
    dictunits.Add "4", "������"
    dictunits.Add "5", "����"
    dictunits.Add "6", "�����"
    dictunits.Add "7", "����"
    dictunits.Add "8", "������"
    dictunits.Add "9", "������"
    dictunits.Add "01", "����"
    dictunits.Add "02", "���"
    dictunits.Add "03", "���"
    dictunits.Add "04", "������"
    dictunits.Add "05", "����"
    dictunits.Add "06", "�����"
    dictunits.Add "07", "����"
    dictunits.Add "08", "������"
    dictunits.Add "09", "������"
    dictunits.Add "10", "������"
    dictunits.Add "11", "�����������"
    dictunits.Add "12", "����������"
    dictunits.Add "13", "����������"
    dictunits.Add "14", "������������"
    dictunits.Add "15", "�����������"
    dictunits.Add "16", "������������"
    dictunits.Add "17", "����������"
    dictunits.Add "18", "������������"
    dictunits.Add "19", "������������"
	
    Set dictunitsfem = New Scripting.dictionary
    dictunitsfem.Add "00", ""
    dictunitsfem.Add "1", "����"
    dictunitsfem.Add "2", "���"
    dictunitsfem.Add "3", "���"
    dictunitsfem.Add "4", "������"
    dictunitsfem.Add "5", "����"
    dictunitsfem.Add "6", "�����"
    dictunitsfem.Add "7", "����"
    dictunitsfem.Add "8", "������"
    dictunitsfem.Add "9", "������"
    dictunitsfem.Add "01", "����"
    dictunitsfem.Add "02", "���"
    dictunitsfem.Add "03", "���"
    dictunitsfem.Add "04", "������"
    dictunitsfem.Add "05", "����"
    dictunitsfem.Add "06", "�����"
    dictunitsfem.Add "07", "����"
    dictunitsfem.Add "08", "������"
    dictunitsfem.Add "09", "������"
    dictunitsfem.Add "10", "������"
    dictunitsfem.Add "11", "�����������"
    dictunitsfem.Add "12", "����������"
    dictunitsfem.Add "13", "����������"
    dictunitsfem.Add "14", "������������"
    dictunitsfem.Add "15", "�����������"
    dictunitsfem.Add "16", "������������"
    dictunitsfem.Add "17", "����������"
    dictunitsfem.Add "18", "������������"
    dictunitsfem.Add "19", "������������"
End Function
       

Public cents(input_ as String)
	input__ = replace(input_, ".", "0")
	If StrComp(input_, "000", vbTextCompare) = 0 Then
		cents="00 ������"
		Exit Function
	End if
	
	if Not StrComp(Mid(input__, 2, 1), "1", vbTextCompare) = 0 Then       
        if StrComp(Mid(input__, 3, 1), "1", vbTextCompare) = 0  Then
            cents =  subnumbfem(input__) + " " +  "�������"
			Exit Function
        if 	StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0   or	StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0   or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0
            cents =  subnumbfem(input__) + " " +  "�������"	
			Exit Function
    cents= subnumbfem(input__) + " " + "������"
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
			conjuntion=subnumbfem(prefix) & " " & dict(groups) & "�"
			Exit Function
		End if
		if  StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 or  StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then 
			conjuntion=subnumbfem(prefix) & " " & dict(groups) & "�"
			Exit Function
		end if
		conjuntion=subnumbfem(prefix) & " " & dict(groups) 
		Exit Function
	end if 
	
	if groups = 1 then
		conjuntion=subnumbfem(prefix) & " " & dict(groups) & "�"
		Exit Function
	end if
		
	if groups = 3 or 	groups = 4 or groups = 5 then
		if Not StrComp(Mid(prefix, 2, 1), "1", vbTextCompare) = 0 Then 
			if StrComp(Mid(prefix, 3, 1), "1", vbTextCompare) = 0 Then
				conjuntion=subnumbfem(prefix) & " " & dict(groups) 
				Exit Function
			End if		
		
			if  StrComp(Mid(prefix, 3, 1), "2", vbTextCompare) = 0 or  StrComp(Mid(prefix, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(prefix, 3, 1), "4", vbTextCompare) = 0 Then 
				conjuntion=subnumbfem(prefix) & " " & dict(groups) & "�"
				Exit Function
			End if
		conjuntion=subnumbfem(prefix) & " " & dict(groups)  & "��"
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
            result = result & " �����"
            numbers= result
			Exit Function
		end if 
        if StrComp(Mid(input__, 3, 1), "2", vbTextCompare) = 0 or StrComp(Mid(input__, 3, 1), "3", vbTextCompare) = 0 or StrComp(Mid(input__, 3, 1), "4", vbTextCompare) = 0 then
            print("flow3")
            result = result & " �����"
            numbers= result
			Exit Function
        end if     
        result = result & " ������"
        numbers= result
		Exit Function
    else        
        result = result & " ������"
        numbers= result
		Exit Function
    print(result)
    result = result & ������"
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




