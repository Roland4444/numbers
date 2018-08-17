public function str_getchar(input_ as String, indexstart as integer )
	str_getchar = Mid(input_, indexstart+1, 1)
end function 


public function str_cut(input_ as String, from as integer, finish as integer )
	dim delta as integer
	from = from + 1
    finish = finish + 1
	if finish = from then
		finish = finish + 1
	end if 
	str_cut = Mid(input_, from, finish - from+1)
end function 
