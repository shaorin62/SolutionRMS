<SCRIPT LANGUAGE="vbscript">
Sub gCreateTextbox(obj, vntData)
	Dim ArrStr, intColCnt, intRowCnt
	Dim strInHtml, THeader, TValue, TId
	
	If isArray(vntData) then
		'콤보 데이터를 만든다.
		intColCnt=-1: intColCnt=ubound(vntData,1)
		intRowCnt=-1: intRowCnt=ubound(vntData,2)
		
		Redim ArrStr(intRowCnt)
			
	    For i=1 To intRowCnt
	     TId		= vntData(0,i)
	     TValue		= vntData(1,i)
	     THeader	= vntData(2,i)  
	       ArrStr(i) = THeader &" <INPUT TYPE='TEXT' VALUE='"& TValue &"' ID='" & TId & "'>  "
	    Next
	    
	    strInHtml = Join(ArrStr)	
	    obj.innerHTML= obj.innerHTML & strInHtml
    End if
 End Sub
  
</SCRIPT>