    <%
	
		    finalRowSpan = 0
	
			if not rsLandClearing.EOF then 
				while not rsLandClearing.EOF
					finalRowSpan = finalRowSpan + 1
					rsLandClearing.MoveNext
				Wend
				rsLandClearing.MoveFirst
			end if
			
			
	
	%>
