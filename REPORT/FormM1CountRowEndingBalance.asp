    <%
	
		    finalRowSpan = 0
			
			if not rsEndingBalance.EOF then 
				while not rsEndingBalance.EOF
					finalRowSpan = finalRowSpan + 1
					rsEndingBalance.MoveNext
				Wend
				rsEndingBalance.MoveFirst
			end if
			
	
	%>
