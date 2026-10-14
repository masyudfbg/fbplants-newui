    <%
	
		    finalRowSpan = 0
			
			if not rsOutput.EOF then 
				while not rsOutput.EOF
					finalRowSpan = finalRowSpan + 1
					rsOutput.MoveNext
				Wend
				rsOutput.MoveFirst
			end if
			
	
	%>
