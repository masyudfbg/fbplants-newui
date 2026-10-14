    <%
	
		    finalRowSpan = 0
			
			if not rsMandays.EOF then 
				while not rsMandays.EOF
					finalRowSpan = finalRowSpan + 1
					rsMandays.MoveNext
				Wend
				rsMandays.MoveFirst
			end if
			
	
	%>
