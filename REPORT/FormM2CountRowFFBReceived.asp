    <%
	
		    finalRowSpan = 0
			
			if not rsTerimaTBS.EOF then 
				while not rsTerimaTBS.EOF
					finalRowSpan = finalRowSpan + 1
					rsTerimaTBS.MoveNext
				Wend
			    rsTerimaTBS.MoveFirst
			end if
			
	
	%>
