    <%
	
		    finalRowSpan = 0
			
			if not rsFertilizer.EOF then 
				while not rsFertilizer.EOF
					finalRowSpan = finalRowSpan + 1
					rsFertilizer.MoveNext
				Wend
				rsFertilizer.MoveFirst
			end if
			
	
	%>
