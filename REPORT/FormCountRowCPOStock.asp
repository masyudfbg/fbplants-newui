    <%
	
		    finalRowSpan = 0
			
			if not rsCPOStock.EOF then 
				while not rsCPOStock.EOF
					finalRowSpan = finalRowSpan + 1
					rsCPOStock.MoveNext
				Wend
			    rsCPOStock.MoveFirst
			end if
			
		

			
	%>
