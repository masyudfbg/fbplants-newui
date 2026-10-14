    <%
	

			if not rsNettoFinal.EOF then 
				while not rsNettoFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsNettoFinal.MoveNext
				Wend
			    rsNettoFinal.MoveFirst
			end if			
			
	
	%>
