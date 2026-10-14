    <%
	
			
			if not rsSortasiFinal.EOF then 
				while not rsSortasiFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsSortasiFinal.MoveNext
				Wend
			    rsSortasiFinal.MoveFirst
			end if
			
	
	%>
