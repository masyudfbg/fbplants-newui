    <%
	
		    finalRowSpan = 0
			
			if not rsshellstock.EOF then 
				while not rsshellstock.EOF
					finalRowSpan = finalRowSpan + 1
					rsshellstock.MoveNext
				Wend
			    rsshellstock.MoveFirst
			end if
			
		

			
	%>
