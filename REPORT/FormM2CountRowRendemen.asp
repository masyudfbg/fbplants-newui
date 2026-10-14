    <%
	
		    finalRowSpan = 0
			
			if not rsRendemenFinal.EOF then 
				while not rsRendemenFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRendemenFinal.MoveNext
				Wend
			    rsRendemenFinal.MoveFirst
			end if
			
			if not rsRendemenFinal.EOF then 
				while not rsRendemenFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRendemenFinal.MoveNext
				Wend
			    rsRendemenFinal.MoveFirst
			end if			

			
	%>
