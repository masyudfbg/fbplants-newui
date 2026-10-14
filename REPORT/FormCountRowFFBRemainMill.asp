    <%
	
		    finalRowSpan = 0
			
			if not rsSisaTBS.EOF then 
				while not rsSisaTBS.EOF
					finalRowSpan = finalRowSpan + 1
					rsSisaTBS.MoveNext
				Wend
			    rsSisaTBS.MoveFirst
			end if
			
			if not rsClosingTBS.EOF then 
				while not rsClosingTBS.EOF
					finalRowSpan = finalRowSpan + 1
					rsClosingTBS.MoveNext
				Wend
			    rsClosingTBS.MoveFirst
			end if
			
	%>
