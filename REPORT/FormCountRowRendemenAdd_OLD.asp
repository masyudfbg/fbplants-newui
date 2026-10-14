    <%
	
		    finalRowSpan = 0
			
			if not rsRendemenFinalAdd.EOF then 
				while not rsRendemenFinalAdd.EOF
					finalRowSpan = finalRowSpan + 1
					rsRendemenFinalAdd.MoveNext
				Wend
			    rsRendemenFinalAdd.MoveFirst
			end if
			
			if not rsRendemenFinalAdd.EOF then 
				while not rsRendemenFinalAdd.EOF
					finalRowSpan = finalRowSpan + 1
					rsRendemenFinalAdd.MoveNext
				Wend
			    rsRendemenFinalAdd.MoveFirst
			end if			

			
	%>
