    <%
	
		    finalRowSpan = 0
			
			if not rsProses.EOF then 
				while not rsProses.EOF
					finalRowSpan = finalRowSpan + 1
					rsProses.MoveNext
				Wend
			    rsProses.MoveFirst
			end if
			
	
	%>
