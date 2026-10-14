    <%
	
		    finalRowSpan = 0
			
			if not rsHourProduction.EOF then 
				while not rsHourProduction.EOF
					finalRowSpan = finalRowSpan + 1
					rsHourProduction.MoveNext
				Wend
			    rsHourProduction.MoveFirst
			end if
			
	
	%>
