    <%
	
			
			if not rsHourCapacityFinal.EOF then 
				while not rsHourCapacityFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsHourCapacityFinal.MoveNext
				Wend
			    rsHourCapacityFinal.MoveFirst
			end if
			
	
	%>
