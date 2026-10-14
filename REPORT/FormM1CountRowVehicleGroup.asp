    <%
	
		    finalRowSpan = 0
			
			if not rsVehicleGroup.EOF then 
				while not rsVehicleGroup.EOF
					finalRowSpan = finalRowSpan + 1
					rsVehicleGroup.MoveNext
				Wend
				rsVehicleGroup.MoveFirst
			end if
			
	
	%>
