    <%
	
		    finalRowSpan = 0
			
			if not rsOilKernelProduction.EOF then 
				while not rsOilKernelProduction.EOF
					finalRowSpan = finalRowSpan + 1
					rsOilKernelProduction.MoveNext
				Wend
			    rsOilKernelProduction.MoveFirst
			end if
			

			
	%>
