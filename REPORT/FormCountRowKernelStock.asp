    <%
	
		    finalRowSpan = 0
			
			if not rsKernelStock.EOF then 
				while not rsKernelStock.EOF
					finalRowSpan = finalRowSpan + 1
					rsKernelStock.MoveNext
				Wend
			    rsKernelStock.MoveFirst
			end if
			
		

			
	%>
