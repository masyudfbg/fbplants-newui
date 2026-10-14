    <%
	
		    finalRowSpan = 0
			
			'if not rsKernel_FFAQuality.EOF then 
			'	while not rsKernel_FFAQuality.EOF
			'		finalRowSpan = finalRowSpan + 1
			'		rsKernel_FFAQuality.MoveNext
			'	Wend
			'    rsKernel_FFAQuality.MoveFirst
			'end if
			
			if not rsCpo_OilLosses.EOF then 
				while not rsCpo_OilLosses.EOF
					finalRowSpan = finalRowSpan + 1
					rsCpo_OilLosses.MoveNext
				Wend
			    rsCpo_OilLosses.MoveFirst
			end if		
            'response.write  "No : " & finalRowSpan
			if not rsKernel_KernelLosses.EOF then 
				while not rsKernel_KernelLosses.EOF
					finalRowSpan = finalRowSpan + 1
					rsKernel_KernelLosses.MoveNext
				Wend
			    rsKernel_KernelLosses.MoveFirst
			end if	
			'response.write  "No : " &  finalRowSpan
		

			
	%>
