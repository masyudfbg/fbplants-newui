    <%
	
		    finalRowSpan = 0
			
			'if not rsKernel_FFAQuality.EOF then 
			'	while not rsKernel_FFAQuality.EOF
			'		finalRowSpan = finalRowSpan + 1
			'		rsKernel_FFAQuality.MoveNext
			'	Wend
			'    rsKernel_FFAQuality.MoveFirst
			'end if
			
			if not rsKernel_MoistureQuality.EOF then 
				while not rsKernel_MoistureQuality.EOF
					finalRowSpan = finalRowSpan + 1
					rsKernel_MoistureQuality.MoveNext
				Wend
			    rsKernel_MoistureQuality.MoveFirst
			end if		

			if not rsKernel_DirtQuality.EOF then 
				while not rsKernel_DirtQuality.EOF
					finalRowSpan = finalRowSpan + 1
					rsKernel_DirtQuality.MoveNext
				Wend
			    rsKernel_DirtQuality.MoveFirst
			end if	
			
		

			
	%>
