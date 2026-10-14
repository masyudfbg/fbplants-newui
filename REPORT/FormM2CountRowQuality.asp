    <%
	
		    finalRowSpan = 0
			
			if not rsCPO_FFAQuality.EOF then 
				while not rsCPO_FFAQuality.EOF
					finalRowSpan = finalRowSpan + 1
					rsCPO_FFAQuality.MoveNext
				Wend
			    rsCPO_FFAQuality.MoveFirst
			end if
			
			if not rsCPO_MoistureQuality.EOF then 
				while not rsCPO_MoistureQuality.EOF
					finalRowSpan = finalRowSpan + 1
					rsCPO_MoistureQuality.MoveNext
				Wend
			    rsCPO_MoistureQuality.MoveFirst
			end if		

			if not rsCPO_DirtQuality.EOF then 
				while not rsCPO_DirtQuality.EOF
					finalRowSpan = finalRowSpan + 1
					rsCPO_DirtQuality.MoveNext
				Wend
			    rsCPO_DirtQuality.MoveFirst
			end if	
			
		

			
	%>
