    <%
	
		    finalRowSpan = 0
			
			if not rsRatioFinal.EOF then 
				while not rsRatioFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRatioFinal.MoveNext
				Wend
				rsRatioFinal.MoveFirst
			end if
			
			if v_txtDetailInfo = "Yes" then
				while not rsRatioFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRatioFinal.MoveNext
				Wend
			    rsRatioFinal.MoveFirst	
								
			end if			
			
	
	%>
