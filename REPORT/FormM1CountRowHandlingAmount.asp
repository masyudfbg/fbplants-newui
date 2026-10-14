    <%
	
		    finalRowSpan = 0
			
			if not rsRatioAmountFinal.EOF then 
				while not rsRatioAmountFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRatioAmountFinal.MoveNext
				Wend
				rsRatioAmountFinal.MoveFirst
			end if
			
			if v_txtDetailInfo = "Yes" then
				while not rsRatioAmountFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsRatioAmountFinal.MoveNext
				Wend
			    rsRatioAmountFinal.MoveFirst	
								
			end if			
			
	
	%>
