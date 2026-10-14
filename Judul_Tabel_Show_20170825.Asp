   <%If Request("txttable") ="Bank" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">CONTROL JOB</TD>	
	<TD align="left">ACCOUNT CODE</TD>	
	<TD align="left">ADDRESS</TD>	
	<TD align="left">CITY</TD>	
	<TD align="left">STATE</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">DATE CREATED</TD>	
	<TD align="left">ACTIVATION</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Roles" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">ROLES CODE</TD>	
	<TD align="left">ROLES NAME</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Stock" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">ITEM TYPE</TD>	
	<TD align="left">MANUFACTURE PART NUMBER</TD>	
	<TD align="left">MANUFACTURE CODE</TD>	
	<TD align="left">SUPPLIER NUMBER</TD>	
	<TD align="left">SUPPLIER CODE</TD>	
	<TD align="left">UOM</TD>	
	<TD align="left">ValidUntil</TD>	
	<TD align="left">STOCK GROUP</TD>
	<TD align="left">IS LESS ZERO</TD>
	<td align="left">ACTIVE MATERIAL</td>	
	</tr>
   <%ElseIf Request("txttable") ="StockGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">CONTROL JOB</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Supplier" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">CONTROL JOB</TD>	
	<TD align="left">DATE CREATED</TD>	
	<TD align="left">ACTIVATION</TD>	
	<TD align="left">ADDRESS</TD>	
	<TD align="left">CITY</TD>	
	<TD align="left">STATE</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">BANK</TD>	
	<TD align="left">BANK ACCOUNT</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Store" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">ACTIVATION</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Gang" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">ACTIVATION</TD>	
	<TD align="left">GANGTYPE</TD>	
	</tr>
	<%ElseIf Request("txttable") ="RangeTaxGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">FC CODE</TD>	
	<TD align="left">FC NAME</TD>	
	<TD align="left">RANGE FROM</TD>	
	<TD align="left">RANGE UNTIL</TD>	
	<TD align="left">TARIFF PERCENTAGE</TD>	
	<TD align="left">REMARKS</TD>
	</tr>
	<%ElseIf Request("txttable") ="MaritalStatus" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">FC CODE</TD>	
	<TD align="left">FC NAME</TD>	
	</tr>
	<%ElseIf Request("txttable") ="UntaxableIncome" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">MARITAL STATUS CODE</TD>	
	<TD align="left">FCNAME</TD>	
	<TD align="left">UTI EMPLOYEE</TD>	
	<TD align="left">UTI SPOUSE</TD>	
	<TD align="left">UTI CHILD</TD>	
	<TD align="left">REMARKS</TD>
	</tr>
   <%ElseIf Request("txttable") ="Employee" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">SECTION NAME</TD>	
	<TD align="left">GANG</TD>	
	<TD align="left">DATE BIRTH</TD>	
	<TD align="left">RELIGION</TD>	
	<TD align="left">RACE</TD>	
	<TD align="left">IDENTITY CARD</TD>	
	<TD align="left">HIGHEST EDUCATION</TD>	
	<TD align="left">YEAR OF GRADUATION</TD>	
	<TD align="left">DATE OF JOIN</TD>	
	<TD align="left">DATE OF DESIGNATION</TD>	
	<TD align="left">PRESENT GRADE</TD>	
	<TD align="left">GRADE DETAIL</TD>	
	<TD align="left">PRES RATE</TD>	
	<TD align="left">PREV GRADE</TD>	
	<TD align="left">PREV DETAIL</TD>	
	<TD align="left">PREV RATE</TD>	
	<TD align="left">MARITAL</TD>	
	<TD align="left">SPOUSE</TD>	
	<TD align="left">DEPENDENT CHILD</TD>	
	<TD align="left">JAMSOSTEK</TD>	
	<TD align="left">JAMSOSTEK CATEGORY</TD>	
	<TD align="left">POSITION</TD>		
	<TD align="left">DETERMINATE</TD>
	<TD align="left">NPWP</TD>
	<TD align="left">NATIONALITY</TD>	
	<TD align="left">CALCULATE TAX</TD>
	<TD align="left">TAX STATUS</TD>
    <TD align="left">GENDER</TD>
	<TD align="left">PAYMENT TYPE</TD>
	<TD align="left">ACCOUNTNO</TD>	
	<TD align="left">BANKNAME</TD>	
	<TD align="left">OWNERSHIPOFACCOUNT</TD>	
	<TD align="left">PPH PAID BY COMPANY</TD>
	<TD align="left">ADDRESS</TD>
	<TD align="left">BPJS KESEHATAN CATEGORY</TD>
	<TD align="left">NO BPJS KESEHATAN</TD>
	</tr>
	<%ElseIf Request("txttable") ="EmployeeAllowdedDetail" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">EMPLOYEE CODE</TD>	
	<TD align="left">ALLOWDED CODE</TD>	
	<TD align="left">RATE 1</TD>	
	<TD align="left">RATE 2</TD>	
	<TD align="left">RATE 3</TD>	
	<TD align="left">RATE 4</TD>
	<TD align="left">RATE 5</TD>
	<TD align="left">REMARKS</TD>
	</tr>
   <%ElseIf Request("txttable") ="Workshop" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">ACTIVATION</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Vehicle" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">VEHICLEGROUPCODE</TD>	
	<TD align="left">DATECREATED</TD>	
	<TD align="left">ACTIVATION</TD>
	<TD align="left">DATETERMINATE</TD>	
	<TD align="left">REGISTRATIONNO</TD>	
	<TD align="left">MAKE</TD>	
	<TD align="left">MODEL</TD>	
	<TD align="left">ENGINENO</TD>	
	<TD align="left">CHASISNO</TD>	
	<TD align="left">YEAROFMADE</TD>	
	<TD align="left">YEAROFPURCHASE</TD>	
	<TD align="left">OWNERSHIP</TD>	
	<TD align="left">VEHICLESTATUS</TD>	
	<TD align="left">CONTROLALLOCATION</TD>	
	<TD align="left">CONTROLJOB</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Contractor" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">CONTROLJOB</TD>	
	<TD align="left">DATECREATED</TD>	
	<TD align="left">ACTIVATION</TD>	
	<TD align="left">ADDRESS</TD>	
	<TD align="left">CITY</TD>	
	<TD align="left">STATE</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">BANKNAME</TD>	
	<TD align="left">BANKACCOUNT</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Building" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">BUILDING GROUP CODE</TD>	
	<TD align="left">ITEM</TD>	
	<TD align="left">ACTIVATION</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Allowded" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">BA</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">GROUP</TD>	
	<TD align="left">ALLDED</TD>	
	<TD align="left">CROP</TD>	
	<TD align="left">CONTROL JOB</TD>	
	<TD align="left">MAN AUTO</TD>	
	<TD align="left">TYPE 1</TD>	
	<TD align="left">TYPE 2</TD>	
	<TD align="left">TYPE 3</TD>	
	<TD align="left">TYPE 4</TD>	
	<TD align="left">MIN AMT</TD>	
	<TD align="left">MAX AMT</TD>	
	<TD align="left">RATE 1</TD>	
	<TD align="left">RATE 2</TD>	
	<TD align="left">RATE 3</TD>	
	<TD align="left">RATE 4</TD>	
	<TD align="left">REMARKS</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Field" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">DIVISION</TD>	
	<TD align="left">HECTARAGEPLANTED</TD>	
	<TD align="left">TERRAINTYPE</TD>	
	<TD align="left">SOILTYPE</TD>	
	<TD align="left">OWNERSHIP</TD>	
	<TD align="left">ACTIVATION</TD>	
	<TD align="left">CROPTYPE</TD>	
	<TD align="left">PLANTINGDATE</TD>	
	<TD align="left">PLANTINGMATERIAL</TD>	
	<TD align="left">TOTALSTAND</TD>	
	<TD align="left">STATUS</TD>	
	<TD align="left">PREVSEMESTERABW</TD>	
	<TD align="left">PRESSEMESTERABW</TD>	
	<TD align="left">HARVESTINGBASEDABW</TD>	
	<TD align="left">MAINROAD</TD>	
	<TD align="left">COLLECTIONROAD</TD>	
	<TD align="left">MAINDRAIN</TD>	
	<TD align="left">COLLECTIONDRAIN</TD>	
	<TD align="left">SUBSIDIARYDRAIN</TD>	
	<TD align="left">COLLECTIONPOINT</TD>	
	<TD align="left">HARVESTINGBRIDGE</TD>	
	<TD align="left">PERMANENTBRIDGE</TD>	
	<TD align="left">WOODBRIDGE</TD>	
	<TD align="left">BUIST</TD>	
	<TD align="left">OWLNEST</TD>
	<TD align="left">HARVESTING SECTION</TD>
	<TD align="left">FIELD INFO 01</TD>	
	<TD align="left">FIELD INFO 02</TD>	
	<TD align="left">FIELD INFO 03</TD>	
	<TD align="left">FIELD INFO 04</TD>	
	<TD align="left">FIELD INFO 05</TD>
	<TD align="left">ISACTIVE_HARVESTED</TD>	
	<TD align="left">VALID FROM</TD>		
	</tr>
   <%ElseIf Request("txttable") ="Station" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">JOB CODE</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Machine" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="1">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">FIXED ASSET</TD>	
	<TD align="left">OWNERSHIP</TD>	
	<TD align="left">DATE PURCHASE</TD>	
	<TD align="left">PURCHASE COST</TD>	
	<TD align="left">STATION CODE</TD>	
	<TD align="left">CONTROL JOB</TD>
	<td align="left">ACTIVATION</td>	
	</tr>
   <%END IF%>