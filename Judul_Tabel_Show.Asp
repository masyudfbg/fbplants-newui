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
	<TD align="left">TRIGGER SYSTEM</TD>	
	<TD align="left">BANK TYPE</TD>	
	</tr>
   <%ElseIf Request("txttable") ="Roles" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">ROLES CODE</TD>	
	<TD align="left">ROLES NAME</TD>	
	</tr>
   <%ElseIf Request("txttable") ="BankStock" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">ITEM TYPE</TD>
	<TD align="left">UOM</TD>
	<TD align="left">STOCK GROUP</TD>
	<TD align="left">STOCK SUB.GROUP</TD>
	<td align="left">ACTIVE MATERIAL</td>
	<TD align="left">MANUFACTURE PART NUMBER</TD>	
	<TD align="left">MANUFACTURE CODE</TD>	
	<TD align="left">SUPPLIER NUMBER</TD>	
	<TD align="left">SUPPLIER CODE</TD>	
	<TD align="left">IS LESS ZERO</TD>
	<TD align="left">VALID UNTIL</TD>
	<TD align="left">MIN VALUE</TD>
	<TD align="left">MAX VALUE</TD>
	<TD align="left">STOCK MINIMUM</TD>
	<TD align="left">IS BUDGET CONTROL</TD>
	</tr>
   <%ElseIf Request("txttable") ="StockGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">CONTROL JOB</TD>
	<TD align="left">ACTIVATION</TD>
	<TD align="left">GROUP TYPE</TD>
	</tr>
	<%ElseIf Request("txttable") ="StockSubGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">STOCK GROUP</TD>
	<TD align="left">ACTIVATION</TD>
	
	</tr>
	<%ElseIf Request("txttable") ="FixedAsset" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">ASSET CODE</TD>
	<TD align="left">ASSET GROUP CODE</TD>
	<TD align="left">ESTIMATE LIFE YEAR</TD>
	</tr>
	<%ElseIf Request("txttable") ="Asset" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>
	<TD align="left">UOM</TD>	
	<TD align="left">ASSET GROUP</TD>
	</tr>
	<%ElseIf Request("txttable") ="AssetGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">ESTIMATE LIFE YEAR</TD>	
	</tr>
	<%ElseIf Request("txttable") ="EmpPosition" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">ACTIVATION</TD>	



	</tr>
	<%ElseIf Request("txttable") ="SectionName" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	



	</tr>
	<%ElseIf Request("txttable") ="Product" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">CROP CODE</TD>
	<TD align="left">UNIT OF MEASUREMENT</TD>	
	</tr>
	<%ElseIf Request("txttable") ="UnitOfMeasure" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">IS ACTIVE</TD>	
	</tr>
	<%ElseIf Request("txttable") ="MillQuality" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">DIRT</TD>	
	<TD align="left">MOISTURE & IMPORITY</TD>	
	<TD align="left">FFA</TD>	
	<TD align="left">DOBI</TD> 	
	</tr>
	<%ElseIf Request("txttable") ="VehicleGroup" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>
	<TD align="left">VEHICLE TYPE</TD>	
	</tr>
	<%ElseIf Request("txttable") ="OpHarvRate" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>


	<TD align="left">CODE</TD>	
	<TD align="left">BJR SENSUS</TD>	
	<TD align="left">TIPE TANAH</TD>
	<TD align="left">BJR MINIMAL</TD>
	<TD align="left">BJR MAKSIMAL</TD>
	<TD align="left">BASIS QUANTITY</TD>
	<TD align="left">BASIS RATE</TD>
	<TD align="left">EXTRA BASIS</TD>
	<TD align="left">BASIS HA</TD>
	</tr>
   <%ElseIf Request("txttable") ="Supplier" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
	<TD align="left">TYPE</TD>	
	<TD align="left">CONTROL JOB</TD>	
	<TD align="left">DATE CREATED</TD>	
	<TD align="left">ACTIVATION</TD>	
	<TD align="left">ADDRESS</TD>	
	<TD align="left">CITY</TD>	
	<TD align="left">STATE</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">BANK</TD>	
	<TD align="left">BANK ACCOUNT</TD>
	<TD align="left">BANK OWNERSHIP</TD>
	<TD align="left">PHONE</TD>
	<TD align="left">FAX</TD>
	<TD align="left">CONTACT PERSON</TD>
	<TD align="left">EMAIL</TD>
	<TD align="left">NPWP</TD>
	
	</tr>
	<%ElseIf Request("txttable") ="Transporter" then%>
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
	<TD align="left">BANK OWNERSHIP</TD>	
	</tr>
	<%ElseIf Request("txttable") ="Pemeriksa" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">NAME</TD>	
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
	<TD align="left">EMPLOYEE TYPE</TD>	
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
	<TD align="left">CORP PRESENT GRADE</TD>	
	<TD align="left">CORP GRADE DETAIL</TD>	
	<TD align="left">CORP PREV GRADE</TD>	
	<TD align="left">CORP PREV DETAIL</TD>	
	<TD align="left">MARITAL</TD>	
	<TD align="left">SPOUSE</TD>	
	<TD align="left">DEPENDENT CHILD</TD>	
	<TD align="left">BPJS KETENAGAKERJAAN</TD>	
	<TD align="left">BPJS KETENAGAKERJAAN CATEGORY</TD>	
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
	<TD align="left">VEHICLESUBGROUPCODE</TD>	
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
	<TD align="left">WB REGISTERED</TD>	
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
	<TD align="left" rowspan="2">NO</TD>
	<TD colspan="2" rowspan="2">&nbsp;</TD>
	<TD align="left" rowspan="2">CODE</TD>	
	<TD align="left" rowspan="2">NAME</TD>	
	<TD align="left" rowspan="2">DIVISION</TD>	
	<TD align="left" rowspan="2">HECTARAGEPLANTED</TD>	
	<TD align="left" rowspan="2">TERRAINTYPE</TD>	
	<TD align="left" rowspan="2">SOILTYPE</TD>	
	<TD align="left" rowspan="2">OWNERSHIP</TD>	
	<TD align="left" rowspan="2">ACTIVATION</TD>	
	<TD align="left" rowspan="2">CROPTYPE</TD>	
	<TD align="left" rowspan="2">PLANTINGDATE</TD>	
	<TD align="left" rowspan="2">PLANTINGMATERIAL</TD>	
	<TD align="left" rowspan="2">TOTALSTAND</TD>	
	<TD align="left" rowspan="2">STATUS</TD>	
	<TD align="left" rowspan="2">PREVSEMESTERABW</TD>	
	<TD align="left" rowspan="2">PRESSEMESTERABW</TD>	
	<TD align="left" rowspan="2">HARVESTINGBASEDABW</TD>	
	<TD align="left" rowspan="2">MAINROAD</TD>	
	<TD align="left" rowspan="2">COLLECTIONROAD</TD>	
	<TD align="left" rowspan="2">MAINDRAIN</TD>	
	<TD align="left" rowspan="2">COLLECTIONDRAIN</TD>	
	<TD align="left" rowspan="2">SUBSIDIARYDRAIN</TD>	
	<TD align="left" rowspan="2">COLLECTIONPOINT</TD>	
	<TD align="left" rowspan="2">HARVESTINGBRIDGE</TD>	
	<TD align="left" rowspan="2">PERMANENTBRIDGE</TD>	
	<TD align="left" rowspan="2">WOODBRIDGE</TD>	
	<TD align="left" rowspan="2">BUIST</TD>	
	<TD align="left" rowspan="2">OWLNEST</TD>
	<TD align="left" rowspan="2">FIELD INFO 01</TD>	
	<TD align="left" rowspan="2">FIELD INFO 02</TD>	
	<TD align="left" rowspan="2">FIELD INFO 03</TD>	
	<TD align="left" rowspan="2">FIELD INFO 04</TD>	
	<TD align="left" rowspan="2">FIELD INFO 05</TD>
	<TD align="left" rowspan="2">ISACTIVE_HARVESTED</TD>	
	<TD align="left" rowspan="2">VALID FROM</TD>
	<TD align="center" colspan="17">LAND USE</TD>
	<TD align="center" colspan="6">INFORMASI PANEN</TD>
	<TD align="left" rowspan="2">Year</TD>
	<TD align="left" rowspan="2">Month</TD>
	
	</tr>
	<tr bgcolor="#0C5B24" style="color:#fff;">
		<td align="left">AREAL_LC</td>
		<td align="left">CADANGAN</td>
		<td align="left">OKUPASI</td>
		<td align="left">BIBITAN</td>	
		<td align="left">PABRIK</td>
		<td align="left">EMPLASMEN</td>
		<td align="left">PARIT</td>
		<td align="left">JALAN</td>
		<td align="left">PRASARANA LAINNYA</td>
		<td align="left">QUARRY</td>
		<td align="left">TANGGUL</td>
		<td align="left">ENCLAVE</td>
		<td align="left">TAMBANG</td>
		<td align="left">KONSERVASI</td>
		<td align="left">BUKIT SUNGAI</td>
		<td align="left">NO IMAGERY</td>
		<td align="left">LUAS GROSS</td>
		
		<td align="left">LUAS PANEN</td>
		<td align="left">POKOK PRODUKTIF</td>
		<td align="left">SEKSI PANEN</td>
		<td align="left">URUTAN SEKSI PANEN</td>
		<td align="left">PLANTING POINT</td>
		<td align="left">FULLSTAND</td>
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
	 <%ElseIf Request("txttable") ="WorkingTime" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">TimeIN</TD>	
	<TD align="left">TimeOut</TD>	
	<TD align="left">ISRESTIME</TD>	
	<TD align="left">ISNEXTDAY</TD>	
	<TD align="left">ADD IN</TD>	
	<TD align="left">ADD OUT</TD>
	<td align="left">ACTIVATION</td>		
	</tr>
	 <%ElseIf Request("txttable") ="WorkingTime_Template" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">GANGCODE</TD>	
	<TD align="left">PERIOD</TD>	
	<TD colspan="2" align="left">MONDAY</TD>	
	<TD colspan="2" align="left">TUESDAY</TD>	
	<TD colspan="2" align="left">WEDNESDAY</TD>	
	<TD colspan="2" align="left">THRUSDAY</TD>
	<TD colspan="2" align="left">FRIDAY</TD>
	<TD colspan="2" align="left">SATURDAY</TD>
	<TD colspan="2" align="left">SUNDAY</TD>
	<td align="left">ACTIVATION</td>	
	 <%ElseIf Request("txttable") ="Employee_WorkingTime" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD colspan="2" align="left">PERIOD</TD>	
	<TD colspan="2" align="left">01</TD>	
	<TD colspan="2" align="left">02</TD>	
	<TD colspan="2" align="left">03</TD>	
	<TD colspan="2" align="left">04</TD>
	<TD colspan="2" align="left">05</TD>
	<TD colspan="2" align="left">06</TD>
	<TD colspan="2" align="left">07</TD>
	<TD colspan="2" align="left">08</TD>
	<TD colspan="2" align="left">09</TD>
	<TD colspan="2" align="left">10</TD>
	<TD colspan="2" align="left">11</TD>
	<TD colspan="2" align="left">12</TD>
	<TD colspan="2" align="left">13</TD>
	<TD colspan="2" align="left">14</TD>
	<TD colspan="2" align="left">15</TD>
	<TD colspan="2" align="left">16</TD>
	<TD colspan="2" align="left">17</TD>
	<TD colspan="2" align="left">18</TD>
	<TD colspan="2" align="left">19</TD>
	<TD colspan="2" align="left">20</TD>
	<TD colspan="2" align="left">21</TD>
	<TD colspan="2" align="left">22</TD>
	<TD colspan="2" align="left">23</TD>
	<TD colspan="2" align="left">24</TD>
	<TD colspan="2" align="left">25</TD>
	<TD colspan="2" align="left">26</TD>
	<TD colspan="2" align="left">27</TD>
	<TD colspan="2" align="left">28</TD>
	<TD colspan="2" align="left">29</TD>
	<TD colspan="2" align="left">30</TD>
	<TD colspan="2" align="left">31</TD>
	<td align="left">ACTIVATION</td>	
	<%ElseIf Request("txttable") ="Koperasi" then%>
	<tr bgcolor="#0C5B24" style="color:#fff;" >	
		<td align="left">No</td>
		<th colspan="2">&nbsp;</td>
		<td align="left">Kode</td>
		<td align="left">Deskripsi</td>
		<td align="left">Tanggal Berdiri</td>
		<td align="left">No Badan Hukum</td>
		<td align="left">Tanggal Badan Hukum</td>
		<td align="left">Alamat</td>
		<td align="left">Kota</td>
		<td align="left">Negara</td>
		<td align="left">Telp</td>
		<td align="left">Fax</td>
		<td align="left">Bentuk Koperasi</td>
		<td align="left">Jenis Koperasi</td>
		<td align="left">Kelompok Koperasi</td>
		<td align="left">Sektor Usaha</td>
		<td align="left">NIK</td>
		<td align="left">Sertifikasi</td>
		<td align="left">Status</td>
	</tr>
   <%END IF%>