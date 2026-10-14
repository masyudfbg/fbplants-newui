<tr bgcolor="#0000CC">
<%If txtTable="PaymentVoucherDetail" then%>     
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Voucher Code</font></td>
     <td><font color="#FFFFFF">Jobcode</font></td>
     <td><font color="#FFFFFF">Job Desc.</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Amount</font></td>
     <td><font color="#FFFFFF">Ref.</font></td>
     <td><font color="#FFFFFF">Remark</font></td>
<%elseif txtTable="PurchaseOrderDetails" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Supplier</font></td>
     <td><font color="#FFFFFF">Stock Code</font></td>
     <td><font color="#FFFFFF">Stock Desc</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>
     <td><font color="#FFFFFF">Amount</font></td>
     <td><font color="#FFFFFF">Remarks</font></td>
<%elseif txtTable="ReceiveNote" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">PO.NO</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Item Desc</font></td>
     <td><font color="#FFFFFF">Qty GRN</font></td>
     <td><font color="#FFFFFF">Qty Return</font></td>
     <td><font color="#FFFFFF">Amt GRN</font></td>
     <td><font color="#FFFFFF">Amt Return</font></td>
<%elseif txtTable="Invoice" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Supplier</font></td>
     <td><font color="#FFFFFF">Item. Code</font></td>
     <td><font color="#FFFFFF">Item. Desc</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>
     <td><font color="#FFFFFF">Amount</font></td>
<%elseif txtTable="IssueDetails" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">SIV Code</font></td>
     <td><font color="#FFFFFF">Stock Code</font></td>
     <td><font color="#FFFFFF">Item Desc</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>
     <td><font color="#FFFFFF">Price</font></td>
     <td><font color="#FFFFFF">Amount</font></td>
     <td><font color="#FFFFFF">Job Desc</font></td>
<%elseif txtTable="Attendance_GAD" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Emp.Code</font></td>
     <td><font color="#FFFFFF">Emp.Name</font></td>
     <td><font color="#FFFFFF">Job Desc.</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Mandays</font></td>
     <td><font color="#FFFFFF">Units</font></td>
     <td><font color="#FFFFFF">Rate</font></td>
     <td><font color="#FFFFFF">Output</font></td>
     <td><font color="#FFFFFF">OT</font></td>
<%elseif txtTable="EmpAllded" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Emp Code</font></td>
     <td><font color="#FFFFFF">Emp Name</font></td>
     <td><font color="#FFFFFF">All.Desc.</font></td>
     <td><font color="#FFFFFF">Amount</font></td>
<%elseif txtTable="HarvestingQuality" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Emp.Code</font></td>
     <td><font color="#FFFFFF">Emp Name</font></td>
     <td><font color="#FFFFFF">Block</font></td>
     <td align="right"><font color="#FFFFFF">Day</font></td>
     <td align="right"><font color="#FFFFFF">Month</font></td>
     <td align="right"><font color="#FFFFFF">year</font></td>
     <td align="right"><font color="#FFFFFF">Under Ripe</font></td>
     <td align="right"><font color="#FFFFFF">Over Ripe</font></td>
     <td align="right"><font color="#FFFFFF">Abnormal</font></td>
     <td align="right"><font color="#FFFFFF">Long Stalk</font></td>
     <td align="right"><font color="#FFFFFF">Eaten by Rat</font></td>
     <td align="right"><font color="#FFFFFF">unHarvest</font></td>
     <td align="right"><font color="#FFFFFF">Un Coll Circle</font></td>
     <td align="right"><font color="#FFFFFF">Un Coll Piece</font></td>
     <td align="right"><font color="#FFFFFF">UnArrange FFB</font></td>
     <td align="right"><font color="#FFFFFF">UnArrange Frond</font></td>
<%elseif txtTable="WorkshopActivity" then%>      
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">WS Code</font></td>
     <td><font color="#FFFFFF">Job Desc</font></td>	 
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Time & Description</font></td>
     <td align="right"><font color="#FFFFFF">Hours</font></td>
     <td align="right"><font color="#FFFFFF">Amount</font></td>
<%elseif txtTable="VehicleActivity" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">VH Code</font></td>
     <td><font color="#FFFFFF">Job Code</font></td>	 
     <td><font color="#FFFFFF">Job Desc</font></td>	 
     <td><font color="#FFFFFF">Time</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td align="right"><font color="#FFFFFF">Hours</font></td>
     <td align="right"><font color="#FFFFFF">Unit</font></td>
     <td align="right"><font color="#FFFFFF">Output</font></td>
     <td align="right"><font color="#FFFFFF">Amount</font></td>
<%elseif txtTable="AGREEMENTDETAIL" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Code Line</font></td>
     <td><font color="#FFFFFF">Job Code</font></td>
     <td><font color="#FFFFFF">Job Desc</font></td>	 
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Due Date</font></td>
     <td><font color="#FFFFFF">Retention</font></td>
     <td align="right"><font color="#FFFFFF">Qty</font></td>
     <td align="right"><font color="#FFFFFF">Rate</font></td>
<%elseif txtTable="WIP" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td align="right"><font color="#FFFFFF">Line</font></td>
     <td align="right"><font color="#FFFFFF">Quantity</font></td>	 	 
<%elseif txtTable="Nursery" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Desc</font></td>
     <td><font color="#FFFFFF">DATEPLANTED</font></td>
     <td><font color="#FFFFFF">CROP</font></td>	 
     <td><font color="#FFFFFF">PLANTINGMATERIAL</font></td>
     <td><font color="#FFFFFF">CLONE</font></td>
     <td><font color="#FFFFFF">SOURCE</font></td>
     <td><font color="#FFFFFF">TYPE</font></td>
     <td><font color="#FFFFFF">STATUS</font></td>
     <td><font color="#FFFFFF">SOURCETYPE</font></td>
     <td><font color="#FFFFFF">SOURCECODE</font></td>
     <td align="right"><font color="#FFFFFF">QUANTITY</font></td>
<%elseif txtTable="NurseryReceived" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Received Date</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>
<%elseif txtTable="NurseryCulling" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Culling Date</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>
<%elseif txtTable="NurseryIssue" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Code</font></td>
     <td><font color="#FFFFFF">Issue Date</font></td>
     <td><font color="#FFFFFF">Loc Type</font></td>
     <td><font color="#FFFFFF">Loc Code</font></td>
     <td><font color="#FFFFFF">Jobcode</font></td>
     <td><font color="#FFFFFF">Quantity</font></td>	 
 <%elseif txtTable="Allowded" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">All. Code</font></td>
     <td><font color="#FFFFFF">Desc.</font></td>
     <td><font color="#FFFFFF">Job Code</font></td>
<%elseif txtTable="Harvesting" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Chit.No</font></td>
     <td><font color="#FFFFFF">Blok</font></td>
     <td><font color="#FFFFFF">Crop</font></td>
     <td><font color="#FFFFFF">Product</font></td>	 
     <td><font color="#FFFFFF">Collec.</font></td>
     <td><font color="#FFFFFF">Harv</font></td>
     <td><font color="#FFFFFF">Reception</font></td>
     <td><font color="#FFFFFF">Veh.</font></td>
     <td><font color="#FFFFFF">Loader</font></td>
     <td><font color="#FFFFFF">Ramp.</font></td>
     <td align="right"><font color="#FFFFFF">Weight</font></td>
     <td align="right"><font color="#FFFFFF">QUANTITY</font></td>
<%elseif txtTable="HarvestingRound" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Year</font></td>
     <td><font color="#FFFFFF">Month</font></td>
     <td><font color="#FFFFFF">Field</font></td>
     <td><font color="#FFFFFF">Crop</font></td>	 
     <td><font color="#FFFFFF">No of Round.</font></td>
     <td><font color="#FFFFFF">Hect.</font></td>
     <td><font color="#FFFFFF">StandPerHa</font></td>
<%elseif txtTable="Weather" then%>
     <td><font color="#FFFFFF">No</font></td>
     <td><font color="#FFFFFF">Date</font></td>
     <td><font color="#FFFFFF">Div.</font></td>
     <td><font color="#FFFFFF">Inche</font></td>
     <td><font color="#FFFFFF">Days</font></td>	 
     <td><font color="#FFFFFF">Avg.Rain</font></td>
     <td><font color="#FFFFFF">Avg.Temp.</font></td>
     <td><font color="#FFFFFF">Avg.Wind</font></td>
<%end if%>
