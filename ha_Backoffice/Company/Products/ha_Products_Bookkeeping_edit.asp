<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

  <!--#include virtual="includes/connection.asp"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsHelp.inc"-->
  <!--#include virtual="ha_backoffice/includes/FunctionsMisc.inc"-->
  
<%

  	Set connLocal = Server.CreateObject("ADODB.Connection")
	connLocal.Open "driver=SQL Server; server=127.0.0.1,7043; uid=davewj2o; pwd=get2it; database=ha_PublishedTables"
	Dim rs: Set rs = Server.CreateObject("ADODB.Recordset")

	If IsNull(Request("ID")) or Request("ID")=0 Then
	Else
		ID = Request("ID")
		
		strSQL = "SELECT * FROM ha_Products WHERE ID = " & ID
		rs.Open strSQL, connLocal, 3, 3

		If Not rs.EOF Then
			eItemID					= rs("ItemID")
			eDescription 			= rs("Description")

			eBreakoutHead1Name 		= rs("BreakoutHead1Name")
			eBreakoutHead1Sequence 	= rs("BreakoutHead1Seq")
			eBreakoutHead2Name 		= rs("BreakoutHead2Name")
			eBreakoutHead2Sequence 	= rs("BreakoutHead2Seq")
			
			eGLAccountHarvest 		= rs("GLAccount")
			eGLSequenceHarvest 		= rs("GLSequence")
			eGLAccountZip2Tax 		= rs("GLAccountZip2Tax")
			eGLSequenceZip2Tax 		= rs("GLSequenceZip2Tax")

			eCreatedBy				= Null2Space(rs("CreatedBy"))
			eCreatedDate			= Null2Space(rs("CreatedDate"))
			eEditedBy				= Null2Space(rs("EditedBy"))
			eEditedDate				= Null2Space(rs("EditedDate"))
		End If
	End If
  
	rs.Close
%>

<html>
<head>
  <title>Harvest American Backoffice - Edit Bookkeeping</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  
  <link href="/ha_BackOffice/includes/ha_BackofficePopup.css" type="text/css" rel="stylesheet" media="screen" />
  <script type="text/javaScript" src="/ha_BackOffice/includes/haFunctions.js"></script>
  
  <style type="text/css">
    input
		{
		width: 50px;   
		}
		
    input.long
		{
		width: 300px;   
		}
		
  </style>
</head>

<body onload="SetScreen(850,650)">
  <form action="ha_Products_Requirements_post.asp?ID=<%=ID%>" method="post" name="frm">

	<span class="popupHeading">Product Bookkeeping Edit</span>

    <table width="95%" border="0" cellspacing="2" cellpadding="2">

		<tr><td>&nbsp;</td></tr>
		
		<tr>
		  <td width="40%" align="right">
			Item ID:
		  </td>
		  <td width="60%" style="color: #1F3A4A; font-weight: bold;">
			<%=eItemID%>
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Description:
		  </td>
		  <td style="color: #1F3A4A; font-weight: bold;">
			<%=eDescription%>
		  </td>		 
		</tr>
		
		<tr><td>&nbsp;</td></tr>
  
		<tr>
		  <td align="right">
			Breakout Heading 1 Name:
		  </td>
		  <td>
			<input type="text" name="BreakoutHead1Name" id="BreakoutHead1Name"
				class="long" value="<%=eBreakoutHead1Name%>">
		  </td>		 
		</tr>
  
		<tr>
		  <td align="right">
			Breakout Heading 1 Sequence:
		  </td>
		  <td>
			<input type="text" name="BreakoutHead1Sequence" id="BreakoutHead1Sequence"
				value="<%=eBreakoutHead1Sequence%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Breakout Heading 2 Name:
		  </td>
		  <td>
			<input type="text" name="BreakoutHead2Name" id="BreakoutHead2Name"
				class="long" value="<%=eBreakoutHead2Name%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			Breakout Heading 2 Sequence:
		  </td>
		  <td>
			<input type="text" name="BreakoutHead2Sequence" id="BreakoutHead2Sequence"
				value="<%=eBreakoutHead2Sequence%>">
		  </td>		 
		</tr>
		
		<tr><td>&nbsp;</td></tr>

		<tr>
		  <td align="right">
			General Ledger Account - Harvest:
		  </td>
		  <td>
			<input type="text" name="GLAccountHarvest" id="GLAccountHarvest"
				class="long" value="<%=eGLAccountHarvest%>">
		  </td>		 
		</tr>
  
		<tr>
		  <td align="right">
			General Ledger Sequence - Harvest:
		  </td>
		  <td>
			<input type="text" name="GLSequenceHarvest" id="GLSequenceHarvest"
				value="<%=eGLSequenceHarvest%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			General Ledger Account - Zip2Tax:
		  </td>
		  <td>
			<input type="text" name="GLAccountZip2Tax" id="GLAccountZip2Tax"
				class="long" value="<%=eGLAccountZip2Tax%>">
		  </td>		 
		</tr>
		
		<tr>
		  <td align="right">
			General Ledger Sequence - Zip2Tax:
		  </td>
		  <td>
			<input type="text" name="GLSequenceZip2Tax" id="GLSequenceZip2Tax"
				value="<%=eGLSequenceZip2Tax%>">
		  </td>		 
		</tr>
		
	</table>
    	
    <span class="popupCredit"><%=CreditLine(eCreatedBy, eCreatedDate, eEditedBy, eEditedDate)%></span>	

    <span class="popupButtons">
      <a href="javascript:document.frm.submit();" class="bo_Button80">Submit</a>
      <a href="javascript:window.close();" class="bo_Button80">Close</a>
    </span>	
	
  </form>
</body>
</html>
