
<%

    strPath           			= "http://www.HarvestAmerican.info/"
    strPathBase	      			= strPath & "ha_BackOffice/"
    strPathBOBase     			= strPath & "ha_BackOffice/"
    strPathIncludes	  			= strPathBase & "includes/"
    strPathBOIncludes 			= strPathBase & "includes/"
    strPathImages     			= strPathBase & "images/"
    strPathServers    			= strPathBase & "Servers/"
    strPathTelephones 			= strPathBase & "Telephones/"
    strPathNotifications 		= strPathBase & "ServiceNotifications/"
    strPathEmployees 			= strPathBase & "Employees/"
    strPathCompany 				= strPathBase & "Company/"
	
	Dim strPage
	strPage= Request.ServerVariables("SCRIPT_NAME")
	strPage= Right(strPage,Len(strPage) - InStrRev(strPage,"/"))

	Function iif(condition, consequent, alternative)
		If condition Then iif=consequent Else iif=alternative
	End Function
	
	Function RequestParse(r, d)
		If isnull(r) or r = "" or instr(r,"'") Then
			RequestParse = d
		Else
			RequestParse = r
		End If
	End Function	
%>


<script language="javascript" type="text/javascript">
    var strPathBase 	= '<%=strPathBase%>';
    var strPathIncludes = '<%=strPathIncludes%>';
    var strPathImages 	= '<%=strPathImages%>';
</script>


  <!--#include virtual="ha_backoffice/includes/FunctionsPermissions.inc"-->
