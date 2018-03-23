<!--#include virtual="ha_backoffice/includes/Config.asp"-->
<!--#include virtual="includes/connection.asp"-->

<%
	ColorTab = 2
	PageHeading = "Telephone PBX Configuration Notes"
%>

<html>
<head>
    <title>Harvest American Backoffice - Telephone PBX Configuration Notes</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

    <link href="<%=strPathBOIncludes%>ha_Backoffice.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%=strPathBOIncludes%>menu/dropdowntabfiles/halfmoontabs.css" />

    <script type="text/javascript" src="<%=strPathBOIncludes%>menu/dropdowntabfiles/dropdowntabs.js"></script>
	<script type="text/javascript" src="<%=strPathBOIncludes%>haFunctions.js" language="javascript"></script>
	<script language='JavaScript'>

	function clickPopup(ID)
		{
			var URL = '/ha_Backoffice/Company/DocumentMaintenance/ha_Document_Show.asp' +
			'?PhotoID=' + ID;
			window.open(URL,'','height=490, width=600')
		}
		
	</script>
    <style type="text/css">
		span.subHead
			{
			font-weight: 	bold;
			color:			#336600;
			}
		span.newStyleHeading
			{
			font-size: 		11pt;
			font-weight: 	bold;
			color:			#336600;
			}	
    </style>

</head>

<body class="gray_desktop">
  <table align="center" cellspacing="0" cellpadding="0" class="MainBody">
    <tr>
		<td>

			<!--#include virtual="ha_Backoffice/includes/bodyparts/ha_PageHeading.inc"-->
			<table cellpadding="10">
				<tr>
					<td>					
					<span class="newStyleHeading">PBX Maintenance Console</span>
                        <p>
                        The <b>Panasonic KX-TDE-100</b> is managed by a computer program called the <b>PBX Unified Maintenance Console</b><br>
                        which has been installed on <b>Camden05</b>. It has three different levels of access:<br>
                        <ul>
                            <li><b>INSTALLER</b> - Installer Level -  Highest Permission</li>
                            <li><b>ADMIN</b> - Administrator Level - Middle Permission</li>
                            <li><b>USER</b> - User Level - Lowest Permission</li>
                        </ul><br>
                        The levels each have their own programmer code (INSTALLER, ADMIN, USER) and their own password (Default: 1234).<br> 
                        </p>
                    <hr>
					</td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Changing Maintenance Console Levels</span>
                        <p>
                        For most day to day uses the Administrator Level is fine, when you need to add new phones or<br>
                        reconfigure parameters for the internal PBX cards the installer level is used.<br>
                        To Change Permission Level:<br>
                        <ol>
                            <li>Open <b>PBX Unified Maintenance Console</b></li>
                            <li>Go to <b>Options(T)</b> and uncheck <b>Automatic login from next time?</b></li>
                            <li>Exit and restart the <b>PBX Unified Maintenance Console</b></li>
                            <li>For more information read the prompt given upon reload or just press <b>OK</b></li>
                            <li>In the next window enter the code for the permission level you want and press <b>OK</b></li>
                        </ol>
                        </p>
                     <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Saving Configuration to PBX SD Card</span>
                        <p>
                            The PBX Automatically saves configuration data to its SD Card upon exiting the Maintenance Console. <br>
                            It can also be saved from the Maintenance Console manually: 
                            <ol>
                                <li>Open the <b>PBX Unified Maintenance Console</b> and press <b>Connect(C)</b></li>
                                <li>Connect via <b>LAN</b> and press <b>Connect(O)</b></li>
                                <li>Select <b>Tool(T)</b> then <b>SD Memory Backup</b> or click the <b>SD Memory Backup</b> button.</li>
                            </ol>
                        </p>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Resetting the PBX</span>
                        <p>
                        There are two types of resets for the PBX they are:<br>
                        <ol>
                            <li>Normal Reset - Just Reboots the PBX</li>
                            <li>System Initialise Reset - Reboot the PBX and Restore Factory Default</li>
                        </ol>
                        Unless you intend to delete all of the PBX configured settings the <b>System Initialize</b> switch should <br>
                        remain in the <b>Normal</b> position. <br>
                        <br>
                        There are three ways of resetting the PBX (in order of severity):<br>
                        <ol>
                            <li>The Reset Button</li>
                            <li>The Power Switch</li>
                            <li>The Power Switch with Unplugging</li>
                        </ol>
                        <br>
                        Using the reset button:<br>
                        <ol>
                            <li>Open the front case of the PBX</li>
                            <li>Hold the reset button for about 2 seconds</li>
                        </ol>
                        <br>
                        Using the power switch:<br> 
                        <ol>
                            <li>Open the front case of the PBX</li>
                            <li>Turn the power switch off then back on</li>
                        </ol>
                        <br>
                        Using the power switch with Unplugging:<br>
                        <ol>
                            <li>Open the front case of the PBX</li>
                            <li>Turn the power switch off</li>
                            <li>Unplug the PBX for at least 5 minutes</li>
                            <li>Plug the PBX back in and turn on the power switch</li>
                        </ol>
                        </p>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Backing Up PBX Configuration</span>
                        <p>
                            <ol>
                                <li>Open the <b>PBX Unified Maintenance Console</b> with the <b>INSTALLER</b> level.</li>
                                <li>Select the <b>Utility</b> menu and then select <b>File Transfer PBX (SD Card) to PC</b></li>
                                <li>Files are transfered one at a time. Transfer all files to the same folder. The files to transfer are:<br>
                                    <ul>
                                        <li>DMSYS - System Configuration</li>
                                        <li>LIC00-LIC99 (As of writing our PBX only has LIC00) - License Keys</li>
                                    </ul>
                                </li>
                                <li>After selecting the file and selecting its save location press <b>Save</b> to backup file.</li>
                                <li>Zip all transfered files with this name format: [Date][PBX]Backup.zip</li>
                            </ol>
                        </p>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">Restoring PBX Configuration</span>
                        <p>
                            <ol>
                                <li>Open the <b>PBX Unified Maintenance Console</b> with the <b>INSTALLER</b> level.</li>
                                <li>Follow the directions above to backup the old configuration if applicable.</li>
                                <li>Select the <b>Utility</b> menu and then select <b>File Transfer PC to PBX (SD Card)</b></li>
                                <li>Unzip the former PBX backup / Make sure you have DMSYS and any LIC files available.</li>
                                <li>Transfer LIC files first then transfer DMSYS</li>
                            </ol>
                        </p>
                    <hr>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="newStyleHeading">PBX Manuals</span>
                        <p>
                        Panasonic offers different manuals for the different configuration levels and ways of configuring the PBX.<br>
                        They are:<br>
                        <ul>
                            <li>User Manual - User Feature Set</li>
                            <li>Operator Manual - Administrator Feature Set</li>
                            <li>Installer Manual - Details Wiring and Card Installation</li>
                            <li>PC Programming Manual - Installer Level for Maintenance Console</li>
                            <li>PT Programming Manual - System Programming using Panasonic Telephone</li>
                            <li>Feature Guide - Describes all basic, optional and programmable features.</li>
                        </ul>
                        </p>
                    <hr>
                    </td>
				</tr>   
			</table>		

		
		
		</td>
	</tr>
  </table>
  <!--#include virtual="ha_BackOffice/includes/BodyParts/ha_PageTail.inc"-->
</body>
</html>

