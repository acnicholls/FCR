<%@ Language=VBScript %>
<% Option Explicit %>
<% dim strSQL, intCounter, objrsType, objrsDoors, objrsLocation %>


<HTML>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->
	<HEAD>
		<!--This header first links to the stylesheet, then declares the title
			there is then a meta tag, input by the devel env. and a divtag containing the logo
			-->
<link rel="stylesheet" type="text/css" href="../styles/carrent.css" />
		<title>Fast Car Rental - Insert New Car</title>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<div align='center'><img src="../images/logo.jpg"></div>
		<%	
	if session("InsertFailure") = true then
		Response.Write("<div align='center'><font> Please be sure to enter all data </font><div>")
	end if

 %>
	</HEAD>
	<!--
	The body contains a form that collect the vehicles information and when


	-->
	<BODY>
    <div align="center">
	<form action='insertcar.asp' method='POST'>
		<table class='insert' rows=2 cols=1 width='40%' valign='top' align='center' height='250'>
			<tr class='insert' height='96%'>
				<td class='insert' align='center' >
					<font>Location: </font><Select name='cboLocation' size=1>
					<% 
					strSQL = "SELECT DISTINCT fldLocation FROM tblCar"
					set objrsLocation = Server.CreateObject("ADODB.Recordset")	
					call objrsLocation.open(strSQL, gobjConn, adopenStatic, adLockreadonly, adCmdText)
					objrsLocation.movefirst
					for intCounter = 1 to objrsLocation.recordcount
						Response.Write("<option value='" & objrsLocation.fields(0) & "'>" & objrsLocation.fields(0))
						objrsLocation.movenext
					next
					objrsLocation.close
					set objrsLocation = nothing
					%>
					</select><br>
					<font>Manufacturer: </font><input type='text' name='txtManu' size='20'><br>
					<font>Colour: </font><input type='text' name='txtcolour' size=15><br>
					<font>Make: </font><input type='text' name='txtMake' size='20'><br>
					<font>Type: </font><Select name='cboType' size=1>
					<% 
					strSQL = "SELECT DISTINCT fldType FROM tblCar"
					set objrsType = Server.CreateObject("ADODB.Recordset")
					call objrsType.open(strSQL, gobjConn, adopenStatic, adLockreadonly, adCmdText)
					objrsType.movefirst
					for intCounter = 0 to objrsType.recordcount - 1
						Response.Write("<option value='" & objrsType.fields(0) & "'>" & objrsType.fields(0))
						objrsType.movenext
					next
					objrsType.close
					set objrsType = nothing

					%>
					</select><br>
					<font>Rate: </font><input type='text' name='txtRate' size=5><br>
					<font>Auto: </font><input type='checkbox' name='chkAuto'><br>
					<font>AC: </font><input type='checkbox' name='chkAC'><br>
					<font>Doors: </font><Select name='cboDoors' size=1>
					<% 
					strSQL = "SELECT DISTINCT fldDoors FROM tblCar"
					set objrsDoors = Server.CreateObject("ADODB.Recordset")	
					call objrsDoors.open(strSQL, gobjConn, adopenStatic, adLockreadonly, adCmdText)
					objrsDoors.movefirst
					for intCounter = 1 to objrsDoors.recordcount
						Response.Write("<option value='" & objrsDoors.fields(0) & "'>" & objrsDoors.fields(0))
						objrsDoors.movenext
					next
					objrsDoors.close
					set objrsDoors = nothing

					%>
					</select><br>
					<font>Year: </font><input type='text' name='txtYear' size=5><br>
					<font>KM: </font><input type='text' name='txtKM' size=10><br>
				</td>
			</tr>
			<tr>
				<td class='insert' align='center'>
				<input type=submit value='Submit'>
				<input type=reset value='Reset'>
				</td>
			</tr>
		</table>
		</form>
        </div>
	</BODY>
</HTML>
