<% @LANGUAGE = "vbscript" %>

<% Option Explicit %>

<html>
<head>
<script type='text/javascript'>
			function RefreshHome() {
			    var selectCtrl = document.getElementById("cboFilter");
			    var selectedItem = selectCtrl.options[selectCtrl.selectedIndex];
			    document.location = selectedItem.value;
			}
			
			function AddCar() {
				//this sub redirects the user to the insertcar page
				document.location = "newcar.asp"
			}
			
</script>
<!-- #Include file="../App_Data/inc_files/FCRmdb.inc" -->



<title>Fast Car Rentals - Inventory</title>
<link rel="stylesheet" type="text/css" href="../styles/carrent.css" />
<!-- add the title logo -->
<div align="center">
<img src='../images/logo.jpg'>
</div>
</head>
<!-- Start the body of the page -->
<body>
<br><br>
<div align='center'>
<table class='id' width='70%'>
	<tr>
		<td class='id' width='35%' align='left' valign='center' NOWRAP>
		<!-- this next line adds the user name to the 'hello' -->
			<% Response.Write ("<h2>Hello, " & Session("UserName") & "!</h2>") %>
		</td>
		<td class='id' width='15%' align='center' valign='center'>
			<font class='h4'>Enter a new car</font><br>
			<input type='button' value='New Car' ONCLICK='AddCar()'>
		</td>
		<td class='id' width='50%' align='center'>
			<font class='h4'>Select a location to filter the inventory list.</font><br>
			<select id="cboFilter" name='cboFilter' ONCHANGE='RefreshHome()' size=1>

			<% 'while this function is not dynamic it works for the current scope of the project, 
				'further expansion by the company will be met with this obstacle.
				'my reluctance to spend the time using a recordset to fill the combo box
				'is tempered by the amount of time alotted for this project
			Select case Request("DesiredLocation")
				'this select statement fills the combo box, appropriatly selectingthe proper 
				'location, as per the user's request
				case "All", ""
					Response.Write ("<option value='home.asp' SELECTED>All")
					Response.Write ("<option value='home.asp?DesiredLocation=Victoria'>Victoria")
					Response.Write ("<option value='home.asp?DesiredLocation=Vancouver'>Vancouver")
					Response.Write ("<option value='home.asp?DesiredLocation=Nanaimo'>Nanaimo")
					Response.Write ("<option value='home.asp?DesiredLocation=Hope'>Hope")
				case "Victoria"
					Response.Write ("<option value='home.asp'>All")
					Response.Write ("<option value='home.asp?DesiredLocation=Victoria' SELECTED>Victoria")
					Response.Write ("<option value='home.asp?DesiredLocation=Vancouver'>Vancouver")
					Response.Write ("<option value='home.asp?DesiredLocation=Nanaimo'>Nanaimo")
					Response.Write ("<option value='home.asp?DesiredLocation=Hope'>Hope")
				case "Vancouver"
					Response.Write ("<option value='home.asp'>All")
					Response.Write ("<option value='home.asp?DesiredLocation=Victoria'>Victoria")
					Response.Write ("<option value='home.asp?DesiredLocation=Vancouver' SELECTED>Vancouver")
					Response.Write ("<option value='home.asp?DesiredLocation=Nanaimo'>Nanaimo")
					Response.Write ("<option value='home.asp?DesiredLocation=Hope'>Hope")
				case "Nanaimo"
					Response.Write ("<option value='home.asp'>All")
					Response.Write ("<option value='home.asp?DesiredLocation=Victoria'>Victoria")
					Response.Write ("<option value='home.asp?DesiredLocation=Vancouver'>Vancouver")
					Response.Write ("<option value='home.asp?DesiredLocation=Nanaimo' SELECTED>Nanaimo")
					Response.Write ("<option value='home.asp?DesiredLocation=Hope'>Hope")
				case "Hope"
					Response.Write ("<option value='home.asp'>All")
					Response.Write ("<option value='home.asp?DesiredLocation=Victoria'>Victoria")
					Response.Write ("<option value='home.asp?DesiredLocation=Vancouver'>Vancouver")
					Response.Write ("<option value='home.asp?DesiredLocation=Nanaimo'>Nanaimo")
					Response.Write ("<option value='home.asp?DesiredLocation=Hope' SELECTED>Hope")	
			end select
			%>
			</select>
		</td>
	</tr>
</table>
</form>
</div><br><br>
<div align='center'>

<table class='inv' bgcolor=#00ced1>
	<caption align='center' valign='center'><B><big><font>Inventory</font></big></b></caption>
	<tr>
		<!--one--><th align='center'>Car ID</td>
		<!--two--><th align='center'>Manufacturer</td>
		<!--three--><th align='center'>Make Name</td>
		<!--four--><th align='center'>Vehicle Type</td>
		<!--five--><th align='center'>Year</td>
		<!--six--><th align='center'>Colour</td>
		<!--seven--><th align='center'>Auto</td>
		<!--eight--><th align='center'>Doors</td>
		<!--nine--><th align='center'>AC</td>
		<!--ten--><th align='center'>Rates</td>
		<!--eleven--><th align='center'>Date Rented</td>
		<!--twelve--><th align='center'>KM</td>
		<!--thirteen--><th align='center'>Location</td>
		<!--fourteen--><th align='center'>Client ID</td>
	</tr>
<%
	Dim intRecords, intFields, strSQL
	select case request("DesiredLocation")
		case "All", ""
			set gobjRS = server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM tblCar"
			call gobjRS.Open(strSQL, gobjConn, adOpenStatic, adLockReadOnly, adcmdText)
		case else
			set gobjRS = server.CreateObject("ADODB.Recordset")
			strSQL = "Select * FROM tblCar WHERE fldLocation ='" & Request("DesiredLocation") & "'"
			call gobjRS.Open(strSQL, gobjConn, adOpenStatic, adLockReadOnly, adcmdText)
	end select
	for intRecords = 0 to gobjRS.RecordCount - 1
	Response.Write("<tr>")
	
		for intFields = 0 to gobjRS.Fields.count - 1
			
			select case cint(intFields)
				case 0
					'first check to see if it's rented
					if isnull(gobjRS.Fields(13)) then
						Response.Write ("<th class='lot'")
					else
						Response.Write ("<th class='rented'")
					end if
					Response.Write ("><a href='rentcar.asp?strCarID=" & gobjRS.Fields(intfields) & "'>")
					Response.Write (gobjRS.Fields(intFields))
					Response.Write ("</a>")
					Response.Write ("</th>")
				case 1,2,3,4,5,10,13
					Response.Write ("<td>")
					if not isnull(gobjRS.Fields(intfields)) then
						Response.Write (gobjRS.Fields(intfields))
					else
						Response.Write ("&nbsp;")
					end if
					Response.Write ("</td>")
				case 6,8
					response.write ("<td align='center'><input type='checkbox'")
					if cbool(gobjRS.Fields(intfields)) = true then
						Response.Write(" CHECKED")
					end if
					Response.write ("></td>")
				case 7
					Response.Write ("<td align='center'>" & gobjRS.Fields(intfields).Value & "</td>")
				case 9
					Response.Write ("<td align='center'>" & formatcurrency(gobjRS.Fields(intfields),2) & "</td>")
				case 11
					Response.Write ("<td align='right'>" & gobjRS.Fields(intfields) & "</td>")
				case 12
					Response.Write ("<td>")
					if gobjRS.fields(intfields) > " " then
						Response.Write (gobjrs.fields(intfields))
					else
						Response.Write ("&nbsp;")
					end if
					Response.Write ("</td>")
			end select
		next
	Response.Write ("</tr>")
	gobjRS.MoveNext
	next
	gobjRS.Close
	set gobjrs = nothing
%>

</table>
</div>
</body>
</html>

