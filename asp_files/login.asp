<html>
<head>
<title>Fast Car Rentals -- Please Login...</title>
<link rel="stylesheet" type="text/css" href="../styles/carrent.css" />
</head>

<body>
<FORM ACTION="checklogin.asp" METHOD="POST">
<div align="center"> 

<img src="../images/logo.jpg" />

<br><br><br><br><br>

<% if Session("loginFailure") = True then %>

<h1>ACCESS DENIED!</h1>

</div>
</form>
</body>


<% else %>

<h2>Please enter your user name (oktork) and password (master)</h2>

User Name : <input type="text" name="txtUserName" length=25><br>
Password : <input type="password" name="txtPassword" length=25><br><br><br>
<input type="submit" value="Login">
<input type="reset" value="Reset">
</div>
</FORM>
</body>



<% end if %>

</html>