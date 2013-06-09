<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>
<%@ page import="java.util.Map.Entry" %>
<%@ page session="false" %>
<html>
<head>
	<title>Home</title>
    <script type="text/javascript" src="../../resources/script.js"></script>
</head>
<body>
<h3>
	Data Import  
</h3>
<div id="container">
<a href="download?type=office">Offices</a>
<a href="download?type=staff">Staff</a>
<a href="download?type=client">Clients</a>
<a href="download?type=loanProducts">Loan Products</a>
<a href="download?type=loanAccounts">Loan Accounts</a>

<input type="submit" id="downloadSheet" value="Download"  />
</form>
<form method="post" action="import" enctype="multipart/form-data">
            <input type="file" name="file"/>
            <input type="submit"/>
</form>
</div>
</body>
</html>
