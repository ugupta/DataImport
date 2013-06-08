<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Map" %>
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
<form id="entities" action="download" >
Select Entity: &nbsp;&nbsp;&nbsp;
<select name="entity">
  <c:forEach items="${entities}" var="entity">
    <option value="${entity.key}">${entity.value}</option>
  </c:forEach>
</select>
<br/>
<input type="submit" id="downloadSheet" value="Download"  />
</form>
<form method="post" action="form" enctype="multipart/form-data">
            <input type="file" name="file"/>
            <input type="submit"/>
</form>
</div>
</body>
</html>
