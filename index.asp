<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>TIENS运维管理平台</title>
    <link rel="stylesheet" href="/cssware/gobal.css">
    <link rel="stylesheet" href="/cssware/nav.css">
    <!--#include virtual="database/connect.asp"-->
</head>

<body>
    <div id="nav">
        <div>
            <%
				set cnn = Server.CreateObject("ADODB.Connection")
				cnn.open connectstring
				sql = "select * from MenuStrcture"
				set rs = cnn.execute(sql)
				While Not rs.EOF
					Response.Write "<span><a href='" & rs.Fields("MenuUrl") & "'>" & rs.Fields("MenuName")& "</a></span>"
					rs.MoveNext
				Wend
			%>
        </div>
    </div>
</body>
<%	
	rs.Close
	Set rs = Nothing
	cnn.Close
	Set cnn = Nothing
%>

</html>