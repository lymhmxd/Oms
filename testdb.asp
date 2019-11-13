<!DOCTYPE html>

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
</head>
<html>
<!--#include virtual="database/connect.asp"-->
<%
    set cnn = Server.CreateObject("ADODB.Connection")
    cnn.open connectstring
    sql = "SELECT * FROM Connection"
    set rs = cnn.execute(sql)
    While Not rs.EOF
        response.write connectstr(rs.Fields(2),rs.Fields(3),rs.Fields(4),rs.Fields(5),rs.Fields(6)) & "<br />"
        rs.MoveNext
    Wend
%>
<%
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
%>
</html>