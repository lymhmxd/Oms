<!DOCTYPE html>
<html>
<!--#include virtual="database/connect.asp"-->
<%
    set cnn = Server.CreateObject("ADODB.Connection")
    cnn.open connectstring
    sql = "select * from MenuStrcture"
    set rs = cnn.execute(sql)
    While Not rs.EOF
        Response.Write rs.Fields("MenuUrl") & ", " & rs.Fields("PMenuId") & "<br>"
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
%>
</html>