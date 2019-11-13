<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<title>TIENS运维管理平台</title>
	<link rel="stylesheet" href="/cssware/gobal.css">
	<link rel="stylesheet" href="/cssware/nav.css">
	<link rel="stylesheet" href="/cssware/body.css">

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
				rs.Close
				Set rs = Nothing
			%>
		</div>
	</div>
	<div id="content">
		<table>
			<tr>
				<%
					sql="SELECT * FROM Dashboard"
					set rs = cnn.execute(sql)
					While Not rs.EOF
						Response.Write "<th>" & rs.Fields("title") & "</th>"
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</tr>
			</table>
		<%
			sqlConnection = "SELECT * FROM Connection"
			set rsConnection = cnn.execute(sqlConnection) '打开Connection的游标用于遍历每一个实例的Dashboard
			While Not rsConnection.EOF
				dim connectstringDashboard
				set cnn_Dashboard = Server.CreateObject("ADODB.Connection")
		    	cnnDashboard.open connectstr(rsConnection.Fields("Driver"),rsConnection.Fields("server"),rsConnection.Fields("database"),rsConnection.Fields("uid"),rsConnection.Fields("pwd"))
				sqldashboard="select * from Dashboard"
				set rsdashboard = cnn.execute(sqldashboard) '打开dashboard的游标用于轮询每一条执行的sql
				While Not rsdashboard.EOF
					sql_Dashboard = rsdashboard.Fields("sqlchar")
					set rs_Dashboard = cnn_Dashboard.execute(sql_Dashboard) '打开目标数据库的游标用于遍历每一条结果
					While Not rs_Dashboard.EOF
						Response.Write rs_Dashboard.Fields("status") '输出遍历结果
						rs_Dashboard.MoveNext
					Wend
					rs_Dashboard.Close
					Set rs_Dashboard = Nothing
					cnn_Dashboard.Close
					Set cnn_Dashboard = Nothing
					rsdashboard.MoveNext
				Wend
				rsdashboard.Close
				Set rsdashboard = Nothing
		    Wend
			rsConnection.Close
			Set rsConnection = Nothing
		%>
</body>
<%	
	cnn.Close
	Set cnn = Nothing
%>

</html>