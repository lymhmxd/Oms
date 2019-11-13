<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>TIENS运维管理平台</title>
	<link rel="stylesheet" href="/cssware/gobal.css">
	<link rel="stylesheet" href="/cssware/nav.css">
	<link rel="stylesheet" href="/cssware/body.css">

	<!--#include virtual="database/connect.asp"-->
</head>

<body style="background-color: lightyellow;">
	<div id="nav">
		<div>
			<%
				set cnnmariadb = Server.CreateObject("ADODB.Connection")
				cnnmariadb.open connectstring
				sqlMenuStrcture = "select * from MenuStrcture"
				set rsMenuStrcture = cnnmariadb.execute(sqlMenuStrcture)
				While Not rsMenuStrcture.EOF
					Response.Write "<span><a href='" & rsMenuStrcture.Fields("MenuUrl") & "'>" & rsMenuStrcture.Fields("MenuName") & "</a></span>"
					rsMenuStrcture.MoveNext
				Wend
				rsMenuStrcture.Close
				Set rsMenuStrcture = Nothing
			%>
		</div>
	</div>
	<div id="content">
		<table cellspacing=0px cellpadding=0px >
			<tr>
				<th>位置</th>
				<%
					sql="SELECT * FROM Dashboard"
					set rs = cnnmariadb.execute(sql)
					While Not rs.EOF
						Response.Write "<th>" & rs.Fields("title") & "</th>"
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</tr>
			<% 	Codepage="65001"
				sqlConnection = "SELECT * FROM Connection"
				set rsConnection = cnnmariadb.execute(sqlConnection) '打开Connection的游标
				While Not rsConnection.EOF
					'建立到目标数据库的链接
					set cnnTarget = Server.CreateObject("ADODB.Connection") 
					cnnTarget.open connectstr(rsConnection.Fields("Driver"),rsConnection.Fields("server"),rsConnection.Fields("database"),rsConnection.Fields("uid"),rsConnection.Fields("pwd"))
					'在主数据库中找到要对目标数据库进行的查询集合，并遍历每一条查询
					sqlDashboard="select * from Dashboard"
					set rsDashboard = cnnmariadb.execute(sqlDashboard)
					Response.Write "<tr><td>" & rsConnection("Name") & "</td>"
					While Not rsDashboard.EOF
						set rsResult = cnnTarget.execute(rsDashboard.Fields("sqlchar")) '打开目标数据库的游标
						Response.Write "<td>"
						While Not rsResult.EOF
							Response.Write rsResult.Fields("status") & "~" & rsResult.Fields("amount") & "~" & rsResult.Fields("Percent") & "<br />" '输出目标数据库的遍历结果
							rsResult.MoveNext
						Wend
						Response.Write "</td>"
						rsResult.Close
						Set rsResult = Nothing
						rsDashboard.MoveNext
					Wend
					Response.Write "</tr>"
					rsdashboard.Close
					Set rsdashboard = Nothing
					cnnTarget.close
					Set cnnTarget = Nothing
					rsConnection.MoveNext
				Wend
				rsConnection.Close
				Set rsConnection = Nothing
			%>
		</table>
</body>
<%	
	cnnmariadb.Close
	Set cnnmariadb = Nothing
%>

</html>