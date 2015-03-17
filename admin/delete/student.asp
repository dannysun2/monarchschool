<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<meta http-equiv="X-UA-Compatible" content="IE=7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>The Monarch School</title>
<style type="text/css" media="all">
@import url("../css/style.css");
</style>
<!--[if lt IE 8]><style type="text/css" media="all">@import url("../css/ie.css");</style><![endif]-->
</head>
<body>
	<div id="hld">
		<div class="wrapper">
			<div id="header">
				<div class="hdrl"></div>
				<div class="hdrr"></div>
				<h1><a href="../index.asp">The Monarch School</a></h1>
				<ul id="nav">
					<li><a href="../index.asp">Dashboard</a></li>
					<li><a href="../add/student.asp">Add Student</a></li>
					<li><a href="../reports.asp">Reports</a></li>		
					<li><a href="../help.asp">Help</a></li>
				</ul>
				<div class="user">
					Today is 
						<script>
							var d=new Date();
							var weekday=new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
							var monthname=new Array("January","February","March","April","May","June","July","August","September","October","November","December");
								document.write(weekday[d.getDay()] + " ");
								document.write(monthname[d.getMonth()] + " ");
								document.write(d.getDate() + ", ");
								document.write(d.getFullYear());
						</script>
				</div>
			</div>
<%
dim cn
dim tokenvalue

sub pass1
sid = request.querystring("sid")

SQLString="SELECT * FROM student WHERE student_id="+cstr(sid)

set rs=Server.CreateObject("ADODB.Recordset")
sid=request.querystring("sid")

rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
name = cstr(rs("stu_fname")) + " " + cstr(rs("stu_lname"))
rs.movenext
wend
rs.close
set rs=nothing
%>
			<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>confirm</h2>
				</div>
				<div class="block_content">
					<form name="delete" action="student.asp" method="POST">	
					<div class="message warning"><p>Are you sure you want to delete the student <u><% =name %></u>? Deleting Student will result in a permanent loss of both student and associated grades.</p></div>
					<input type="hidden" name="token" value="2">
					<input type="hidden" name="sid" value="<% = cstr(sid) %>"><p>
					<p>
					<input type="submit" class="submit long" value="Submit" />
					</p>
					</form>
<%
end sub
sub pass2

studentid=request.form("sid")
set cn=Server.CreateObject("ADODB.Connection")
cn.open "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

SQLDelete1="DELETE FROM student WHERE student_id="+cstr(studentid)
SQLDelete2="DELETE FROM grades WHERE student_id="+cstr(studentid)
%>
<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>Alert</h2>
				</div>
				<div class="block_content">
					<form name="success" action="" method="">	
						<div class="message success"><p>Student record successfully deleted!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="javascript:window.location.href='../index.asp'" />
					</p>
					</form>
<%
cn.execute SQLDelete1
cn.execute SQLDelete2
end sub

tokenvalue=request.form("token")
if tokenvalue="" then
tokenvalue=request.querystring("token")
end if

select case tokenvalue
case "1"
	call pass1
case "2"
	call pass2
end select
%>
</div>
				<div class="bendl"></div>
				<div class="bendr"></div>				
			</div>
			<br>
      <div id="footer">
        <center>
          <img src="../images/footer.png">
          <b><p>The Monarch School</b></p>
        </center>
      </div>
    </div>
  </div>
</body>
</html>