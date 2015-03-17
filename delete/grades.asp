<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<html lang="en">
<head>
<title>The Monarch School</title>
<meta http-equiv="x-ua-compatible" content="IE=8">
<meta charset="utf-8">
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
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
gid = request.querystring("gid")

SQLString="SELECT * FROM student WHERE student_id="+cstr(sid)

set rs=Server.CreateObject("ADODB.Recordset")
sid=request.querystring("sid")

rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
name = cstr(rs("stu_fname")) + " " + cstr(rs("stu_mname")) + " " + cstr(rs("stu_lname"))
rs.movenext
wend
rs.close
set rs=nothing

SQLString="SELECT *,CONVERT(varchar(10),date_taken,101) as dt_taken FROM grades WHERE grades_id="+cstr(gid)

set rs=Server.CreateObject("ADODB.Recordset")
sid=request.querystring("sid")

rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
dt_taken = cstr(rs("dt_taken"))
rs.movenext
wend
rs.close
set rs=nothing
%>
			<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2><% =name %> - delete test from <% =dt_taken %></h2>
				</div>
				<div class="block_content">
					<form name="delete" action="grades.asp" method="POST">	
					<div class="message warning"><p>Deleting grades will result in a permanent loss.</p></div>
					<p>
					<input type="hidden" name="token" value="2">
					<input type="hidden" name="sid" value="<% = cstr(sid) %>">
					<input type="hidden" name="gid" value="<% = cstr(gid) %>">
					<input type="submit" class="submit long" value="Submit" />
					</p>
					</form>
<%
end sub
sub pass2

studentid=request.form("sid")
gradeid=request.form("gid")

set cn=Server.CreateObject("ADODB.Connection")
cn.open "gl1181","gl1181","YVT52ddnJ"

SQLDelete1="DELETE FROM grades WHERE student_id="+cstr(studentid) +" AND grades_id=" +cstr(gradeid)
%>
<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
				<h2>Alert</h2>
				</div>
				<div class="block_content">
					<form name="success" action="" method="">	
					<div class="message success"><p>Grades successfully deleted!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="javascript:window.location.href='../view/student.asp?token=3&sid=<% =cstr(studentid) %>'" />
					</p>
					</form>
					
<%
cn.execute SQLDelete1
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