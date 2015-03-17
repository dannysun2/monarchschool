<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<meta http-equiv="X-UA-Compatible" content="IE=7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>The Monarch School</title>
<style type="text/css" media="all">
@import url("../css/style.css");
@import url("../css/facebox.css");
@import url("../css/date_input.css");
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
					<li><a href="student.asp">Add Student</a></li>
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
dim rs

function convertDate(dateString)
convertDate = dateString
convertDate = mid(convertDate,7,4) + "-" + mid(convertDate,1,2) + "-" + mid(convertDate,4,2) + " 05:00:00"
end function

Function removeSlashes(gradeString)
	removeSlashes = Replace(gradeString,chr(47),"*")
	removeSlashes = Replace(removeSlashes,chr(92),"*")
end function

sub pass1

sid = request.querystring("sid")
if sid = "" then sid=request.form("sid")

name = ""

set rs = Server.CreateObject("ADODB.Recordset")
sql_string="Select * from student WHERE student_id=" +sid
rs.open sql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
level=""
while NOT rs.EOF
    level=cstr(rs("level"))
	name = cstr(rs("stu_fname")) + " " + cstr(rs("stu_lname"))
    rs.movenext
wend
rs.close
set rs=nothing

date_taken = request.form("date_taken")
inhibit = request.form("inhibit")
shift = request.form("shift")
emotion = request.form("emotion")
initiate = request.form("initiate")
work_mem = request.form("work_mem")
plan_org = request.form("plan_org")
org_of_mat = request.form("org_of_mat")
monitor = request.form("monitor")

dt_error = ""
inhibit_error=""
shift_error=""
emotion_error=""
ini_error=""
wm_error=""
po_error=""
om_error=""
mon_error=""
error_count=0

if tokenvalue >1 then
name = request.form("name")

if date_taken = "" and tokenvalue > 1 then
dt_error = "test date is required."
error_count = error_count + 1

elseif not isdate(date_taken) then
error_count = error_count + 1
dt_error = "Invalid test date."

elseif date_taken <> "" and tokenvalue > 1 then

if mid(date_taken,2,1) = "/" then date_taken = "0" + date_taken

if mid(date_taken,5,1) = "/" then date_taken = mid(date_taken,1,3) + "0" + mid(date_taken,4,6)

sql_string = "select * from grades where student_id = '" + cstr(sid) + "' and date_taken = '" + cstr(convertDate(date_taken)) + " 05:00:00'"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

c=0
while not rs.EOF
c=c+1
  rs.MoveNext
wend

rs.Close
set rs=nothing
	
	if c>0 then
	dt_error = "Test for the selected student and date already exists. Please modify the existing test or select a different date."
	error_count = error_count + 1
	end if
end if

if initiate = "" then
ini_error = "Initiate grade is required."
error_count = error_count + 1

else
	initiate = removeSlashes(initiate)
	if not isnumeric(initiate) then
	ini_error = "Initiate grade must be numeric."
	error_count = error_count + 1
	elseif cint(initiate) > 100 then
	ini_error = "Initiate grade may not be over 100."
	error_count = error_count + 1
	elseif cint(initiate) < 0 then
	ini_error = "Initiate grade may not be under 0."
	error_count = error_count + 1
	elseif cint(initiate) >=0 and cint(initiate) <= 100 then
		for i = 1 to len(initiate)
			if mid(initiate,i,1) = "." then
				error_count = error_count + 1
				ini_error = "Initiate grade may not contain decimals."
			end if
		next
	end if
end if

if work_mem = "" then
wm_error = "Working Memory grade is required."
error_count = error_count + 1
else
	work_mem = removeSlashes(work_mem)
	if not isnumeric(work_mem) then
	wm_error = "Working Memory grade must be numeric."
	error_count = error_count + 1
	elseif cint(work_mem) > 100 then
	wm_error = "Working Memory grade may not be over 100."
	error_count = error_count + 1
	elseif cint(work_mem) < 0 then
	wm_error = "Working Memory grade may not be under 0."
	error_count = error_count + 1
	elseif cint(work_mem) >=0 and cint(work_mem) <= 100 then
		for i = 1 to len(work_mem)
			if mid(work_mem,i,1) = "." then
				error_count = error_count + 1
				wm_error = "Working Memory grade may not contain decimals."
			end if
		next
	end if
end if

if plan_org = "" then
po_error = "Plan & Organize grade is required."
error_count = error_count + 1
else
	plan_org = removeSlashes(plan_org)
	if not isnumeric(plan_org) then
	po_error = "Plan & Organize grade must be numeric."
	error_count = error_count + 1
	elseif cint(plan_org) > 100 then
	po_error = "Plan & Organize grade may not be over 100."
	error_count = error_count + 1
	elseif cint(plan_org) < 0 then
	po_error = "Plan & Organize grade may not be under 0."
	error_count = error_count + 1
	elseif cint(plan_org) >=0 and cint(plan_org) <= 100 then
		for i = 1 to len(plan_org)
			if mid(plan_org,i,1) = "." then
				error_count = error_count + 1
				po_error = "Plan & Organize grade may not contain decimals."
			end if
		next
	end if
end if

if org_of_mat = "" then
om_error = "Organization of Materials grade is required."
error_count = error_count + 1
else
	org_of_mat = removeSlashes(org_of_mat)
	if not isnumeric(org_of_mat) then
	om_error = "Organization of Materials grade must be numeric."
	error_count = error_count + 1
	elseif cint(org_of_mat) > 100 then
	om_error = "Organization of Materials grade may not be over 100."
	error_count = error_count + 1
	elseif cint(org_of_mat) < 0 then
	om_error = "Organization of Materials grade may not be under 0."
	error_count = error_count + 1
	elseif cint(org_of_mat) >=0 and cint(org_of_mat) <= 100 then
		for i = 1 to len(org_of_mat)
			if mid(org_of_mat,i,1) = "." then
				error_count = error_count + 1
				om_error = "Organization of Materials grade may not contain decimals."
			end if
		next
	end if
end if

if monitor = "" then
mon_error = "Monitor grade is required."
error_count = error_count + 1
else
	monitor = removeSlashes(monitor)
	if not isnumeric(monitor) then
	mon_error = "Monitor grade must be numeric."
	error_count = error_count + 1
	elseif cint(monitor) > 100 then
	mon_error = "Monitor grade may not be over 100."
	error_count = error_count + 1
	elseif cint(monitor) < 0 then
	mon_error = "Monitor grade may not be under 0."
	error_count = error_count + 1
	elseif cint(monitor) >=0 and cint(monitor) <= 100 then
		for i = 1 to len(monitor)
			if mid(monitor,i,1) = "." then
				error_count = error_count + 1
				mon_error = "Monitor grade may not contain decimals."
			end if
		next
	end if
end if

if inhibit = "" then
inhibit_error = "Inhibit grade is required."
error_count = error_count + 1
else
	inhibit = removeSlashes(inhibit)
	if not isnumeric(inhibit) then
	inhibit_error = "Inhibit grade must be numeric."
	error_count = error_count + 1
	elseif cint(inhibit) > 100 then
	inhibit_error = "Inhibit grade may not be over 100."
	error_count = error_count + 1
	elseif cint(inhibit) < 0 then
	inhibit_error = "Inhibit grade may not be under 0."
	error_count = error_count + 1
	elseif cint(inhibit) >=0 and cint(inhibit) <= 100 then
		for i = 1 to len(inhibit)
			if mid(inhibit,i,1) = "." then
				error_count = error_count + 1
				inhibit_error = "Inhibit grade may not contain decimals."
			end if
		next
	end if
end if

if shift = "" then
shift_error = "Shift grade is required."
error_count = error_count + 1
else
	shift = removeSlashes(shift)
	if not isnumeric(shift) then
	shift_error = "Shift grade must be numeric."
	error_count = error_count + 1
	elseif cint(shift) > 100 then
	shift_error = "Shift grade may not be over 100."
	error_count = error_count + 1
	elseif cint(shift) < 0 then
	shift_error = "Shift grade may not be under 0."
	error_count = error_count + 1
	elseif cint(shift) >=0 and cint(shift) <= 100 then
		for i = 1 to len(shift)
			if mid(shift,i,1) = "." then
				error_count = error_count + 1
				shift_error = "Shift grade may not contain decimals."
			end if
		next
	end if
end if

if emotion = "" then
emotion_error = "Emotional Control grade is required."
error_count = error_count + 1
else
	emotion = removeSlashes(emotion)
	if not isnumeric(emotion) then
	emotion_error = "Emotional Control grade must be numeric."
	error_count = error_count + 1
	elseif cint(emotion) > 100 then
	emotion_error = "Emotional Control grade may not be over 100."
	error_count = error_count + 1
	elseif cint(emotion) < 0 then
	emotion_error = "Emotional Control grade may not be under 0."
	error_count = error_count + 1
	elseif cint(emotion) >=0 and cint(emotion) <= 100 then
		for i = 1 to len(emotion)
			if mid(emotion,i,1) = "." then
				error_count = error_count + 1
				emotion_error = "Emotional Control grade may not contain decimals."
			end if
		next
	end if
end if

end if

if error_count = 0 and tokenvalue = 2 then
call pass2
else
%>
			<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2><% =name %> - Add Grades</h2>
				</div>
				<div class="block_content">
					<form name="addgrades" action="grades.asp" method="post">
						<p>
							<label>Inhibit Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="inhibit" name="inhibit" value="<% =cstr(inhibit) %>"/>
							<font color='red'> <% =inhibit_error %> </font>
						</p>
						<p>
							<label>Shift Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="shift" name="shift" value="<% =cstr(shift) %>"/>
							<font color='red'> <% =shift_error %> </font>
						</p>
						<p>
							<label>Emotional Control Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="emotion" name="emotion" value="<% =cstr(emotion) %>"/>
							<font color='red'> <% =emotion_error %> </font>
						</p>
						<p>
							<label>Initiate Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="initiate" name="initiate" value="<% =cstr(initiate) %>"/>
							<font color='red'> <% =ini_error %> </font>
						</p>
						<p>
							<label>Working Memory Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="work_mem" name="work_mem" value="<% =cstr(work_mem) %>"/>
							<font color='red'> <% =wm_error %> </font>
						</p>
						<p>
							<label>Plan & Organize Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="plan_org" name="plan_org" value="<% =cstr(plan_org) %>"/>
							<font color='red'> <% =po_error %> </font>
						</p>	
						<p>
							<label>Organization of Materials Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="org_of_mat" name="org_of_mat" value="<% =cstr(org_of_mat) %>"/>
							<font color='red'> <% =om_error %> </font>
						</p>
						<p>
							<label>Monitor Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="monitor" name="monitor" value="<% =cstr(monitor) %>"/>
							<font color='red'> <% =mon_error %> </font>
						</p>	
							<label>Date Taken:</label> 
							<input type="text" class="text date_picker" id="date_taken" name="date_taken" maxlength="10" value="<% if date_taken = "" then %><% =date() %><% else %><% =cstr(date_taken) %><% end if %>"/>
							<font color='red'> <% =dt_error %> </font>
							<p>
						<hr/>
						<p>
							<input type="submit" class="submit long" value="Submit" />
							<input type="reset" class="submit mid" value="Reset"/>
							<input type="hidden" id="token" name="token" value="2">
							<input type="hidden" name="sid" id="sid" value="<% = cstr(sid) %>">
							<input type="hidden" id="level" name="level" value="<% = cstr(level) %>">
							<input type="hidden" id="name" name="name" value="<% = cstr(name) %>">
						</p>
					</form>
<%
end if
end sub
sub pass2

date_taken = request.form("date_taken")
if mid(date_taken,2,1) = "/" then date_taken = "0" + date_taken

if mid(date_taken,5,1) = "/" then date_taken = mid(date_taken,1,3) + "0" + mid(date_taken,4,6)
dt = convertDate(date_taken)
studentid = request.form("sid")

set cn=Server.CreateObject("ADODB.Connection")
cn.open "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

	SQLString="INSERT INTO grades "
    SQLString=SQLString+"(student_id,date_taken,level,inhibit,shift,emotion,initiate,work_mem,plan_org,org_of_materials,monitor) VALUES ("
    SQLString=SQLString +cstr(studentid) +","
    SQLString=SQLString+ chr(39)+cstr(dt) +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("level") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("inhibit") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("shift") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("emotion") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("initiate") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("work_mem") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("plan_org") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("org_of_mat") +chr(39)+","
    SQLString=SQLString+ chr(39)+request.form("monitor") +chr(39)+")"

    cn.execute SQLString
%>
			<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>Alert</h2>
				</div>
				<div class="block_content">
					<form name="success" action="" method="post">	
						<div class="message success"><p>Grades were successfully added to the database!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="javascript:window.location.href='../view/student.asp?token=3&sid=<% =cstr(studentid) %>'" />
					</p>
					</form>
<%
end sub

tokenvalue=request.form("token")
if tokenvalue="" then
tokenvalue=request.querystring("token")
end if
select case tokenvalue
case "1"
	call pass1
case "2"
  call pass1
end select
%>
				</div>
				<div class="bendl"></div>
				<div class="bendr"></div>	
			</div>
			<br><div id="footer">
				<center>
					<img src="../images/footer.png">
					<p><b>The Monarch School</b></p>
				</center>
			</div>
		</div>
	</div>
	<!--[if IE]><script type="text/javascript" src="../js/excanvas.js"></script><![endif]-->	
	<script type="text/javascript" src="../js/jquery.js"></script>
	<script type="text/javascript" src="../js/jquery.img.preload.js"></script>
	<script type="text/javascript" src="../js/jquery.date_input.pack.js"></script>
	<script type="text/javascript" src="../js/facebox.js"></script>
	<script type="text/javascript" src="../js/jquery.select_skin.js"></script>
	<script type="text/javascript" src="../js/jquery.pngfix.js"></script>
	<script type="text/javascript" src="../js/custom.js"></script>	
</body>
</html>