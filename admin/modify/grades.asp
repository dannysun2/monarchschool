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
dim rs
dim tokenvalue

Function removeSlashes(gradeString)
	removeSlashes = Replace(gradeString,chr(47),"*")
	removeSlashes = Replace(removeSlashes,chr(92),"*")
end function

sub pass1
  
sid=request.querystring("sid")
if sid = "" then sid = request.form("sid")
gid=request.querystring("gid")
if gid = "" then gid = request.form("gid")
ini_error=""
wm_error=""
po_error=""
om_error=""
mon_error=""
inhibit_error=""
shift_error=""
emotion_error=""
error_count=0
SQLString="SELECT * FROM student WHERE student_id="+cstr(sid)

set rs=Server.CreateObject("ADODB.Recordset")

rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
name = cstr(rs("stu_fname")) + " " + cstr(rs("stu_lname"))
rs.movenext
wend
rs.close
set rs=nothing

  if tokenvalue = 1 then
  set rs=Server.CreateObject("ADODB.Recordset")
  SQLString2="SELECT *,DATE_FORMAT(date_taken,'%m/%d/%Y') AS dt_taken FROM grades WHERE student_id=" +cstr(sid) +" AND grades_id=" +cstr(gid)
  rs.open SQLString2,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

   c=0
   while NOT rs.EOF
	inhibit_val = rs("inhibit")
	shift_val = rs("shift")
	emotion_val = rs("emotion")
	initiate_val = rs("initiate")
	workmem_val = rs("work_mem")
	planorg_val = rs("plan_org")
	monitor_val = rs("monitor")
	orgofmat_val = rs("org_of_materials")
	dt_taken = cstr(rs("dt_taken"))
      c=c+1
      rs.movenext
   wend
   rs.close
   set rs=nothing

elseif tokenvalue = 2 then
	inhibit_val = cstr(request.form("inhibit"))
	shift_val = cstr(request.form("shift"))
	emotion_val = cstr(request.form("emotion"))
	initiate_val = cstr(request.form("initiate"))
	workmem_val = cstr(request.form("work_mem"))
	planorg_val = cstr(request.form("plan_org"))
	monitor_val = cstr(request.form("monitor"))
	orgofmat_val = cstr(request.form("org_of_mat"))
	
	
	if initiate_val = "" then
ini_error = "Initiate grade is required."
error_count = error_count + 1
else
	initiate_val = removeSlashes(initiate_val)
	if not isnumeric(initiate_val) then
	ini_error = "Initiate grade must be numeric."
	error_count = error_count + 1
	elseif cint(initiate_val) > 100 then
	ini_error = "Initiate grade may not be over 100."
	error_count = error_count + 1
	elseif cint(initiate_val) < 0 then
	ini_error = "Initiate grade may not be under 0."
	error_count = error_count + 1
	elseif cint(initiate_val) >=0 and cint(initiate_val) <= 100 then
		for i = 1 to len(initiate_val)
			if mid(initiate_val,i,1) = "." then
				error_count = error_count + 1
				ini_error = "Initiate grade may not contain decimals."
			end if
		next
	end if
end if

if workmem_val = "" then
wm_error = "Working Memory grade is required."
error_count = error_count + 1
else
	workmem_val = removeSlashes(workmem_val)
	if not isnumeric(workmem_val) then
	wm_error = "Working Memory grade must be numeric."
	error_count = error_count + 1
	elseif cint(workmem_val) > 100 then
	wm_error = "Working Memory grade may not be over 100."
	error_count = error_count + 1
	elseif cint(workmem_val) < 0 then
	wm_error = "Working Memory grade may not be under 0."
	error_count = error_count + 1
	elseif cint(workmem_val) >=0 and cint(workmem_val) <= 100 then
		for i = 1 to len(workmem_val)
			if mid(workmem_val,i,1) = "." then
				error_count = error_count + 1
				wm_error = "Working Memory grade may not contain decimals."
			end if
		next
	end if
end if

if planorg_val = "" then
po_error = "Plan & Organize grade is required."
error_count = error_count + 1
else
	planorg_val = removeSlashes(planorg_val)
	if not isnumeric(planorg_val) then
	po_error = "Plan & Organize grade must be numeric."
	error_count = error_count + 1
	elseif cint(planorg_val) > 100 then
	po_error = "Plan & Organize grade may not be over 100."
	error_count = error_count + 1
	elseif cint(planorg_val) < 0 then
	po_error = "Plan & Organize grade may not be under 0."
	error_count = error_count + 1
	elseif cint(planorg_val) >=0 and cint(planorg_val) <= 100 then
		for i = 1 to len(planorg_val)
			if mid(planorg_val,i,1) = "." then
				error_count = error_count + 1
				po_error = "Plan & Organize grade may not contain decimals."
			end if
		next
	end if
end if

if orgofmat_val = "" then
om_error = "Organization of Materials grade is required."
error_count = error_count + 1
else
	orgofmat_val = removeSlashes(orgofmat_val)
	if not isnumeric(orgofmat_val) then
	om_error = "Organization of Materials grade must be numeric."
	error_count = error_count + 1
	elseif cint(orgofmat_val) > 100 then
	om_error = "Organization of Materials grade may not be over 100."
	error_count = error_count + 1
	elseif cint(orgofmat_val) < 0 then
	om_error = "Organization of Materials grade may not be under 0."
	error_count = error_count + 1
	elseif cint(orgofmat_val) >=0 and cint(orgofmat_val) <= 100 then
		for i = 1 to len(orgofmat_val)
			if mid(orgofmat_val,i,1) = "." then
				error_count = error_count + 1
				om_error = "Organization of Materials grade may not contain decimals."
			end if
		next
	end if
end if

if monitor_val = "" then
mon_error = "Monitor grade is required."
error_count = error_count + 1
else
	monitor_val = removeSlashes(monitor_val)
	if not isnumeric(monitor_val) then
	mon_error = "Monitor grade must be numeric."
	error_count = error_count + 1
	elseif cint(monitor_val) > 100 then
	mon_error = "Monitor grade may not be over 100."
	error_count = error_count + 1
	elseif cint(monitor_val) < 0 then
	mon_error = "Monitor grade may not be under 0."
	error_count = error_count + 1
	elseif cint(monitor_val) >=0 and cint(monitor_val) <= 100 then
		for i = 1 to len(monitor_val)
			if mid(monitor_val,i,1) = "." then
				error_count = error_count + 1
				mon_error = "Monitor grade may not contain decimals."
			end if
		next
	end if
end if

if inhibit_val = "" then
inhibit_error = "Inhibit grade is required."
error_count = error_count + 1
else
	inhibit_val = removeSlashes(inhibit_val)
	if not isnumeric(inhibit_val) then
	inhibit_error = "Inhibit grade must be numeric."
	error_count = error_count + 1
	elseif cint(inhibit_val) > 100 then
	inhibit_error = "Inhibit grade may not be over 100."
	error_count = error_count + 1
	elseif cint(inhibit_val) < 0 then
	inhibit_error = "Inhibit grade may not be under 0."
	error_count = error_count + 1
	elseif cint(inhibit_val) >=0 and cint(inhibit_val) <= 100 then
		for i = 1 to len(inhibit_val)
			if mid(inhibit_val,i,1) = "." then
				error_count = error_count + 1
				inhibit_error = "Inhibit grade may not contain decimals."
			end if
		next
	end if
end if

if shift_val = "" then
shift_error = "Shift grade is required."
error_count = error_count + 1
else
	shift_val = removeSlashes(shift_val)
	if not isnumeric(shift_val) then
	shift_error = "Shift grade must be numeric."
	error_count = error_count + 1
	elseif cint(shift_val) > 100 then
	shift_error = "Shift grade may not be over 100."
	error_count = error_count + 1
	elseif cint(shift_val) < 0 then
	shift_error = "Shift grade may not be under 0."
	error_count = error_count + 1
	elseif cint(shift_val) >=0 and cint(shift_val) <= 100 then
		for i = 1 to len(shift_val)
			if mid(shift_val,i,1) = "." then
				error_count = error_count + 1
				shift_error = "Shift grade may not contain decimals."
			end if
		next
	end if
end if

if emotion_val = "" then
emotion_error = "Emotional Control grade is required."
error_count = error_count + 1
else
	emotion_val = removeSlashes(emotion_val)
	if not isnumeric(emotion_val) then
	emotion_error = "Emotional Control grade must be numeric."
	error_count = error_count + 1
	elseif cint(emotion_val) > 100 then
	emotion_error = "Emotional Control grade may not be over 100."
	error_count = error_count + 1
	elseif cint(emotion_val) < 0 then
	emotion_error = "Emotional Control grade may not be under 0."
	error_count = error_count + 1
	elseif cint(emotion_val) >=0 and cint(emotion_val) <= 100 then
		for i = 1 to len(emotion_val)
			if mid(emotion_val,i,1) = "." then
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
					<h2><% =name %> - Test <% =dt_taken %></h2>
				</div>
				<div class="block_content">
					<form name="modifygrades" action="grades.asp" method="post">
						<p>
							<label>Inhibit Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="inhibit" name="inhibit" value="<% =cstr(inhibit_val) %>"/>
							<font color='red'> <% =inhibit_error %> </font>
						</p>
						<p>
							<label>Shift Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="shift" name="shift" value="<% =cstr(shift_val) %>"/>
							<font color='red'> <% =shift_error %> </font>
						</p>
						<p>
							<label>Emotional Control Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="emotion" name="emotion" value="<% =cstr(emotion_val) %>"/>
							<font color='red'> <% =emotion_error %> </font>
						</p>

						<p>
							<label>Initiate Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="initiate" name="initiate" value="<% =cstr(initiate_val) %>"/>
							<font color='red'> <% =ini_error %> </font>
						</p>
						<p>
							<label>Working Memory Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="work_mem" name="work_mem" value="<% =cstr(workmem_val) %>"/>
							<font color='red'> <% =wm_error %> </font>
						</p>
						<p>
							<label>Plan & Organize Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="plan_org" name="plan_org" value="<% =cstr(planorg_val) %>"/>
							<font color='red'> <% =po_error %> </font>
						</p>
						<p>
							<label>Organization of Materials Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="org_of_mat" name="org_of_mat" value="<% =cstr(orgofmat_val) %>"/>
							<font color='red'> <% =om_error %> </font>
						</p>
						<p>
							<label>Monitor Grade:</label><br />
							<input type="text" class="text big" maxlength="3" id="monitor" name="monitor" value="<% =cstr(monitor_val) %>" />
							<font color='red'> <% =mon_error %> </font>
						</p>	
						<hr/>
						<p>
							<input type="submit" class="submit long" value="Submit" />
							<input type="reset" class="submit mid" value="Reset"/>
							<input type="hidden" name="token" value="2">
							<input type="hidden" name="sid" value="<% = cstr(sid) %>"><p>
							<input type="hidden" name="gid" value="<% = cstr(gid) %>"><p>
						</p>
					</form>
<%
end if
end sub
sub pass2
studentid = cstr(request.form("sid"))
set cn=Server.CreateObject("ADODB.Connection")
cn.open "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

SQLString="UPDATE grades SET "
SQLString=SQLString+ " inhibit = "+ chr(39) + request.form("inhibit")  + chr(39)  
SQLString=SQLString+ ", shift = "+ chr(39) + request.form("shift")  + chr(39)  
SQLString=SQLString+ ", emotion = "+ chr(39) + request.form("emotion")  + chr(39)     
SQLString=SQLString+ ", initiate = "+ chr(39) + request.form("initiate")  + chr(39)   
SQLString=SQLString+ ", work_mem = "+ chr(39) + request.form("work_mem")  + chr(39) 
SQLString=SQLString+ ", plan_org = "+ chr(39) + request.form("plan_org")  + chr(39)
SQLString=SQLString+ ", monitor = "+ chr(39) + request.form("monitor")  + chr(39) 
SQLString=SQLString+ ", org_of_materials = "+ chr(39) + request.form("org_of_mat")  + chr(39) 
SQLString=SQLString+ " WHERE student_id=" +cstr(request.form("sid")) +" AND grades_id=" +cstr(request.form("gid"))

cn.execute SQLString,numa


%>
<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>Alert</h2>
				</div>
				<div class="block_content">
					<form name="success" action="" method="post">	
					<div class="message success"><p>Grades were successfully modified!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="javascript:window.location.href='../view/student.asp?token=3&sid=<% =cstr(studentid) %>'" />
					</p>
					</form>
<%


cn.close
set cn=nothing
end sub

sub passerror
     response.write "<p>INVALID TOKEN VALUE. token="+cstr(tokernvalue)
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