<!DOCTYPE html>
<html lang="en" xml:lang="en" xmlns="http://www.w3.org/1999/xhtml"></html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
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

function convertDate(dateString)
convertDate = mid(dateString,7,4) + "-" + mid(dateString,1,2) + "-" + mid(dateString,4,2) + " 05:00:00"
end function

sub pass1
error_count = 0

if tokenvalue = 1 then
set rs=Server.CreateObject("ADODB.Recordset")
sid=request.querystring("sid")
SQLString="SELECT *,DATE_FORMAT(dob_date,'%m/%d/%Y') AS dobDate,DATE_FORMAT(ed_date,'%m/%d/%Y') AS edDate FROM student WHERE student_id="+cstr(sid)
rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

fname = rs("stu_fname")
mname = rs("stu_mname")
lname = rs("stu_lname")
gender = rs("stu_gender")
level = rs("level")
dob = rs("dobdate")
disability = rs("stu_diagnosis")
enrolldate = rs("eddate")
name = cstr(rs("stu_fname")) + " " + cstr(rs("stu_lname"))

rs.close
set rs=nothing
elseif tokenvalue = 2 then

name = request.form("name")
sid=request.form("sid")
fname = request.form("f_name")
mname = request.form("m_name")
lname = request.form("l_name")
gender = request.form("gender")
level = request.form("level")
dob = request.form("stu_dob")
disability = request.form("diagnosis")
enrolldate = request.form("enrolldate")
fname_error=""
mname_error=""
lname_error=""
invalidListName = ",<.>?;:'@#~]}[{=+)(*&^%$£!`¬| _0123456789" 

if fname = "" then
fname_error = "Student must have a first name.  "
error_count = error_count + 1
else
fname = Replace(fname,chr(92),"*")
fname = Replace(fname,chr(47),"*")
for i = 1 to len(invalidListName) 
            if instr(fname,mid(invalidListName,i,1))>0 then
				if fname_error = "" then fname_error = "Invalid characters found in first name: "
                fname_error = fname_error + mid(invalidListName,i,1) + ", "
				error_count = error_count + 1
                
            end if 
        next 
end if

if mname = "" then
else
mname = Replace(mname,chr(92),"*")
mname = Replace(mname,chr(47),"*")
for i = 1 to len(invalidListName) 
            if instr(mname,mid(invalidListName,i,1))>0 then 
				if mname_error = "" then mname_error = "Invalid characters found in middle name: "
                mname_error = mname_error + mid(invalidListName,i,1) + ", "
				error_count = error_count + 1
                 
            end if 
        next 
end if

if lname = "" then
lname_error = "Student must have a last name.  "
error_count = error_count + 1
else
lname = Replace(lname,chr(92),"*")
lname = Replace(lname,chr(47),"*")
for i = 1 to len(invalidListName) 
            if instr(lname,mid(invalidListName,i,1))>0 then 
				if lname_error = "" then lname_error = "Invalid characters found in last name: "
                lname_error = lname_error + mid(invalidListName,i,1) + ", "
				error_count = error_count + 1
                
            end if 
        next 
end if

if dob = "" then
dob_error = "Student must have a Date of Birth."
error_count = error_count + 1
elseif not isdate(dob) then
dob_error = "Invalid Date of Birth."
error_count = error_count + 1
end if

if enrolldate = "" then
ed_error = "Student must have an Enrollment Date."
error_count = error_count + 1
elseif not isdate(enrolldate) then
ed_error = "Invalid Enrollment Date."
error_count = error_count + 1
end if
end if

if tokenvalue = 2 and error_count = 0 then
call pass2
else
if fname_error <> "" then fname_error = left(fname_error,len(fname_error)-2)
if lname_error <> "" then lname_error = left(lname_error,len(lname_error)-2)
if mname_error <> "" then mname_error = left(mname_error,len(mname_error)-2)
%>
			<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2><% =name %> - Modify Info</h2>
				</div>
				<div class="block_content">
					<form name="modifystudent" action="student.asp" method="post">
						<p>
							<label>First Name:</label><br />
							<input type="text" class="text big" id="f_name" name="f_name" value="<% =cstr(fname) %>" />
							<font color='red'> <% =fname_error %> </font>
						</p>
						<p>
							<label>Middle Name:</label><br />
							<input type="text" class="text big" id="m_name" name="m_name" value="<% =cstr(mname) %>" />
							<font color='red'> <% =mname_error %> </font>
						</p>
						<p>
							<label>Last Name:</label><br />
							<input type="text" class="text big" id="l_name" name="l_name" value="<% =cstr(lname) %>"/>
							<font color='red'> <% =lname_error %> </font>
						</p>	
						<p>
							<label>Select Gender:</label> <br />
							<select name="gender" id="gender" class="styled" width=100%>
								<option value="M">Male</option>
								<option value="F" <% if gender = "F" then %> <%="selected"%> <% end if %>>Female</option>					
							</select>
						</p>
							<p>
								<label>Select Level:</label> <br />
							<select name="level" id="level" class="styled" width="100%">
								<option value="Novice" <% if level = "Novice" then %> <%="selected"%> <% end if %>>Novice</option>
								<option value="Apprentice" <% if level = "Apprentice" then %> <%="selected"%> <% end if %>>Apprentice</option>	
								<option value="Challenger" <% if level = "Challenger" then %> <%="selected"%> <% end if %>>Challenger</option>	
								<option value="Voyager" <% if level = "Voyager" then %> <%="selected"%> <% end if %>>Voyager</option>		
							</select>
						</p>
							<p>
							<label>Date of Birth:</label> 
							<input type="text" placeholder="MM/DD/YYYY" class="text date_picker" id="stu_dob" name="stu_dob" value="<% =cstr(dob) %>"/>
							<font color='red'> <% =dob_error %> </font>
							</p>
							<p><label>Enrollment Date:</label> 
							<input type="text" placeholder="MM/DD/YYYY" class="text date_picker" id="enrolldate" name="enrolldate" value="<% =cstr(enrolldate) %>"/>
							<font color='red'> <% =ed_error %> </font>
						</p>
						<p><label>Select Diagnosis:</label> <br />
							<select name="diagnosis" id="diagnosis" class="styled" width=100%>
								<option value="Asperger Syndrome" <% if disability = "Asperger Syndrome" then %> <%="selected"%> <% end if %>>Asperger Syndrome</option>
								<option value="Anxiety Disorder" <% if disability = "Anxiety Disorder" then %> <%="selected"%> <% end if %>>Anxiety Disorder</option>
								<option value="Autism" <% if disability = "Autism" then %> <%="selected"%> <% end if %>>Autism</option>
								<option value="Bipolar Disorder" <% if disability = "Bipolar Disorder" then %> <%="selected"%> <% end if %>>Bipolar Disorder</option>
								<option value="Tourette Syndrome" <% if disability = "Tourette Syndrome" then %> <%="selected"%> <% end if %>>Tourette Syndrome</option>
								<option value="Traumatic Brain Injury" <% if disability = "Traumatic Brain Injury" then %> <%="selected"%> <% end if %>>Traumatic Brain Injury</option>
								<option value="Other" <% if disability = "Other" then %> <%="selected"%> <% end if %>>Other</option>
							</select>
						</p>						
						<hr/>
						<p>
						<input type="submit" class="submit long" value="Submit" />
						<input type="reset" class="submit mid" value="Reset"/>
						<input type="hidden" id="token" name="token" value="2">
						<input type="hidden" name="sid" value="<% =cstr(sid) %>">
						<input type="hidden" id="name" name="name" value="<% = cstr(name) %>">
						</p>
					</form>
<%
end if
end sub

sub pass2

set cn=Server.CreateObject("ADODB.Connection")
cn.open "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

f_name = request.form("f_name")
f_name = ucase(mid(f_name,1,1)) +mid(f_name,2)

m_name = request.form("m_name")
m_name = ucase(mid(m_name,1,1)) +mid(m_name,2)

l_name = request.form("l_name")
l_name = ucase(mid(l_name,1,1)) +mid(l_name,2)

dob = request.form("stu_dob")
if mid(dob,2,1) = "/" then dob = "0" + dob
if mid(dob,5,1) = "/" then dob = mid(dob,1,3) + "0" + mid(dob,4,6)
if mid(dob,2,1) = "-" then dob = "0" + dob
if mid(dob,5,1) = "-" then dob = mid(dob,1,3) + "0" + mid(dob,4,6)
dob = convertDate(dob)


ed = request.form("enrolldate")
if mid(ed,2,1) = "/" then ed = "0" + ed
if mid(ed,5,1) = "/" then ed = mid(ed,1,3) + "0" + mid(ed,4,6)
if mid(ed,2,1) = "-" then ed = "0" + ed
if mid(ed,5,1) = "-" then ed = mid(ed,1,3) + "0" + mid(ed,4,6)
ed = convertDate(ed)

SQLString="UPDATE student SET "
SQLString=SQLString+ " stu_fname = "+ chr(39) +f_name  + chr(39)   
SQLString=SQLString+ ", stu_lname = "+ chr(39) +l_name  + chr(39) 
SQLString=SQLString+ ", stu_mname = "+ chr(39) +m_name  + chr(39)
SQLString=SQLString+ ", stu_gender = "+ chr(39) + request.form("gender")  + chr(39) 
SQLString=SQLString+ ", level = "+ chr(39) + request.form("level")  + chr(39) 
SQLString=SQLString+ ", dob_date = "+ chr(39) +dob + chr(39)
SQLString=SQLString+ ", stu_diagnosis = "+ chr(39) + request.form("diagnosis")  + chr(39) 
SQLString=SQLString+ ", ed_date = "+ chr(39) +ed + chr(39)   
SQLString=SQLString+ " WHERE student_id=" +cstr(request.form("sid"))

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
					<div class="message success"><p>Student record successfully modified!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="javascript:window.location.href='../index.asp'" />
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
case else
	call passerror
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
          <b><p align='middle'>The Monarch School</b></p>
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
	<script type="text/javascript" src="../js/jquery.placeholder.js?v=2.0.7"></script>
  <script>
   // To test the @id toggling on password inputs in browsers that don’t support changing an input’s @type dynamically (e.g. Firefox 3.6 or IE), uncomment this:
   // $.fn.hide = function() { return this; }
   // Then uncomment the last rule in the <style> element (in the <head>).
   $(function() {
    // Invoke the plugin
    $('input, textarea').placeholder();
    // That’s it, really.
    // Now display a message if the browser supports placeholder natively
    var html;
    if ($.fn.placeholder.input && $.fn.placeholder.textarea) {
     html = '<strong>Your current browser natively supports <code>placeholder</code> for <code>input</code> and <code>textarea</code> elements.</strong> The plugin won’t run in this case, since it’s not needed. If you want to test the plugin, use an older browser ;)';
    } else if ($.fn.placeholder.input) {
     html = '<strong>Your current browser natively supports <code>placeholder</code> for <code>input</code> elements, but not for <code>textarea</code> elements.</strong> The plugin will only do its thang on the <code>textarea</code>s.';
    }
    if (html) {
     $('<p class="note">' + html + '</p>').insertAfter('form');
    }
   });
  </script>
</body>
</html>