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
</script>
<script type="text/javascript" src="../js/placeholder.min.js"></script>
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

function convertDate(dateString)
convertDate = mid(dateString,7,4) + "-" + mid(dateString,1,2) + "-" + mid(dateString,4,2) + " 05:00:00"
end function

sub pass1

error_count = 0

if tokenvalue = 2 then

gender = request.form("gender")
level = request.form("level")
fname = request.form("f_name")
mname = request.form("m_name")
lname = request.form("l_name")
level = request.form("level")

dob = request.form("stu_dob")

ed = request.form("enrolldate")

disability = request.form("diagnosis")
fname_error=""
mname_error=""
lname_error=""
ed_error = ""
dob_error = ""
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
				if mname_error = "" then mname_error = "Invalid characters found in first name: "
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
				if lname_error = "" then lname_error = "Invalid characters found in first name: "
                lname_error = lname_error + mid(invalidListName,i,1) + ", "
				error_count = error_count + 1
                
            end if 
        next 
end if

if dob = "" then
dob_error = "Student must have a Date of Birth."
error_count = error_count + 1
end if

if ed = "" then
ed_error = "Student must have an Enrollment Date."
error_count = error_count + 1
end if

end if

if tokenvalue = 2 and error_count=0 then
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
					<h2>Add Student</h2>
				</div>
				<div class="block_content">
					<form name="addstudent" action="student.asp" method="post">
						<p>
							<label>First Name:</label><br />
							<input type="text" class="text big" id="f_name" name="f_name" value="<% =fname %>"/>
							<font color='red'> <% =fname_error %> </font>
						</p>
						<p>
							<label>Middle Name:</label><br />
							<input type="text" class="text big" id="m_name" name="m_name" value="<% =mname %>"/>
							<font color='red'> <% =mname_error %> </font>
						</p>
						<p>
							<label>Last Name:</label><br />
							<input type="text" class="text big" id="l_name" name="l_name" value="<% =lname %>"/>
							<font color='red'> <% =lname_error %> </font>
						</p>	
						<p>
							<p><label>Select Gender:</label> <br />
							<select name="gender" id="gender" class="styled" width="100%">
								<option value="M">Male</option>
								<option value="F" <% if gender = "F" then %> <%="selected"%> <% end if %>>Female</option>					
							</select></p>
							<p><label>Select Level:</label> <br />
							<select name="level" id="level" class="styled" width="100%">
								<option value="Novice" <% if level = "Novice" then %> <%="selected"%> <% end if %>>Novice</option>
								<option value="Apprentice" <% if level = "Apprentice" then %> <%="selected"%> <% end if %>>Apprentice</option>	
								<option value="Challenger" <% if level = "Challenger" then %> <%="selected"%> <% end if %>>Challenger</option>	
								<option value="Voyager" <% if level = "Voyager" then %> <%="selected"%> <% end if %>>Voyager</option>		
							</select></p>
							
							<p><label>Date of Birth:</label> 
							<input type="text" class="text big" placeholder="MM/DD/YYYY" id="stu_dob" name="stu_dob" value="<% =dob %>"/>
							<font color='red'> <% =dob_error %> </font>
							</p>
							<p><label>Enrollment Date:</label> 
							<input type="text" class="text date_picker" id="enrolldate" name="enrolldate" value="<% =ed %>"/>
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
							</select></p>						
						<hr/>
						<p>
							<input type="submit" class="submit long" value="Submit" />
							<input type="reset" class="submit mid" value="Reset"/>
							<input type="hidden" id="token" name="token" value="2">
						</p>
					</form>
<%
end if
end sub
sub pass2

set cn=Server.CreateObject("ADODB.Connection")
cn.open "gl1181","gl1181","YVT52ddnJ"

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
ed = convertDate(ed)

Insert_string="INSERT INTO student (stu_fname,stu_mname,stu_lname,stu_gender,level,dob_date,stu_diagnosis,ed_date)"
Insert_string=Insert_string+" VALUES ("
Insert_string=Insert_string+chr(39)+cstr(f_name)+chr(39)+"," 
Insert_string=Insert_string+chr(39)+cstr(m_name)+chr(39)+"," 
Insert_string=Insert_string+chr(39)+cstr(l_name)+chr(39)+"," 
Insert_string=Insert_string+chr(39)+cstr(Request.form("gender"))+chr(39)+"," 
Insert_string=Insert_string+chr(39)+cstr(Request.form("level"))+chr(39)+","
Insert_string=Insert_string+chr(39)+cstr(dob)+chr(39)+","
Insert_string=Insert_string+chr(39)+cstr(Request.form("diagnosis"))+chr(39)+"," 
Insert_string=Insert_string+chr(39)+cstr(ed)+chr(39)+")"

cn.execute Insert_string,numa

if Err = 0 and numa = 1 then
%>
<div class="block2">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>Alert</h2>
				</div>
				<div class="block_content">
					<form name="success" action="" method="post">	
					<div class="message success"><p>Student was successfully added to the database!</p></div>
					<p>
					<input Type="button" value="Back" class="submit mid" onClick="history.go(-2);return true;" />
					</p>
					</form>

<%
else
  If cn.Errors.Count > 0 Then
     for i = 0 to cn.Errors.Count - 1
         response.write "<p>"
         etext=ucase(cn.errors(i))
         k=instr(etext,"DUPLICATE")
         if k > 0 then
           response.write "<p>DUPLICATE user_id IN THE login DATABASE!!<br>"
           response.write "userid="+cstr(request.form("user_id")) + " already entered. Click the <b>BACK</b> button to try again"
           exit for
         else
            response.write "<br><b>"+cn.errors(i)+"</b>"
         end if
      next 
   end if 
end if

end sub

on error resume next
tokenvalue=request.form("token")
select case tokenvalue
case ""
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
<script type="text/javascript">
        Placeholder.init({
            classFocus: "normal",
            classBlur:  "placeholder",
            wait: true
        });
    </script>
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