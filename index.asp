<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<html lang="en">
<head>
<title>The Monarch School</title>
<meta http-equiv="x-ua-compatible" content="IE=8">
<meta charset="utf-8">
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<style type="text/css" media="all">
@import url("css/style.css");
@import url("css/facebox.css");
@import url("css/visualize.css");
</style>
<!--[if lt IE 8]><style type="text/css" media="all">@import url("css/ie.css");</style><![endif]-->

<script language="javascript" src="js/AjaxRequest.js"></script>
<script language="javascript">
function go(col) {
	AjaxRequest.submit(fred,
  	{
    	'url':'js/monarch_ajax.asp'
  		,'onSuccess':function(req){ q_out.innerHTML=req.responseText; }
  		,'onError':function(req){ q_out.innerHTML= req.responseText; }
  	}
		);
	}
</script>
</head>

<body>
	<div id="hld">
		<div class="wrapper">
			<div id="header">
				<div class="hdrl"></div>
				<div class="hdrr"></div>
				<h1><a href="index.asp">The Monarch School</a></h1>
					<ul id="nav">
					<li><a href="index.asp">Dashboard</a></li>
					<li><a href="add/student.asp">Add Student</a></li>
					<li><a href="reports.asp">Reports</a></li>		
					<li><a href="help.asp">Help</a></li>
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
function oneDecimal(someNumber)
	oneDecimal = cstr(formatnumber(someNumber,1))
end function

set rs=Server.CreateObject("ADODB.Recordset")
SQLString="SELECT * FROM grades"
s_pop=0
s_inhi_sum=0
s_shif_sum=0
s_emot_sum=0
s_init_sum=0
s_work_sum=0
s_plan_sum=0
s_mat_sum=0
s_mon_sum=0

n_pop=0
n_inhi_sum=0
n_shif_sum=0
n_emot_sum=0
n_init_sum=0
n_work_sum=0
n_plan_sum=0
n_mat_sum=0
n_mon_sum=0

a_pop=0
a_inhi_sum=0
a_shif_sum=0
a_emot_sum=0
a_init_sum=0
a_work_sum=0
a_plan_sum=0
a_mat_sum=0
a_mon_sum=0

c_pop=0
c_inhi_sum=0
c_shif_sum=0
c_emot_sum=0
c_init_sum=0
c_work_sum=0
c_plan_sum=0
c_mat_sum=0
c_mon_sum=0

v_pop=0
v_inhi_sum=0
v_shif_sum=0
v_emot_sum=0
v_init_sum=0
v_work_sum=0
v_plan_sum=0
v_mat_sum=0
v_mon_sum=0

rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
    s_pop=s_pop+1
	s_inhi_sum=s_inhi_sum+rs("inhibit")
  	s_shif_sum=s_shif_sum+rs("shift")
   	s_emot_sum=s_emot_sum+rs("emotion")
   	s_init_sum=s_init_sum+rs("initiate")
  	s_work_sum=s_work_sum+rs("work_mem")
   	s_plan_sum=s_plan_sum+rs("plan_org")
   	s_mat_sum=s_mat_sum+rs("org_of_materials")
   	s_mon_sum=s_mon_sum+rs("monitor")

    		if rs("level") = "Novice" Then
    			n_pop=n_pop+1
 	  			n_inhi_sum=n_inhi_sum+rs("inhibit")
  				n_shif_sum=n_shif_sum+rs("shift")
   				n_emot_sum=n_emot_sum+rs("emotion")
    			n_init_sum=n_init_sum+rs("initiate")
    			n_work_sum=n_work_sum+rs("work_mem")
    			n_plan_sum=n_plan_sum+rs("plan_org")
    			n_mat_sum=n_mat_sum+rs("org_of_materials")
    			n_mon_sum=n_mon_sum+rs("monitor")

    		elseif rs("level") = "Apprentice" Then
    			a_pop=a_pop+1
 	  			a_inhi_sum=a_inhi_sum+rs("inhibit")
  				a_shif_sum=a_shif_sum+rs("shift")
   				a_emot_sum=a_emot_sum+rs("emotion")
    			a_init_sum=a_init_sum+rs("initiate")
    			a_work_sum=a_work_sum+rs("work_mem")
    			a_plan_sum=a_plan_sum+rs("plan_org")
    			a_mat_sum=a_mat_sum+rs("org_of_materials")
    			a_mon_sum=a_mon_sum+rs("monitor")
  
    		elseif rs("level") = "Challenger" Then
    			c_pop=c_pop+1
    			c_inhi_sum=c_inhi_sum+rs("inhibit")
  				c_shif_sum=c_shif_sum+rs("shift")
   				c_emot_sum=c_emot_sum+rs("emotion")
    			c_init_sum=c_init_sum+rs("initiate")
    			c_work_sum=c_work_sum+rs("work_mem")
    			c_plan_sum=c_plan_sum+rs("plan_org")
    			c_mat_sum=c_mat_sum+rs("org_of_materials")
    			c_mon_sum=c_mon_sum+rs("monitor")
	
    		elseif rs("level") = "Voyager" Then
    			v_pop=v_pop+1
    			v_inhi_sum=v_inhi_sum+rs("inhibit")
  				v_shif_sum=v_shif_sum+rs("shift")
   				v_emot_sum=v_emot_sum+rs("emotion")
    			v_init_sum=v_init_sum+rs("initiate")
    			v_work_sum=v_work_sum+rs("work_mem")
    			v_plan_sum=v_plan_sum+rs("plan_org")
    			v_mat_sum=v_mat_sum+rs("org_of_materials")
    			v_mon_sum=v_mon_sum+rs("monitor")

    		End If
			rs.movenext
wend
rs.close
set rs=nothing

if s_pop <> 0 then
s_inhi_avg=s_inhi_sum/s_pop
s_shif_avg=s_shif_sum/s_pop
s_emot_avg=s_emot_sum/s_pop
s_init_avg=s_init_sum/s_pop
s_work_avg=s_work_sum/s_pop
s_plan_avg=s_plan_sum/s_pop
s_mat_avg=s_mat_sum/s_pop
s_mon_avg=s_mon_sum/s_pop
end if

if n_pop <> 0 then
n_inhi_avg=n_inhi_sum/n_pop
n_shif_avg=n_shif_sum/n_pop
n_emot_avg=n_emot_sum/n_pop
n_init_avg=n_init_sum/n_pop
n_work_avg=n_work_sum/n_pop
n_plan_avg=n_plan_sum/n_pop
n_mat_avg=n_mat_sum/n_pop
n_mon_avg=n_mon_sum/n_pop
end if

if a_pop <> 0 then
a_inhi_avg=a_inhi_sum/a_pop
a_shif_avg=a_shif_sum/a_pop
a_emot_avg=a_emot_sum/a_pop
a_init_avg=a_init_sum/a_pop
a_work_avg=a_work_sum/a_pop
a_plan_avg=a_plan_sum/a_pop
a_mat_avg=a_mat_sum/a_pop
a_mon_avg=a_mon_sum/a_pop
end if

if c_pop <> 0 then
c_inhi_avg=c_inhi_sum/c_pop
c_shif_avg=c_shif_sum/c_pop
c_emot_avg=c_emot_sum/c_pop
c_init_avg=c_init_sum/c_pop
c_work_avg=c_work_sum/c_pop
c_plan_avg=c_plan_sum/c_pop
c_mat_avg=c_mat_sum/c_pop
c_mon_avg=c_mon_sum/c_pop
end if

if v_pop <> 0 then
v_inhi_avg=v_inhi_sum/v_pop
v_shif_avg=v_shif_sum/v_pop
v_emot_avg=v_emot_sum/v_pop
v_init_avg=v_init_sum/v_pop
v_work_avg=v_work_sum/v_pop
v_plan_avg=v_plan_sum/v_pop
v_mat_avg=v_mat_sum/v_pop
v_mon_avg=v_mon_sum/v_pop
end if

set rs2=Server.CreateObject("ADODB.Recordset")
SQLString2="SELECT * FROM student"
rs2.open SQLString2,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

student_count=0
while NOT rs2.eof
	student_count=student_count+1
	rs2.movenext
wend
rs2.close
set rs2=nothing
%>			
			<div class="block">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
					<h2>Statistics - Currently <b><%= cstr(student_count) %></b> Students Enrolled</h2>
						<ul class="tabs">
						<li><a href="#novice">Novice</a></li>
						<li><a href="#apprentice">Apprentice</a></li>
						<li><a href="#challenger">Challenger</a></li>
						<li><a href="#voyager">Voyager</a></li>
						</ul>
				</div>

				<div class="block_content tab_content" id="novice">
					<table class="stats" rel="line" cellpadding="0" cellspacing="0" width="100%">
						<thead>
								<tr>
								<td>&nbsp;</td>
								<th scope="col">Inhibit</th>
								<th scope="col">Shift</th>
								<th scope="col">Emotional&nbsp;Control</th>
								<th scope="col">Initiate</th>
								<th scope="col">Working&nbsp;Memory</th>
								<th scope="col">Plan/Organization</th>
								<th scope="col">Org&nbsp;of&nbsp;Materials</th>
								<th scope="col">Monitor</th>
								</tr>
						</thead>
						<tbody>
							<tr>
								<th>Novice Average</th>		
								<td><%= oneDecimal(n_inhi_avg) %></td>
								<td><%= oneDecimal(n_shif_avg) %></td>
								<td><%= oneDecimal(n_emot_avg) %></td>						
								<td><%= oneDecimal(n_init_avg) %></td>
								<td><%= oneDecimal(n_work_avg) %></td>
								<td><%= oneDecimal(n_plan_avg) %></td>
								<td><%= oneDecimal(n_mat_avg) %></td>
								<td><%= oneDecimal(n_mon_avg) %></td>
							</tr>
							<tr>
								<th>School Average</th>
								<td><%= oneDecimal(s_inhi_avg) %></td>
								<td><%= oneDecimal(s_shif_avg) %></td>
								<td><%= oneDecimal(s_emot_avg) %></td>
								<td><%= oneDecimal(s_init_avg) %></td>
								<td><%= oneDecimal(s_work_avg) %></td>
								<td><%= oneDecimal(s_plan_avg) %></td>
								<td><%= oneDecimal(s_mat_avg) %></td>
								<td><%= oneDecimal(s_mon_avg) %></td>
							</tr>
						</tbody>
					</table>	
				</div>

				<div class="block_content tab_content" id="apprentice">
					<table class="stats" rel="line" cellpadding="0" cellspacing="0" width="100%">
						<thead>
								<tr>
								<td>&nbsp;</td>
								<th scope="col">Inhibit</th>
								<th scope="col">Shift</th>
								<th scope="col">Emotional&nbsp;Control</th>
								<th scope="col">Initiate</th>
								<th scope="col">Working&nbsp;Memory</th>
								<th scope="col">Plan/Organization</th>
								<th scope="col">Org&nbsp;of&nbsp;Materials</th>
								<th scope="col">Monitor</th>
								</tr>
						</thead>
						<tbody>
							<tr>
								<th>Apprentice Average</th>			
								<td><%= oneDecimal(a_inhi_avg) %></td>
								<td><%= oneDecimal(a_shif_avg) %></td>
								<td><%= oneDecimal(a_emot_avg) %></td>						
								<td><%= oneDecimal(a_init_avg) %></td>
								<td><%= oneDecimal(a_work_avg) %></td>
								<td><%= oneDecimal(a_plan_avg) %></td>
								<td><%= oneDecimal(a_mat_avg) %></td>
								<td><%= oneDecimal(a_mon_avg) %></td>
							</tr>
							<tr>
								<th>School Average</th>
								<td><%= oneDecimal(s_inhi_avg) %></td>
								<td><%= oneDecimal(s_shif_avg) %></td>
								<td><%= oneDecimal(s_emot_avg) %></td>
								<td><%= oneDecimal(s_init_avg) %></td>
								<td><%= oneDecimal(s_work_avg) %></td>
								<td><%= oneDecimal(s_plan_avg) %></td>
								<td><%= oneDecimal(s_mat_avg) %></td>
								<td><%= oneDecimal(s_mon_avg) %></td>
							</tr>
						</tbody>
					</table>
				</div>

				<div class="block_content tab_content" id="challenger">
					<table class="stats" rel="line" cellpadding="0" cellspacing="0" width="100%">
						<thead>
								<tr>
								<td>&nbsp;</td>
								<th scope="col">Inhibit</th>
								<th scope="col">Shift</th>
								<th scope="col">Emotional&nbsp;Control</th>
								<th scope="col">Initiate</th>
								<th scope="col">Working&nbsp;Memory</th>
								<th scope="col">Plan/Organization</th>
								<th scope="col">Org&nbsp;of&nbsp;Materials</th>
								<th scope="col">Monitor</th>
								</tr>
						</thead>
						<tbody>
							<tr>
								<th>Challenger Average</th>	
								<td><%= oneDecimal(c_inhi_avg) %></td>
								<td><%= oneDecimal(c_shif_avg) %></td>
								<td><%= oneDecimal(c_emot_avg) %></td>								
								<td><%= oneDecimal(c_init_avg) %></td>
								<td><%= oneDecimal(c_work_avg) %></td>
								<td><%= oneDecimal(c_plan_avg) %></td>
								<td><%= oneDecimal(c_mat_avg) %></td>
								<td><%= oneDecimal(c_mon_avg) %></td>
							</tr>
							<tr>
								<th>School Average</th>
								<td><%= oneDecimal(s_inhi_avg) %></td>
								<td><%= oneDecimal(s_shif_avg) %></td>
								<td><%= oneDecimal(s_emot_avg) %></td>
								<td><%= oneDecimal(s_init_avg) %></td>
								<td><%= oneDecimal(s_work_avg) %></td>
								<td><%= oneDecimal(s_plan_avg) %></td>
								<td><%= oneDecimal(s_mat_avg) %></td>
								<td><%= oneDecimal(s_mon_avg) %></td>
							</tr>
						</tbody>
					</table>
				</div>
				
				<div class="block_content tab_content" id="voyager">
					<table class="stats" rel="line" cellpadding="0" cellspacing="0" width="100%">
						<thead>
								<tr>
								<td>&nbsp;</td>
								<th scope="col">Inhibit</th>
								<th scope="col">Shift</th>
								<th scope="col">Emotional&nbsp;Control</th>
								<th scope="col">Initiate</th>
								<th scope="col">Working&nbsp;Memory</th>
								<th scope="col">Plan/Organization</th>
								<th scope="col">Org&nbsp;of&nbsp;Materials</th>
								<th scope="col">Monitor</th>
								</tr>
						</thead>
						<tbody>
							<tr>
								<th>Voyager Average</th>	
								<td><%= oneDecimal(v_inhi_avg) %></td>
								<td><%= oneDecimal(v_shif_avg) %></td>
								<td><%= oneDecimal(v_emot_avg) %></td>								
								<td><%= oneDecimal(v_init_avg) %></td>
								<td><%= oneDecimal(v_work_avg) %></td>
								<td><%= oneDecimal(v_plan_avg) %></td>
								<td><%= oneDecimal(v_mat_avg) %></td>
								<td><%= oneDecimal(v_mon_avg) %></td>
							</tr>
							<tr>
								<th>School Average</th>
								<td><%= oneDecimal(s_inhi_avg) %></td>
								<td><%= oneDecimal(s_shif_avg) %></td>
								<td><%= oneDecimal(s_emot_avg) %></td>
								<td><%= oneDecimal(s_init_avg) %></td>
								<td><%= oneDecimal(s_work_avg) %></td>
								<td><%= oneDecimal(s_plan_avg) %></td>
								<td><%= oneDecimal(s_mat_avg) %></td>
								<td><%= oneDecimal(s_mon_avg) %></td>
							</tr>
						</tbody>
					</table>
				</div>
				<div class="bendl"></div>
				<div class="bendr"></div>
			</div>
			
			<div class="block">
				<div class="block_head">
					<div class="bheadl"></div>
					<div class="bheadr"></div>
						<h2>Results</h2>
						<form action="" method="post" name="fred">
						<input type="text" name="LastName" onkeyup="go(1)" onfocus="if (this.value==this.defaultValue) this.value = ''"
						onblur="if (this.value=='') this.value = this.defaultValue" class="text" value="Student's Last Name" />
						</form>
				</div>
				<div class="block_content">
					<div id="q_out"></div>
				</div>	
				<div class="bendl"></div>
				<div class="bendr"></div>
			</div>
			
			<div id="footer">
				<center>
					<img src="images/footer.png">
					<b><p align='middle'>The Monarch School</b></p>
				</center>
			</div>
		</div>
	</div>
	
	<!--[if IE]><script type="text/javascript" src="js/excanvas.js"></script><![endif]-->	
	<script type="text/javascript" src="js/jquery.js"></script>
	<script type="text/javascript" src="js/jquery.img.preload.js"></script>
	<script type="text/javascript" src="js/jquery.date_input.pack.js"></script>
	<script type="text/javascript" src="js/facebox.js"></script>
	<script type="text/javascript" src="js/jquery.visualize.js"></script>
	<script type="text/javascript" src="js/jquery.visualize.tooltip.js"></script>
	<script type="text/javascript" src="js/jquery.select_skin.js"></script>
	<script type="text/javascript" src="js/jquery.pngfix.js"></script>
	<script type="text/javascript" src="js/custom.js"></script>
</body>
</html>