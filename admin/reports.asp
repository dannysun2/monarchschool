<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico">
<meta http-equiv="X-UA-Compatible" content="IE=7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>The Monarch School</title>
<style type="text/css">
@import url("css/style.css");
@import url("css/date_input.css");
</style>
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
dim cn
dim tokenvalue
dim rs

sub pass1
%>

      <div class="block2">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2>Reports</h2>
        </div>
        <div class="block_content">
          <form name="reports" action="results.asp" method="post">    
              <p><label>From:</label> 
              <input type="text" class="text date_picker" id="from_dt" name="from_dt"/>
              <font color='red'> <% =fromdt_error %> </font>
              </p>
              <p><label>To:</label> 
              <input type="text" class="text date_picker" id="to_dt" name="to_dt" value="<% =date() %>" />
              <font color='red'> <% =todt_error %> </font>
            </p>
                <p><label>Select Level:</label> <br />
                <select name="level" id="level" class="styled" width="100%">
                <option value="ALL">ALL</option>
                <option value="Novice">Novice</option>
                <option value="Apprentice">Apprentice</option>  
                <option value="Challenger">Challenger</option>  
                <option value="Voyager">Voyager</option>    
              </select></p>
            <hr/>
            <p>
              <center><input type="submit" class="submit long" value="Submit" /></center>
              <input type="hidden" id="token" name="token" value="2">
            </p>
          </form>
    </div>
        <div class="bendl"></div>
        <div class="bendr"></div>   
      </div>

<%
end sub
sub pass2

from_date = request.form("from_dt")

to_date = request.form("to_dt")
if mid(to_date,2,1) = "/" then to_date = "0" + to_date
if mid(to_date,5,1) = "/" then to_date = mid(to_date,1,3) + "0" + mid(to_date,4,6)

level = request.form("level")

if level = "ALL" then
'SET MIN DATES SQLSTRING
  MINsql_string="select A.* from grades A,(select student_id,min(B.date_taken) as minDate from grades B where "
  MINsql_string = MINsql_string + "date_taken BETWEEN '" + cstr(from_date) + " 00:00:00' AND '" + cstr(to_date) + " 23:59:59' group by student_id HAVING "
  MINsql_string = MINsql_string + "count(grades_id) >= 2) B where A.date_taken = B.minDate and A.student_id = B.student_id" 
 'SET MAX DATES SQLSTRING 
  MAXsql_string="select A.* from grades A,(select student_id,max(B.date_taken) as maxDate from grades B where "
  MAXsql_string = MAXsql_string + "date_taken BETWEEN '" + cstr(from_date) + " 00:00:00' AND '" + cstr(to_date) + " 23:59:59' group by student_id HAVING "
  MAXsql_string = MAXsql_string + "count(grades_id) >= 2) B where A.date_taken = B.maxDate and A.student_id = B.student_id"   
elseif level <> "ALL" and level <> "" then
 'SET MIN DATES SQLSTRING
 MINsql_string="select A.* from grades A,(select student_id,min(B.date_taken) as minDate from grades B where "
  MINsql_string = MINsql_string + "date_taken BETWEEN '" + cstr(from_date) + " 00:00:00' AND '" + cstr(to_date) + " 23:59:59'"
  MINsql_string = MINsql_string + " AND level= " +chr(39) +cstr(level) +chr(39) + " "
  MINsql_string = MINsql_string + "group by student_id HAVING "
  MINsql_string = MINsql_string + "count(grades_id) >= 2) B where A.date_taken = B.minDate and A.student_id = B.student_id"
 'SET MAX DATES SQLSTRING 
   MAXsql_string="select A.* from grades A,(select student_id,max(B.date_taken) as maxDate from grades B where "
  MAXsql_string = MAXsql_string + "date_taken BETWEEN '" + cstr(from_date) + " 00:00:00' AND '" + cstr(to_date) + " 23:59:59'"
  MAXsql_string = MAXsql_string + " AND level= " +chr(39) +cstr(level) +chr(39) + " "
  MAXsql_string = MAXsql_string + "group by student_id HAVING "
  MAXsql_string = MAXsql_string + "count(grades_id) >= 2) B where A.date_taken = B.maxDate and A.student_id = B.student_id"
end if

'BEGIN GETTING MIN AVERAGES FOR EACH TEST
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open MINsql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
MinArray = Array (0,0,0,0,0,0,0,0)
minRowCount = 0

while not rs.EOF
  minRowCount=MinRowCount+1
  for i = 0 to 7
  'the test grade columns start at 4, so rs(i+4)
	MinArray(i) = MinArray(i) + cint(rs(i+4))
  next

  rs.MoveNext
wend

rs.Close
set rs=nothing

'BEGIN GETTING MAX AVERAGES FOR EACH TEST
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open MAXsql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
MaxArray = Array (0,0,0,0,0,0,0,0)
maxRowCount = 0

while not rs.EOF
  maxRowCount=MaxRowCount+1
  for i = 0 to 7
  'the test grade columns start at 4, so rs(i+4)
	MaxArray(i) = MaxArray(i) + cint(rs(i+4))
  next

  rs.MoveNext
wend

rs.Close
set rs=nothing

'PRINT OUT MIN DATA 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open MINsql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

'PRINT OUT THE AVERAGES FOR EACH MIN DATE
if MinRowCount <> 0 then
	for i=0 to 7
	MinArray(i) = (MinArray(i)/MinRowCount)
	next
end if

'PRINT OUT THE MAX DATA
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open MAXsql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

'PRINT OUT THE AVERAGES FOR EACH MAX DATE
if MaxRowCount <> 0 then
	for i=0 to 7
	MaxArray(i) = (MaxArray(i)/MaxRowCount)
	next
end if

'CALCULATE THE % DIFFERENCE BETWEEN THE MIN AND MAX
Lvl_Improvement = Array (0,0,0,0,0,0,0,0)

if MaxRowCount <> 0 then
for i=0 to 7
	Lvl_Improvement(i) = MaxArray(i)-MinArray(i)
next
end if
%>

<div class="block">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2>graphs</h2>
            <ul class="tabs">
            <li><a href="#graph1">Average</a></li>
            <li><a href="#graph2">Margin</a></li>
            </ul>
        </div>

        <div class="block_content tab_content" id="graph1">
            
            <center><div id="chart1" style="width:860px;height:300px; margin: 15px 0px;"></div></center>

        </div>

        <div class="block_content tab_content" id="graph2">
          
          GRAPH 2 PLACEHOLDER

        </div>
          <div class="bendl"></div>
        <div class="bendr"></div>
      </div>

<div class="block">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
            <h2>SQL Query</h2>
        </div>
        <div class="block_content">

      </div>  
        <div class="bendl"></div>
        <div class="bendr"></div>
</div>


<%
rs.Close
set rs=nothing

end sub

tokenvalue=request.form("token")
select case tokenvalue
case ""
   call pass1
case "2"
  call pass2
end select
%>

 <br>
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
  <script type="text/javascript" src="js/jquery.select_skin.js"></script>
  <script type="text/javascript" src="js/jquery.pngfix.js"></script>
  <script type="text/javascript" src="js/custom.js"></script>  
</body>
</html>