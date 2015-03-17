<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<html lang="en">
<head>
<title>The Monarch School</title>
<meta http-equiv="x-ua-compatible" content="IE=8">
<meta charset="utf-8">
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<link type="text/css" rel="stylesheet" href="css/jquery.jqplot.css" />
<!--[if lt IE 9]><script language="javascript" type="text/javascript" src="js/excanvas.min.js"></script><![endif]-->
<script src="js/jquery.js" language="javascript" type="text/javascript" ></script>
<script src="js/jquery.jqplot.js" language="javascript" type="text/javascript" ></script>
<script src="js/jqplot.CategoryAxisRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="js/jqplot.dateAxisRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="js/jqplot.barRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="js/jqplot.pointLabels.js" language="javascript" type="text/javascript" ></script>
<style type="text/css">
@import url("css/style.css");
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

sub pass2

from_date = request.form("from_dt")
to_date = request.form("to_dt")
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

<div class="block3">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2>Average Raw Score Comparison</h2>
        </div>
        <div class="block_content" id="graph1">  
           <center><div id="chart2" style="width:860px;height:300px; margin: 15px 0px;"></div></center>
        </div>
          <div class="bendl"></div>
        <div class="bendr"></div>
      </div>
     
<div class="block3">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2>margin of improvement</h2>
        </div>
        <div class="block_content" id="graph1">  
            <center><div id="chart1" style="width:860px;height:300px; margin: 15px 0px;"></div></center>
        </div>
          <div class="bendl"></div>
        <div class="bendr"></div>
      </div>

      <div id="footer">
        <center>
          <img src="images/footer.png">
          <p><b>The Monarch School</b></p>
        </center>
      </div>
    </div>
  </div>
<center>

<script class="code" type="text/javascript">$(document).ready(function(){
  var line1=[['Inhibit', <% =Lvl_Improvement(0) %>], ['Shift', <% =Lvl_Improvement(1) %>], ['Emotional Control', <% =Lvl_Improvement(2) %>], ['Initiate', <% =Lvl_Improvement(3) %>],
      ['Working Memory', <% =Lvl_Improvement(4) %>], ['Plan/Organize', <% =Lvl_Improvement(5) %>], ['Org of Materials', <% =Lvl_Improvement(6) %>], ['Monitor', <% =Lvl_Improvement(7) %>]];
  var plot1 = $.jqplot('chart1', [line1], {
      animate: true,
      seriesColors: [ "#5CB9E4"],
      seriesDefaults:{
                shadow: false,
                pointLabels: { show: true }
        },
        axesDefaults: {
        tickOptions: {
            showMark: false,
        }
        },
      axes:{
        xaxis:{
          renderer:$.jqplot.CategoryAxisRenderer,
          tickOptions:{
            formatString:'%b&nbsp;%#d',
            showMark: false,
          } 
        },
        yaxis:{
          tickOptions:{
            formatString:'%.1f'
            }
        }
      },
      grid: {
        background: '#ffffff',
        shadow: false,
        borderWidth: 0,
        },
      highlighter: {
        show: true,
        sizeAdjust: 7.5
      },
      cursor: {
        show: false
      }
  });
});
</script>

<script class="code" type="text/javascript">$(document).ready(function(){
        $.jqplot.config.enablePlugins = true;
        var s1 = [<% =MinArray(0) %>, <% =MinArray(1) %>, <% =MinArray(2) %>, <% =MinArray(3) %>, <% =MinArray(4) %>, <% =MinArray(5) %>, <% =MinArray(6) %>, <% =MinArray(7) %>];
        var s2 = [<% =MaxArray(0) %>, <% =MaxArray(1) %>, <% =MaxArray(2) %>, <% =MaxArray(3) %>, <% =MaxArray(4) %>, <% =MaxArray(5) %>, <% =MaxArray(6) %>, <% =MaxArray(7) %>];
        var ticks = ['Inhibit', 'Shift', 'Emotional Control', 'Initiate', 'Working Memory', 'Plan/Organize', 'Org of Materials', 'Monitor'];
        
        plot1 = $.jqplot('chart2', [s1, s2], {
            // Only animate if we're not using excanvas (not in IE 7 or IE 8)..
            animate: true,
            seriesColors: [ "#5CB9E4", "#EC9649"],
            seriesDefaults:{
                renderer:$.jqplot.BarRenderer,
                shadow: false,
                rendererOptions: {fillToZero: true},
                pointLabels: { show: true }
            },
        series:[
        {label:'Oldest Exams'},
        {label:'Latest Exams'},
        ],
        legend: {
            show: true, 
            placement: 'insideGrid',
            location: 'se'
        },
        grid: {
        background: '#ffffff',
        shadow: false,
        borderWidth: 0,
        },
        axesDefaults: {
        tickOptions: {
            showMark: false,
        }
        },
        axes: {
            // Use a category axis on the x axis and use our custom ticks.
            xaxis: {
                renderer: $.jqplot.CategoryAxisRenderer,
                rendererOptions: {
                showGridline: false,
                },
                ticks: ticks
            },
            // Pad the y axis just a little so bars can get close to, but
            // not touch, the grid boundaries.  1.2 is the default padding.
            yaxis: {
                pad: .8,
                tickOptions: {formatString: '%.1f'},
                min: 0,
                max: 100,
                ticks: [0,20,40,60,80,100,120]
            }
        }
        });
    });
</script>

<%
end sub

tokenvalue=request.form("token")
select case tokenvalue
case "2"
   call pass2
case else
end select
%>

</center>
</body>
</html>