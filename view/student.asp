<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<html lang="en">
<head>
<title>The Monarch School</title>
<meta http-equiv="x-ua-compatible" content="IE=8">
<meta charset="utf-8">
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<link type="text/css" rel="stylesheet" href="../css/jquery.jqplot.css" />
<!--[if lt IE 9]><script language="javascript" type="text/javascript" src="../js/excanvas.min.js"></script><![endif]-->
<script src="../js/jquery.js" language="javascript" type="text/javascript" ></script>
<script src="../js/jquery.jqplot.js" language="javascript" type="text/javascript" ></script>
<script src="../js/jqplot.CategoryAxisRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="../js/jqplot.dateAxisRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="../js/jqplot.barRenderer.js" language="javascript" type="text/javascript" ></script>
<script src="../js/jqplot.pointLabels.js" language="javascript" type="text/javascript" ></script>
<style type="text/css">
@import url("../css/style.css");
</style>
</head>
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

Dim month_name 
month_name = Array("January","February","March","April","May","June","July","August","September","October","November","December")

'PULLS STUDENT ID FROM URL
sub pass1
bidvalue=request.querystring("sid")

'OPENS CONNECTION WITH SQL QUERY'
Set rs = Server.CreateObject("ADODB.Recordset")
sql_string="Select *,CONVERT(varchar(10),dob_date,101) as dobdate,CONVERT(varchar(10),ed_date,101) as eddate from student WHERE student_id=" +bidvalue
rs.open sql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
firstname_val=""
lastname_val=""
middlename_val=""
gender_val=""
level=""
dob=""
dob_year=""
diag_val=""
ed=""

while NOT rs.EOF
    firstname_val=cstr(rs("stu_fname"))
    lastname_val=cstr(rs("stu_lname"))
    middlename_val=cstr(rs("stu_mname"))
    gender_val=cstr(rs("stu_gender"))
    level=cstr(rs("level"))
    dob=cstr(rs("dobdate"))
    diag_val=cstr(rs("stu_diagnosis"))
    ed=cstr(rs("eddate"))
    rs.movenext
wend

mybday = CalculateAge(dob)

'CLOSING CONNECTION TO STUDENT TABLE'
rs.Close
set rs=nothing

'OPENS CONNECTION TO GRADES TABLE WITH QUERY'
Set rs = Server.CreateObject("ADODB.Recordset")
sql_string="Select *,CONVERT(varchar(10),date_taken,101) as dtdate from grades WHERE student_id=" +bidvalue + " order by date_taken desc, grades_id desc"
rs.open sql_string, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"

'GRADES COUNT'
Set rs2 = Server.CreateObject("ADODB.Recordset")
sql_string2="Select *,CONVERT(varchar(10),date_taken,101) as dtdate from grades WHERE student_id=" +bidvalue + " order by date_taken desc, grades_id desc"
rs2.open sql_string2, "DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
grades_count=0
while NOT rs2.EOF
    grades_count=grades_count+1
    rs2.movenext
wend
rs2.Close
set rs2=nothing
%>

<div class="block3">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2><%= cstr(firstname_val) +" " +cstr(lastname_val) +" (" +cstr(gender_val) +")" +" - " +cstr(mybday) +" Years Old - " +cstr(diag_val)%> </h2>
          <ul>
             <li><a href="../add/grades.asp?token=1&sid=<%=cstr(bidvalue)%>">Add Grades</a></li>
            <li>Current Level : <a href="../modify/student.asp?token=1&sid=<%=cstr(bidvalue)%>"><b><%=cstr(level)%></b></a></li>
            <li><a href="../delete/student.asp?token=1&sid=<%=cstr(bidvalue)%>">Delete Student</a></li>
          </ul>
        </div>
        <div class="block_content">
            <center><div id="chart1" style="width:860px;height:300px; margin: 15px 0px;"></div></center>
        </div>
        <div class="bendl"></div>
        <div class="bendr"></div>
      </div>
      <div class="block">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
          <h2>Grades - Total of <b><%= cstr(grades_count)%></b> Exams Taken</h2>
          <ul>
            <li>ENROLLMENT DATE: <% = cstr(ed)%></li>
          </ul>
        </div> 
        <div class="block_content">
<%
'CREATES A TABLE POPULATING GRADES TABLE'
response.write "<table cellspacing='0' width='100%'>"
response.write "<td width='20%'><b>Date Taken</b></td><td width='9%'><b>Level</b></td><td width='10%'><b>Inhibit</b></td><td width='10%'><b>Shift</b></td><td width='10%'><b>Emotional Control</b></td><td width='10%'><b>Initiate</b></td><td width='10%'><b>Working Memory</b></td><td width='10%'><b>Plan/Organize</b></td><td width='10%'><b>Org of Materials</b></td><td width='10%'><b>Monitor</b></td><td><b>Modify</b></td><td><b>Delete</b></td></tr>"

'CREATES DELETE AND MODIFY GRADES LINKS FOR EACH GRADE RECORD'
d=0
while NOT rs.EOF
  deletelink="<center><a href='../delete/grades.asp?token=1&sid=" +cstr(rs("student_id")) +"&gid=" +cstr(rs("grades_id"))  +"'><img src='../images/delete_16.png'</a></center>"
  modifylink="<center><a href='../modify/grades.asp?token=1&sid=" +cstr(rs("student_id")) +"&gid=" +cstr(rs("grades_id"))  +"'><img src='../images/pencil_16.png'</a></center>"
  'response.write "<tr><td>"+cstr(rs("grades_id"))+"</td>"
  response.write "<td>"+cstr(rs("dtdate")) +"</td>"
  response.write "<td>"+cstr(rs("level"))+"</td>"
  response.write "<td>"+cstr(rs("inhibit"))+"</td>"
  response.write "<td>"+cstr(rs("shift"))+"</td>"
  response.write "<td>"+cstr(rs("emotion"))+"</td>"
  response.write "<td>"+cstr(rs("initiate"))+"</td>"
  response.write "<td>"+cstr(rs("work_mem"))+"</td>"
  response.write "<td>"+cstr(rs("plan_org"))+"</td>"
  response.write "<td>"+cstr(rs("org_of_materials"))+"</td>"
  response.write "<td>"+cstr(rs("monitor"))+"</td>"
  response.write "<td>"+cstr(modifylink)+"</td>"
  response.write "<td>"+cstr(deletelink)+"</td></tr>"
  if d=0 then
  inhibit_val1 = cstr(rs("inhibit"))
  shift_val1 = cstr(rs("shift"))
  emotion_val1 = cstr(rs("emotion"))
  intiate_val1 = cstr(rs("initiate"))
  work_mem_val1 = cstr(rs("work_mem"))
  plan_org_val1 = cstr(rs("plan_org"))
  oom_val1 = cstr(rs("org_of_materials"))
  mon_val1 = cstr(rs("monitor"))
  newer_dt = cstr(rs("dtdate"))
  elseif d=1 then
  inhibit_val2 = cstr(rs("inhibit"))
  shift_val2 = cstr(rs("shift"))
  emotion_val2 = cstr(rs("emotion"))
  intiate_val2 = cstr(rs("initiate"))
  work_mem_val2 = cstr(rs("work_mem"))
  plan_org_val2 = cstr(rs("plan_org"))
  oom_val2 = cstr(rs("org_of_materials"))
  mon_val2 = cstr(rs("monitor"))  
  older_dt = cstr(rs("dtdate"))
  end if
  d=d+1
  rs.movenext
wend
response.write "</table>"

'CLOSING CONNECTION TO GRADES TABLE'
rs.Close
set rs=nothing
%>
        </div>
        <div class="bendl"></div>
        <div class="bendr"></div>
      </div>
      <div id="footer">
        <center>
          <img src="../images/footer.png">
          <p><b>The Monarch School</b></p>
        </center>
      </div>
    </div>
  </div>
<center>
<script type="text/javascript">

$(document).ready(function(){
      $.jqplot.config.enablePlugins = true;
      <%if d>1 then%>
	  <% ="var s1 = [" + cstr(inhibit_val2) + "," + cstr(shift_val2) + "," + cstr(emotion_val2) + "," + cstr(intiate_val2) + "," + cstr(work_mem_val2) +"," + cstr(plan_org_val2) +"," + cstr(oom_val2) + "," + cstr(mon_val2) %>];
	  <%end if%>
	  <%if d>0 then%>
      var s2 = [<% = cstr(inhibit_val1) %>, <% = cstr(shift_val1) %>, <% = cstr(emotion_val1) %>, <% = cstr(intiate_val1) %>, <% = cstr(work_mem_val1) %>, <% = cstr(plan_org_val1) %>, <% = cstr(oom_val1) %>, <% = cstr(mon_val1) %>];
	  <%end if%>
      var ticks = ['Inhibit', 'Shift', 'Emotional Control', 'Initiate', 'Working Memory', 'Plan/Organize', 'Org of Materials', 'Monitor'];
         
      var plot1 = $.jqplot('chart1', [<%if d>1 then%> <%="s1,"%> <%end if%><%if d>0 then%>s2<%end if%>], {
        animate: true,
        seriesColors: [ "#5CB9E4", "#EC9649"],
        // The "seriesDefaults" option is an options object that will
        // be applied to all series in the chart.
        seriesDefaults:{
            renderer:$.jqplot.BarRenderer,
            shadow: false,
            rendererOptions: {fillToZero: true}
        },
        // Custom labels for the series are specified with the "label"
        // option on the series option.  Here a series option object
        // is specified for each series.
        series:[
        <%if d>1 then%>
      {label:'<% =older_dt %>'},
    <%end if%>
    <%if d>0 then%>
            {label:'<% =newer_dt %>'},
    <%end if%>
        ],
        // Show the legend and put it outside the grid, but inside the
        // plot container, shrinking the grid to accomodate the legend.
        // A value of "outside" would not shrink the grid and allow
        // the legend to overflow the container.
        legend: {
            show: true, 
            placement: 'insideGrid',
            location: 'ne'
        },
        grid: {
        background: '#ffffff',
        shadow: false,
        borderWidth: 0
        },
        axesDefaults: {
        tickOptions: {
            showMark: false
        }
        },
        axes: {
            // Use a category axis on the x axis and use our custom ticks.
            xaxis: {
                renderer: $.jqplot.CategoryAxisRenderer,
                rendererOptions: {
                showGridline: false
                },
                ticks: ticks
            },
            // Pad the y axis just a little so bars can get close to, but
            // not touch, the grid boundaries.  1.2 is the default padding.
            yaxis: {
                pad: .8,
                tickOptions: {formatString: '%d'},
                min: 0,
                max: 100,
                ticks: [0,20,40,60,80,100,120,140]
            }
        }
    });
});
</script>

<%
end sub

sub passerror
     response.write "<p>INVALID TOKEN VALUE. token="+cstr(tokenvalue)
end sub

tokenvalue=request.querystring("token")
select case tokenvalue
case "3"
   call pass1
case else
   call passerror
end select

'<!-- function to calculate age -->

  Function CalculateAge(DateOfBirth) 
    TodaysDate = Date()
    intAge = DateDiff("yyyy", DateOfBirth, TodaysDate)
    If TodaysDate < DateSerial(Year(TodaysDate), Month(DateOfBirth), Day(DateOfBirth)) Then
      intAge = intAge - 1
    End If
    CalculateAge = intAge
  End Function
    
%>

</center>
</body>
</html>