<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<html lang="en">
<head>
<title>The Monarch School</title>
<meta http-equiv="x-ua-compatible" content="IE=8">
<meta charset="utf-8">
<link rel="shortcut icon" type="image/x-icon" href="../images/favicon.ico">
<style type="text/css">
@import url("css/style.css");
@import url("css/facebox.css");
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


tokenvalue=request.form("token")
select case tokenvalue
case ""
   call pass1
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