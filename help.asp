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
</style>
<!--[if lt IE 8]><style type="text/css" media="all">@import url('css/ie.css");</style><![endif]-->
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

<div class="block">
        <div class="block_head">
          <div class="bheadl"></div>
          <div class="bheadr"></div>
            <h2>documentation</h2>
        </div>
        <div class="block_content">
<p><strong>What is the purpose of this system?</strong><br />
This system was designed to track student improvement since their enrollment in the school.</p>

<h3><strong>Student Records</strong></h3>

<p><strong>How do I add a student?</strong><br />
Click Add Student in the menu at the top of the page. Once you do, you will be taken to the Add Student page on which you can enter all of the necessary data for the student you wish to add to the database.. All fields are mandatory except for Middle Name, which may be left blank. Once you have entered all of the data for the student and have also verified that all of the data is correct, press the Submit button. The student will be added and you will be taken back to the site home page, from which you may add another student, if you wish.</p>

<p><small>Note: All data can be modified later from the Modify Student screen. (See the next section.) Note: You can clear all fields and start over by pressing the Reset button.</small></p>

<p><strong>How do I search for students?</strong><br />
On the site&rsquo;s top page, look for a text box labeled &ldquo;Search&rdquo; and click it to activate it. Begin typing the first few letters of the target student&rsquo;s last name. You will see the list of students below actively update as you type.</p>

<p><small>Note: If the student you are looking for does not appear, double-check the name&rsquo;s spelling. Perhaps that student is not yet in the database.</small><br />
<small>Tip: When the search box is blank, type Backspace to see a list of all students currently in the database.</small></p>

<p><strong>How do I modify a student?</strong><br />
Occasionally, you will find that a student record needs to be updated, such as when a student moves into another school program or if erroneous data was entered in the past and needs to be updated. Warning: Not all users have access to modify records. If you do not, please speak with the school&rsquo;s designated administrator. To modify a student&rsquo;s data, you must first search for that student. (See the instructions above in &ldquo;How do I search for students?&rdquo;) Once you see the target student in the list of search results, you will see a pencil edit at the line of that line. Press the pencil icon to bring up the Modify Student screen. This screen is almost identical to the Add Student screen and behaves the same way. Change any data that needs to be updated and press Submit.</p>

<p><strong>How do I delete a student?</strong><br />
There might rarely be a case in which a student record needs to be deleted. Warning: Not all users have access to delete records. If you do not, please speak with the school&rsquo;s designated administrator. To delete a student record (which will, in turn, deleted all test results to which that student is linked), first, search for the student in question and find his or her name in the search results. (See the instructions above in &ldquo;How do I search for students?&rdquo;) Click the icon on the left (under &ldquo;Student ID&rdquo;) to bring up the student&rsquo;s record page. Press &ldquo;Delete Student&rdquo; at the upper right corner of the page to delete the current student from the database permanently.</p>

<p><small>Warning: All test results entered for deleted students will automatically be deleted as well! Be careful with this function.</small></p>

<h3><strong>Test Result Records</strong></h3>

<p>Test results are recorded for each student in the database. Each time a student takes a test and those scores are added to the database, a new test record is created for that student.</p>

<p><strong>How do I view a student&rsquo;s test results?</strong><br />
First, search for the student in question and find his or her name in the search results. (See the instructions above in &ldquo;How do I search for students?&rdquo;) Click the icon on the left to bring up the student record.</p>

<p><strong>How do I add a new student test result?</strong><br />
When a student has taken a new test and you have those results, open the Student Record and click the &ldquo;Add Grades&rdquo; link near the upper right corner of the screen. Enter the grade that the student earned for each.</p>
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
</body>
</html>