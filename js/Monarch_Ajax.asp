<%

'
' MonarchAjax.asp
'
'
'   receives form submission from AjaxRequest.submit
'   (program: glajax_cross.htm)
'


LastName=Request.form("LastName")

set rs=Server.CreateObject("ADODB.Recordset")
SQLString="SELECT * FROM student WHERE "
fsw=0
if len(LastName) > 0 then
   SQLString=SQLString+" stu_lname LIKE '"+cstr(LastName) +"%'"
   fsw=1
end if

if len(LastName) = 0 then
   SQLString="Select * from student"
   fsw=1
end if

response.write "<p><table cellspacing='0' width='100%'>"
response.write "<td><center><b>Student ID</b></center></td>"
response.write "<td align='center'><b>First Name</b></td>"
response.write "<td align='center'><b>Middle Name</b></td>"
response.write "<td align='center'><b>Last Name</b></td>"
response.write "<td><center><b>Gender</center></b></td>"
response.write "<td><b>Level</b></td>"
response.write "<td align='center'><b>Diagnosis</b></td>"
response.write "<td><center><b>Modify</b></center></td></tr>"

'response.write "<p>SQL STRING="+cstr(SQLString)

if fsw=1 then

c=0
rs.open SQLString,"DSN=gl1181;UID=gl1181;PWD=YVT52ddnJ;"
while not rs.eof
	sid=cstr(rs("student_id"))
	viewlink="<center><a href='view/student.asp?token=3&sid="+sid+"'><img src='http://disc-nt.cba.uh.edu/Students/gl1181/project/user_16.png'</a></center>"
	modlink="<center><a href='modify/student.asp?token=1&sid="+sid+"'><img src='http://disc-nt.cba.uh.edu/Students/gl1181/project/pencil_16.png'</a></center>"
	s_fname=cstr(rs("stu_fname"))
	
    response.write "<tr><td align='center'>"
    response.write cstr(viewlink) +"</td><td align='center'>"
    response.write s_fname+"</td><td align='center'>"
    response.write cstr(rs("stu_mname"))+"</td><td align='center'>"
    response.write cstr(rs("stu_lname"))+"</td><td><center>"
    response.write cstr(rs("stu_gender"))+"</center></td><td align='center'>"
    response.write cstr(rs("level"))+"</center></td><td align='center'>"
    response.write cstr(rs("stu_diagnosis"))+"</td><td align='center'>"
    response.write cstr(modlink)+"</td></tr>"
    c=c+1
rs.movenext
wend
rs.close
set rs=nothing

end if
response.write "</table></span><p><center><b>"+cstr(c)+" matching records found" +"</b></center>"

%>