<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(20,10)
dim car_tab(20,10)
dim qul_tab(20,10)
dim fam_tab(20,10)
dim edu_tab(20,10)
dim lan_tab(20,10)
	 
view_condi=Request("view_condi")

curr_date = datevalue(mid(cstr(now()),1,10))

if view_condi = "" then
	view_condi = "��ü"
end if

title_line = "������Ȳ(" + view_condi + ")" + cstr(curr_date)

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_sch = Server.CreateObject("ADODB.Recordset")
Set rs_car = Server.CreateObject("ADODB.Recordset")
Set rs_qul = Server.CreateObject("ADODB.Recordset")
Set RsschCnt = Server.CreateObject("ADODB.Recordset")
Set RscarCnt = Server.CreateObject("ADODB.Recordset")
Set RsqulCnt = Server.CreateObject("ADODB.Recordset")

Set Rs_fam = Server.CreateObject("ADODB.Recordset")
Set rs_app = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_lan = Server.CreateObject("ADODB.Recordset")
Set rs_stay = Server.CreateObject("ADODB.Recordset")
Set RsfamCnt = Server.CreateObject("ADODB.Recordset")
Set RsappCnt = Server.CreateObject("ADODB.Recordset")
Set RseduCnt = Server.CreateObject("ADODB.Recordset")
Set RslanCnt = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'if view_condi = "��ü" then
       Sql = "SELECT * FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_in_date,emp_no,emp_name ASC" 
'   else	   
'	   Sql = "SELECT * FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_company = '"&view_condi&"') and (emp_no < '900000') ORDER BY emp_in_date,emp_no,emp_name ASC" 
'end if

Rs.Open Sql, Dbconn, 1
	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th colspan="20" scope="col">�⺻����</th>
                                <th colspan="7" scope="col">�з»���</th>
                                <th colspan="5" scope="col">��»���</th>
                                <th colspan="5" scope="col">�ڰ��� ��Ȳ</th>
							</tr>
                            <tr>
								<th class="first" scope="col">���</th>
                                <th scope="col">����</th>
                                <th scope="col">�������</th>
								<th scope="col">�ֹι�ȣ</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                                <th scope="col">��å</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">����</th>
                                <th scope="col">�����</th>
                                <th scope="col">��</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">����ó</th>
                                <th scope="col">����óȸ��</th>
                                <th scope="col">�����Ի���</th>
                                <th scope="col">�Ի���</th>
                                <th scope="col">�����з�</th>
                                <th scope="col">���ּ�</th>
                                <th scope="col">�ڵ���</th>
                                <th scope="col">e����</th>
                                <th scope="col">��������</th>
                                
                                <th scope="col">�Ⱓ</th>
                                <th scope="col">�б���</th>
                                <th scope="col">�а�</th>
                                <th scope="col">����</th>
                                <th scope="col">������</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                                
                                <th scope="col">�����Ⱓ</th>
                                <th scope="col">ȸ���</th>
                                <th scope="col">�μ�</th>
                                <th scope="col">����</th>
                                <th scope="col">������</th>
                                
                                <th scope="col">�ڰ�����</th>
                                <th scope="col">���</th>
                                <th scope="col">�հ�����</th>
                                <th scope="col">�߱ޱ��</th>
                                <th scope="col">�ڰ�����ȣ</th>

							</tr>
						</thead>
						<tbody>
			<%
						do until rs.eof
						
						   emp_no = rs("emp_no")

'�з»��� db
for i = 0 to 20
	for j = 0 to 10
		sch_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_school where sch_empno = '"&emp_no&"' order by sch_empno, sch_seq asc"
	Rs_sch.Open Sql, Dbconn, 1	
	while not rs_sch.eof
		k = k + 1
		sch_tab(k,1) = rs_sch("sch_start_date")
		sch_tab(k,2) = rs_sch("sch_end_date")
		sch_tab(k,3) = rs_sch("sch_school_name")
		sch_tab(k,4) = rs_sch("sch_dept")
		sch_tab(k,5) = rs_sch("sch_major")
		sch_tab(k,6) = rs_sch("sch_sub_major")
		sch_tab(k,7) = rs_sch("sch_degree")
		sch_tab(k,8) = rs_sch("sch_finish")
		rs_sch.movenext()
	Wend
    rs_sch.close()		
	k_sch = k				


'��»��� db
for i = 0 to 20
	for j = 0 to 10
		car_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_career where career_empno = '"&emp_no&"' order by career_empno, career_seq asc"
	Rs_car.Open Sql, Dbconn, 1	
	while not rs_car.eof
		k = k + 1
		car_tab(k,1) = rs_car("career_join_date")
		car_tab(k,2) = rs_car("career_end_date")
		car_tab(k,3) = rs_car("career_office")
		car_tab(k,4) = rs_car("career_dept")
		car_tab(k,5) = rs_car("career_position")
		car_tab(k,6) = rs_car("career_task")
		rs_car.movenext()
	Wend
    rs_car.close()	
    k_car = k		

'�ڰݻ��� db
for i = 0 to 20
	for j = 0 to 10
		qul_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_qual where qual_empno = '"&emp_no&"' order by qual_empno, qual_seq asc"
	rs_qul.Open Sql, Dbconn, 1	
	while not rs_qul.eof
		k = k + 1
		qul_tab(k,1) = rs_qul("qual_type")
		qul_tab(k,2) = rs_qul("qual_grade")
		qul_tab(k,3) = rs_qul("qual_pass_date")
		qul_tab(k,4) = rs_qul("qual_org")
		qul_tab(k,5) = rs_qul("qual_no")
		rs_qul.movenext()
	Wend
    rs_qul.close()	
	k_qul = k	
	
	if rs("emp_birthday") = "1900-01-01" then
		   emp_birthday = ""
	   else 
		   emp_birthday = rs("emp_birthday")
	end if
	
	emp_email = rs("emp_email") + "@k-won.co.kr"					   
						    

						   for jj = 1 to 20

							   if jj = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_birthday%></td>

									<td class="left" bgcolor="#EEFFFF"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></td>

                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_grade")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_job")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_position")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_bonbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_team")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_org_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_reside_place")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_reside_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_first_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_in_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_last_edu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_email%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_military_id")%></td>
                                    
								    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,1)%>&nbsp;~&nbsp;<%=sch_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,5)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,6)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,7)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,8)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,1)%>&nbsp;~&nbsp;<%=car_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,5)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,6)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,1)%>&nbsp;</td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,5)%></td>

						         </tr>
            <%
			                    else
								   if sch_tab(jj,1) <> "" or car_tab(jj,1) <> "" or qul_tab(jj,1) <> "" then
		    %>		
                                 <tr>
								    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=sch_tab(jj,1)%>&nbsp;~&nbsp;<%=sch_tab(jj,2)%></td>
                                    <td class="left" ><%=sch_tab(jj,3)%></td>
                                    <td class="left" ><%=sch_tab(jj,4)%></td>
                                    <td class="left" ><%=sch_tab(jj,5)%></td>
                                    <td class="left" ><%=sch_tab(jj,6)%></td>
                                    <td class="left" ><%=sch_tab(jj,7)%></td>
                                    <td class="left" ><%=sch_tab(jj,8)%></td>
                                    
                                    <td class="left" ><%=car_tab(jj,1)%>&nbsp;~&nbsp;<%=car_tab(jj,2)%></td>
                                    <td class="left" ><%=car_tab(jj,3)%></td>
                                    <td class="left" ><%=car_tab(jj,4)%></td>
                                    <td class="left" ><%=car_tab(jj,5)%></td>
                                    <td class="left" ><%=car_tab(jj,6)%></td>
                                    
                                    <td class="left" ><%=qul_tab(jj,1)%>&nbsp;</td>
                                    <td class="left" ><%=qul_tab(jj,2)%></td>
                                    <td class="left" ><%=qul_tab(jj,3)%></td>
                                    <td class="left" ><%=qul_tab(jj,4)%></td>
                                    <td class="left" ><%=qul_tab(jj,5)%></td>
						         </tr>            
            <%            							
							       end if
							 end if
	                       next
							  
						   rs.movenext()
						loop
						rs.close()
		    %>						
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
