<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim sch_tab(10,10)
dim car_tab(20,10)
dim qul_tab(20,10)

acpt_emp_no = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)

emp_no = request("emp_no")
be_pg = request("be_pg")
be_pg1 = "insa_card00.asp"
page = request("page")

view_sort = request("view_sort")
page_cnt = request("page_cnt")


Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_sch = Server.CreateObject("ADODB.Recordset")
Set rs_car = Server.CreateObject("ADODB.Recordset")
Set rs_qul = Server.CreateObject("ADODB.Recordset")
Set RsschCnt = Server.CreateObject("ADODB.Recordset")
Set RscarCnt = Server.CreateObject("ADODB.Recordset")
Set RsqulCnt = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
Set rs = DbConn.Execute(SQL)

if not rs.EOF or not rs.BOF then

    if rs("emp_image") = "" or isnull(rs("emp_image")) then 
	            photo_image = ""
		else
	            photo_image = "/emp_photo/" + rs("emp_image")
    end if

    emp_person2 = rs("emp_person2")
    if emp_person2 <> "" then
	   sex_id = mid(cstr(emp_person2),1,1)
	   if sex_id = "1" then
	         emp_sex = "��"
		  else
		     emp_sex = "��"
	   end if
	end if

    if rs("emp_military_date1") = "1900-01-01" then
           emp_military_date1 = ""
           emp_military_date2 = ""
       else 
           emp_military_date1 = rs("emp_military_date1")
           emp_military_date2 = rs("emp_military_date2")
    end if
    if rs("emp_marry_date") = "1900-01-01" then
           emp_marry_date = ""
       else 
     	   emp_marry_date = rs("emp_marry_date")
    end if

'�з»��� db
for i = 0 to 10
'	com_tab(i) = ""
'	com_sum(i) = 0
	for j = 0 to 10
		sch_tab(i,j) = ""
'		com_in(i,j) = 0
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

end if
title_line = " �λ� ��� ī�� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("�Ҽ��� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_card00.asp" method="post" name="frm">
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
						</colgroup>
						<tbody>
                        <% if not rs.EOF or not rs.BOF then %>
							<tr>
                                <%
								'<th colspan="2" rowspan="4" class="first">&nbsp;</th>
								'<img src="emp_photo/�̻���.jpg" width=110 height=120 alt="">
								emp_email = rs("emp_email") + "@k-won.co.kr"
								%>
                                <td colspan="2" rowspan="4" class="first">
                                <img src="<%=photo_image%>" width=110 height=120 alt="">
                                </td>
								<th>���&nbsp;&nbsp;��ȣ</th>
                                <td class="left"><%=rs("emp_no")%></td> 
								<th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td colspan="2" class="left"><%=rs("emp_org_code")%>)<%=rs("emp_org_name")%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
								<td class="left"><%=rs("emp_jikgun")%>-<%=rs("emp_jikmu")%>&nbsp;</td>
                                <th>�ֹι�ȣ</th>
								<td colspan="2" class="left"><%=rs("emp_person1")%>-<%=rs("emp_person2")%>&nbsp;&nbsp;(<%=emp_sex%>)</td>
                 			</tr>
							<tr>
								<th>����(�ѱ�)</th>
                                <td class="left"><%=rs("emp_name")%>&nbsp;</td>
								<th>����(����)</th>
								<td colspan="2" class="left"><%=rs("emp_ename")%>&nbsp;</td>
                                <th>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;å</th>
                                <td class="left"><%=rs("emp_position")%>&nbsp;</td>
								<th>����(��)/������</th>
								<td colspan="2" class="left">(<%=rs("emp_grade")%>)&nbsp;<%=rs("emp_job")%>&nbsp;/&nbsp;<%=rs("emp_grade_date")%></td>
                 			</tr>                            
							<tr>
                                <th>�����Ի���</th>
                                <td class="left"><%=rs("emp_first_date")%></td>
                                <th>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</th>
                                <td class="left"><%=rs("emp_in_date")%>&nbsp;</td>
                                <th>��ȭ��ȣ</th>
								<td class="left"><%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%>&nbsp;</td>
								<th>�ּ�(��)</th>
								<td colspan="3" class="left"><%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></td>
                            </tr>
                            <tr>
                                <th>�ټӱ����</th>
                                <td class="left"><%=rs("emp_gunsok_date")%>&nbsp;</td>
                                <th>���������</th>
                                <td class="left"><%=rs("emp_end_gisan")%>&nbsp;</td>
                                <th>�ڵ���</th>
								<td class="left"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%>&nbsp;</td>
                                <th>e-�����ּ�</th>
								<td colspan="3" class="left"><%=emp_email%>&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="10" class="left">�� �з� ���� ��</th> 
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_school_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','schoolview','scrollbars=yes,width=800,height=400')">�� �з� ������</a>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="3">�Ⱓ</th>
                                <th colspan="2">�б���</th>
                                <th colspan="2">�а�</th>
                                <th colspan="2">����</th>
                                <th>������</th>  
                                <th>����</th>
                                <th>����</th>
                            </tr>
                				<td colspan="3" class="left"><%=sch_tab(1,1)%>&nbsp;~&nbsp;<%=sch_tab(1,2)%></td>
                                <td colspan="2" class="left"><%=sch_tab(1,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_tab(1,4)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_tab(1,5)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(1,6)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(1,7)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(1,8)%>&nbsp;</td>
                             </tr>
                            </tr>
                				<td colspan="3" class="left"><%=sch_tab(2,1)%>&nbsp;~&nbsp;<%=sch_tab(2,2)%></td>
                                <td colspan="2" class="left"><%=sch_tab(2,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_tab(2,4)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_tab(2,5)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(2,6)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(2,7)%>&nbsp;</td>
                                <td class="left"><%=sch_tab(2,8)%>&nbsp;</td>
                             </tr>                             
                            <tr>
                                <th colspan="10" class="left">�� ���� ��� ���� ��</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_career_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','careerview','scrollbars=yes,width=800,height=400')">�� ��� ������</a>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="3">�����Ⱓ</th>
                                <th colspan="2">ȸ���</th>
                                <th colspan="2">��  ��</th>
                                <th>����</th>
                                <th colspan="4">������</th>
                            </tr>
                            <tr>
                                <td colspan="3" class="left"><%=car_tab(1,1)%>&nbsp;~&nbsp;<%=car_tab(1,2)%></td>
                                <td colspan="2" class="left"><%=car_tab(1,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=car_tab(1,4)%>&nbsp;</td>
                                <td colspan="1" class="left"><%=car_tab(1,5)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=car_tab(1,6)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=car_tab(2,1)%>&nbsp;~&nbsp;<%=car_tab(2,2)%></td>
                                <td colspan="2" class="left"><%=car_tab(2,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=car_tab(2,4)%>&nbsp;</td>
                                <td colspan="1" class="left"><%=car_tab(2,5)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=car_tab(2,6)%>&nbsp;</td>
                             </tr>
                             <tr>                             
                                <th colspan="10" class="left">�� �ڰ��� ���� ��</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_qual_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','qualview','scrollbars=yes,width=800,height=400')">�� �ڰ� ������</a>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="3">�ڰ��� ����</th>
                                <th>���</th>
                                <th colspan="2">�հݳ����</th>
                                <th colspan="2">�߱� �����</th>
                                <th colspan="4">�ڰ� ��Ϲ�ȣ</th>
                            </tr>
                            <tr>
                                <td colspan="3" class="left"><%=qul_tab(1,1)%>&nbsp;</td>
                                <td class="left"><%=qul_tab(1,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(1,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(1,4)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=qul_tab(1,5)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=qul_tab(2,1)%>&nbsp;</td>
                                <td class="left"><%=qul_tab(2,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(2,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(2,4)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=qul_tab(2,5)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=qul_tab(3,1)%>&nbsp;</td>
                                <td class="left"><%=qul_tab(3,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(3,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qul_tab(3,4)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=qul_tab(3,5)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <th>���� �����Ⱓ</th>
                                <td colspan="2" class="left"><%=mid(emp_military_date1,1,7)%>~<%=mid(emp_military_date2,1,7)%>&nbsp;</td>
                                <th>��������/���</th>
                                <td class="left"><%=rs("emp_military_id")%> - <%=rs("emp_military_grade")%>&nbsp;</td>
                                <th>��������</th>
								<td colspan="2" class="left"><%=rs("emp_military_comm")%>&nbsp;</td>
                                <th>��ȥ�����</th>
                                <td class="left"><%=emp_marry_date%>&nbsp;</td>
                                <th>����</th>
                                <td class="left"><%=rs("emp_faith")%>&nbsp;</td>
							</tr>
                      <% end if %>
                      </tbody>
					</table>
				</div>      
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="40%">
					<div class="btnCenter">
                    <a href="#" class="btnType04" onClick="pop_Window('insa_card_print.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card_pop','scrollbars=yes,width=750,height=600')">�λ���ī�� ���</a>
              <% if acpt_emp_no = "900002" then %>
                    <a href="insa_excel_card_print.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>" class="btnType04">�����ٿ�ε�</a>
              <% end if %>
					</div>                  
                  	</td>
				    <td>
                    <div class="btnCenter">
                    <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" class="btnType04" onClick="pop_Window('insa_card01.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg1%>&acpt_user=<%=acpt_user%>','emp_card1_pop','scrollbars=yes,width=1250,height=750')">�� �λ��� ��Ÿ����</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
        	</form>
		</div>				
	</div>        				
	</body>
</html>
