<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

dim fam_tab(10,10)
dim app_tab(50,30)
dim edu_tab(10,10)
dim lan_tab(10,10)

emp_no = request("emp_no")
be_pg = request("be_pg")
page = request("page")

view_sort = request("view_sort")
page_cnt = request("page_cnt")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_fam = Server.CreateObject("ADODB.Recordset")
Set rs_app = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_lan = Server.CreateObject("ADODB.Recordset")
Set rs_stay = Server.CreateObject("ADODB.Recordset")
Set RsfamCnt = Server.CreateObject("ADODB.Recordset")
Set RsappCnt = Server.CreateObject("ADODB.Recordset")
Set RseduCnt = Server.CreateObject("ADODB.Recordset")
Set RslanCnt = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect


Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
Set rs = DbConn.Execute(SQL)
if not rs.EOF or not rs.BOF then
'�Է¹��� ���ϴ� ��¥�ʵ带 ó�� 1900-01-01�� �ϳ���..�Ф�
if rs("emp_end_date") = "1900-01-01" then
   emp_end_date = ""
   else
   emp_end_date = rs("emp_end_date")
end if
if rs("emp_birthday") = "1900-01-01" then
   emp_birthday = ""
   else
   emp_birthday = rs("emp_birthday")
end if
if rs("emp_grade_date") = "1900-01-01" then
   emp_grade_date = ""
   else
   emp_grade_date = rs("emp_grade_date")
end if
if rs("emp_org_baldate") = "1900-01-01" then
   emp_org_baldate = ""
   else
   emp_org_baldate = rs("emp_org_baldate")
end if
if rs("emp_sawo_date") = "1900-01-01" then
   emp_sawo_date = ""
   else
   emp_sawo_date = rs("emp_sawo_date")
end if

'�������� db
for i = 0 to 10
	for j = 0 to 10
		fam_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_family where family_empno = '"&emp_no&"' order by family_empno, family_seq asc"
	rs_fam.Open Sql, Dbconn, 1
	while not rs_fam.eof
		k = k + 1
		fam_tab(k,1) = rs_fam("family_rel")
		fam_tab(k,2) = rs_fam("family_name")
		fam_tab(k,3) = rs_fam("family_birthday")
		fam_tab(k,4) = rs_fam("family_birthday_id")
		fam_tab(k,5) = rs_fam("family_job")
		fam_tab(k,6) = rs_fam("family_person1")
		fam_tab(k,7) = rs_fam("family_person2")
		fam_tab(k,8) = rs_fam("family_live")
		rs_fam.movenext()
	Wend
    rs_fam.close()

'�߷ɻ��� db
for i = 0 to 50
	for j = 0 to 30
		app_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_appoint where app_empno = '"&emp_no&"' order by app_empno, app_seq asc"
	rs_app.Open Sql, Dbconn, 1
	while not rs_app.eof
		k = k + 1
		app_tab(k,1) = rs_app("app_date")
		app_tab(k,2) = rs_app("app_id")
		app_tab(k,3) = rs_app("app_id_type")
		app_tab(k,4) = rs_app("app_to_company")
		app_tab(k,5) = rs_app("app_to_orgcode")
		app_tab(k,6) = rs_app("app_to_org")
		app_tab(k,7) = rs_app("app_to_grade")
		app_tab(k,8) = rs_app("app_to_job")
		app_tab(k,9) = rs_app("app_to_position")
		app_tab(k,10) = rs_app("app_to_enddate")
		app_tab(k,11) = rs_app("app_be_company")
		app_tab(k,12) = rs_app("app_be_orgcode")
		app_tab(k,13) = rs_app("app_be_org")
		app_tab(k,14) = rs_app("app_be_grade")
		app_tab(k,15) = rs_app("app_be_job")
		app_tab(k,16) = rs_app("app_be_position")
		app_tab(k,17) = rs_app("app_be_enddate")
		app_tab(k,18) = rs_app("app_start_date")
		app_tab(k,19) = rs_app("app_finish_date")
		app_tab(k,20) = rs_app("app_reward")
		app_tab(k,21) = rs_app("app_comment")
		rs_app.movenext()
	Wend
    rs_app.close()


'�������� db
for i = 0 to 10
	for j = 0 to 10
		edu_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_edu where edu_empno = '"&emp_no&"' order by edu_empno, edu_seq asc"
	rs_edu.Open Sql, Dbconn, 1
	while not rs_edu.eof
		k = k + 1
		edu_tab(k,1) = rs_edu("edu_name")
		edu_tab(k,2) = rs_edu("edu_office")
		edu_tab(k,3) = rs_edu("edu_finish_no")
		edu_tab(k,4) = rs_edu("edu_start_date")
		edu_tab(k,5) = rs_edu("edu_end_date")
		edu_tab(k,6) = rs_edu("edu_comment")
		rs_edu.movenext()
	Wend
    rs_edu.close()

'���л��� db
for i = 0 to 10
	for j = 0 to 10
		lan_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_language where lang_empno = '"&emp_no&"' order by lang_empno, lang_seq asc"
	rs_lan.Open Sql, Dbconn, 1
	while not rs_lan.eof
		k = k + 1
		lan_tab(k,1) = rs_lan("lang_id")
		lan_tab(k,2) = rs_lan("lang_id_type")
		lan_tab(k,3) = rs_lan("lang_point")
		lan_tab(k,4) = rs_lan("lang_grade")
		lan_tab(k,5) = rs_lan("lang_get_date")
		rs_lan.movenext()
	Wend
    rs_lan.close()

'�Ǳٹ����ּ�
        stay_name = rs("emp_stay_name")
		stay_sido = ""
		stay_gugun = ""
		stay_dong = " "
		stay_addr = ""
		stay_code = rs("emp_stay_code")
        if stay_code <> "" then
		   Sql="select * from emp_stay where stay_code = '"&stay_code&"'"
		   Rs_stay.Open Sql, Dbconn, 1

		  if not rs_stay.eof then
             stay_name = rs_stay("stay_name")
			 stay_sido = rs_stay("stay_sido")
			 stay_gugun = rs_stay("stay_gugun")
			 stay_dong = rs_stay("stay_dong")
			 stay_addr = rs_stay("stay_addr")
		  end if
		  rs_stay.Close()
		end if
end if
title_line = " �λ���ī��-��Ÿ���� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
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
			function insert_off()
			{
				document.getElementById('into_tab').style.display = 'none';
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_card01.asp" method="post" name="frm">
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
								<th colspan="2" class="first">��������</th>
                                <% If rs("emp_type") = "1" then emp_type = "����" end if %>
								<% if rs("emp_type") = "2" then emp_type = "����" end if %>
								<% if rs("emp_type") = "3" then emp_type = "����" end if %>
								<% if rs("emp_type") = "9" then emp_type = "���" end if %>
								<td class="left"><%=emp_type%>&nbsp;</td>
								<th>����</th>
                                <td colspan="6" class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%>-<%=rs("emp_reside_place")%>&nbsp;</td>
								<th>���������</th>
								<td class="left"><%=rs("emp_yuncha_date")%>&nbsp;</td>
                            </tr>
							<tr>
								<th colspan="2" class="first">�������Կ���</th>
                            <%
							    if rs("emp_sawo_id") = "Y" then
								      sawo_id = "����"
								   else
								      sawo_id = "����"
							    end if
							%>
                                <td class="left"><%=sawo_id%>&nbsp;</td>
								<th>����������</th>
								<td class="left"><%=emp_sawo_date%>&nbsp;</td>
                                <th>��ֿ���</th>
                                <td colspan="2" class="left"><%=rs("emp_disabled")%>-<%=rs("emp_disab_grade")%>&nbsp;</td>
								<th>���</th>
								<td class="left"><%=rs("emp_hobby")%>&nbsp;</td>
                                <th>�������</th>
                                <td class="left"><%=emp_birthday%>(<%=rs("emp_birthday_id")%>)&nbsp;</td>
                 			</tr>
							<tr>
								<th colspan="2" class="first">����(�ּ�)</th>
								<td colspan="8" class="left"><%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></td>
                                <th>��󿬶���</th>
								<td class="left"><%=rs("emp_emergency_tel")%>&nbsp;</td>
                            </tr>
                            <tr>
								<th colspan="2" class="first">�Ǳٹ�����</th>
                                <td colspan="2" class="left"><%=stay_name%>&nbsp;</td>
                                <th >�Ǳٹ��� �ּ�</th>
								<td colspan="5" class="left"><%=stay_sido%>&nbsp;<%=stay_gugun%>&nbsp;<%=stay_dong%>&nbsp;<%=stay_addr%>&nbsp;</td>
                                <th >������</th>
                                <td class="left"><%=emp_end_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="10" class="left">�� ���� ���� ��</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_family_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','familyview','scrollbars=yes,width=800,height=400')">�� ���� ������</a>
                            </tr>
                            <tr>
                                <th colspan="3">����</th>
                                <th colspan="2">����</th>
                                <th colspan="2">�������</th>
                                <th colspan="2">����</th>
                                <th colspan="2">�ֹι�ȣ</th>
                                <th>���ſ���</th>
                            </tr>
                            <tr>
                                <td colspan="3" class="left"><%=fam_tab(1,1)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(1,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(1,3)%>(<%=fam_tab(1,4)%>)&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(1,5)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(1,6)%>-<%=fam_tab(1,7)%>&nbsp;</td>
                                <td class="left"><%=fam_tab(1,8)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=fam_tab(2,1)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(2,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(2,3)%>(<%=fam_tab(2,4)%>)&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(2,5)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=fam_tab(2,6)%>-<%=fam_tab(2,7)%>&nbsp;</td>
                                <td class="left"><%=fam_tab(2,8)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <th colspan="10" class="left">�� �߷� ���� ��</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_appoint_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','appointview','scrollbars=yes,width=1200,height=400')">�� �߷� ������</a>
                                </td>
                            </tr>
                            <tr>
				                <th rowspan="2" colspan="2" class="first">�߷���</th>
                                <th rowspan="2" scope="col">�߷ɱ���</th>
                                <th rowspan="2" scope="col">�߷�����</th>
                                <th colspan="3" scope="col">�߷���</th>
				                <th colspan="5" scope="col">�߷���</th>
			                </tr>
                            <tr>
                                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">����/å</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">����/å</th>
                                <th colspan="2" scope="col">���</th>
                            </tr>
                            <tr>
                                <td colspan="2" class="left"><%=app_tab(1,1)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,2)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,3)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,4)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,5)%>)<%=app_tab(1,6)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,7)%>-<%=app_tab(1,9)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,11)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,12)%>)<%=app_tab(1,13)%>&nbsp;</td>
                                <td class="left"><%=app_tab(1,14)%>-<%=app_tab(1,16)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=app_tab(1,18)%>-<%=app_tab(1,19)%><%=app_tab(1,17)%>&nbsp;<%=app_tab(1,20)%>&nbsp;<%=app_tab(1,21)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="2" class="left"><%=app_tab(2,1)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,2)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,3)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,4)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,5)%>)<%=app_tab(2,6)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,7)%>-<%=app_tab(2,9)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,11)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,12)%>)<%=app_tab(2,13)%>&nbsp;</td>
                                <td class="left"><%=app_tab(2,14)%>-<%=app_tab(2,16)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=app_tab(2,18)%>-<%=app_tab(2,19)%><%=app_tab(2,17)%>&nbsp;<%=app_tab(2,20)%>&nbsp;<%=app_tab(2,21)%>&nbsp;</td>
                             </tr>
                             <tr>
                                <th colspan="10" class="left">�� ���� ���� ��</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="pop_Window('insa_edu_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','eduview','scrollbars=yes,width=800,height=400')">�� ���� ������</a>
                                </td>
                             </tr>
                            <tr>
                                <th colspan="3">����&nbsp;������</th>
                                <th colspan="2">�������</th>
                                <th>����&nbsp;������No.</th>
                                <th colspan="2">����&nbsp;�Ⱓ</th>
                                <th colspan="4">����&nbsp;�ֿ�&nbsp;����</th>
                            </tr>
                            <tr>
                                <td colspan="3" class="left"><%=edu_tab(1,1)%>&nbsp;</td>
                                <td colspan="2"class="left"><%=edu_tab(1,2)%>&nbsp;</td>
                                <td class="left"><%=edu_tab(1,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=edu_tab(1,4)%> - <%=edu_tab(1,5)%>&nbsp;</td>
                                <td colspan="4" class="left"><%=edu_tab(1,6)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=edu_tab(2,1)%>&nbsp;</td>
                                <td colspan="2"class="left"><%=edu_tab(2,2)%>&nbsp;</td>
                                <td class="left"><%=edu_tab(2,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=edu_tab(2,4)%>(<%=edu_tab(2,5)%>)&nbsp;</td>
                                <td colspan="4" class="left"><%=edu_tab(2,6)%>&nbsp;</td>
                             </tr>
                              <tr>
                                <th colspan="12" class="left">�� ���� �ɷ� ��</th>
                             </tr>
                             <tr>
                                <th colspan="3">���б���</th>
                                <th colspan="2">��������</th>
                                <th colspan="2">����</th>
                                <th colspan="2">�޼�</th>
                                <th colspan="3">�����</th>
                            </tr>
                            <tr>
                                <td colspan="3" class="left"><%=lan_tab(1,1)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(1,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(1,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(1,4)%>&nbsp;</td>
                                <td colspan="3" class="left"><%=lan_tab(1,5)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=lan_tab(2,1)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(2,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(2,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(2,4)%>&nbsp;</td>
                                <td colspan="3" class="left"><%=lan_tab(2,5)%>&nbsp;</td>
                             </tr>
                            <tr>
                                <td colspan="3" class="left"><%=lan_tab(3,1)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(3,2)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(3,3)%>&nbsp;</td>
                                <td colspan="2" class="left"><%=lan_tab(3,4)%>&nbsp;</td>
                                <td colspan="3" class="left"><%=lan_tab(3,5)%>&nbsp;</td>
                             </tr>
                      <% end if %>
			    	  </tbody>
                    </table>
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >�ݱ�</a>&nbsp;&nbsp;
					</div>
                    <br>
        	</form>
		</div>
	</div>
	</body>
</html>
