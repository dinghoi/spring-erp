<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim view_condi, owner_view, ck_sw, title_line
Dim from_date, to_date
Dim condi, page, pgsize, stpage
Dim totRecord, total_page, start_page
Dim rsCount, be_pg, family_empno, family_seq
Dim rsApp

be_pg = "/insa/insa_reward_punish_mg.asp"

view_condi = request("view_condi")
condi = request("condi")
page = Request("page")

ck_sw = Request("ck_sw")

if ck_sw = "n" then
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
	condi = request.form("condi")
else
	owner_view=request("owner_view")
	view_condi = request("view_condi")
	condi = request("condi")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	condi = "��ü"
	ck_sw = "n"
end if

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

if view_condi <> "" Then
	objBuilder.Append "SELECT COUNT(*) FROM emp_appoint "

	if condi = "��ü" then
		if owner_view = "T" then
			'Sql = "SELECT count(*) FROM emp_appoint where (app_empno = '"+view_condi+"') and (app_id = '����߷�' or app_id = '¡��߷�')"
			objBuilder.Append "WHERE app_empno = '"&view_condi&"' AND (app_id = '����߷�' OR app_id = '¡��߷�') "
		else
			'Sql = "SELECT count(*) FROM emp_appoint where (app_emp_name like '%"+view_condi+"%') and (app_id = '����߷�' or app_id = '¡��߷�')"
			objBuilder.Append "WHERE app_emp_name LIKE '%"&view_condi&"%' AND (app_id = '����߷�' OR app_id = '¡��߷�') "
		end if
	else
		if owner_view = "T" then
			'Sql = "SELECT count(*) FROM emp_appoint where app_empno = '"+view_condi+"' and app_id = '"+condi+"'"
			objBuilder.Append "WHERE app_empno = '"&view_condi&"' AND app_id = '"&condi&"' "
		else
			'Sql = "SELECT count(*) FROM emp_appoint where app_emp_name like '%"+view_condi+"%' and app_id = '"+condi+"'"
			objBuilder.Append "WHERE app_emp_name LIKE '%"&view_condi&"%' AND app_id = '"&condi&"' "
		end if
	end If

	Set rsCount = Dbconn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	totRecord = cint(rsCount(0)) 'Result.RecordCount
end if

IF totRecord mod pgsize = 0 THEN
	total_page = int(totRecord / pgsize) 'Result.PageCount
ELSE
	total_page = int((totRecord / pgsize) + 1)
END If

if view_condi <> "" Then
	objBuilder.Append "SELECT app_empno, app_empno, app_date, app_id, app_id_type, app_reward, "
	objBuilder.Append "	app_start_date, app_finish_date, app_comment, app_to_grade, app_to_position, "
	objBuilder.Append "	app_to_company, app_to_org, app_to_orgcode "
	objBuilder.Append "FROM emp_appoint "

	if condi = "��ü" Then
		if owner_view = "T" then
			'Sql = "SELECT * FROM emp_appoint where (app_empno = '"+view_condi+"') and (app_id = '����߷�' or app_id = '¡��߷�') ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
			objBuilder.Append "WHERE app_empno = '"&view_condi&"' AND (app_id = '����߷�' OR app_id = '¡��߷�') "
		else
			'Sql = "SELECT * FROM emp_appoint where (app_emp_name like '%"+view_condi+"%') and (app_id = '����߷�' or app_id = '¡��߷�') ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
			objBuilder.Append "WHERE app_emp_name LIKE '%"&view_condi&"%' AND (app_id = '����߷�' OR app_id = '¡��߷�') "
		end if
	else
		if owner_view = "T" then
			'Sql = "SELECT * FROM emp_appoint where app_empno = '"+view_condi+"' and app_id = '"+condi+"' ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
			objBuilder.Append "WHERE app_empno = '"&view_condi&"' AND app_id = '"&condi&"' "
		else
			'Sql = "SELECT * FROM emp_appoint where app_emp_name like '%"+view_condi+"%' and app_id = '"+condi+"' ORDER BY app_empno,app_date,app_seq ASC limit "& stpage & "," &pgsize
			objBuilder.Append "WHERE app_emp_name LIKE '%"&view_condi&"%' AND app_id = '"&condi&"' "
		end if
	end If
	objBuilder.Append "ORDER BY app_empno, app_date, app_seq ASC LIMIT "& stpage & "," &pgsize

	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	rsApp.Open objBuilder.ToString(), Dbconn, 1
	objBuilder.Clear()
end if
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html lang="ko">
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

			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("������ �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			function reward_punish_del(val, val2, val3) {

            if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			document.frm.app_empno.value = val;
			document.frm.app_seq.value = val2;
			document.frm.app_emp_name.value = val3;

            document.frm.action = "/insa/insa_reward_punish_del.asp";
            document.frm.submit();
            }
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_reward_punish_mg.asp?ck_sw=n" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>������ �˻���</dt>
                        <dd>
                            <p>
                                <label>
                            <strong>��� : </strong>
                                <select name="condi" id="condi" value="<%=condi%>" style="width:100px">
                                  <option value="��ü" <%If condi = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value="����߷�" <%If condi = "����߷�" then %>selected<% end if %>>����߷�</option>
                                  <option value="¡��߷�" <%If condi = "¡��߷�" then %>selected<% end if %>>¡��߷�</option>
                                </select>
                                </label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">���
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">����
                                </label>
							<strong>���� : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
                            <col width="7%" >
                            <col width="10%" >
                            <col width="6%" >
							<col width="10%" >
							<col width="13%" >
                            <col width="*" >
                            <col width="22%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>���</th>
                                <th>����</th>
                                <th>���Ҽ�</th>
                                <th>�������</th>
                                <th>�������</th>
                                <th>¡��Ⱓ</th>
                                <th>�������</th>
                                <th>����/��å �� �Ҽ�</th>
                                <th>���</th>
                                <th>����</th>
                                <th>���</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						Dim rs_emp, emp_name, emp_bonbu, emp_saupbu, emp_team
						Dim emp_org_code, emp_org_name, emp_job

						if  view_condi <> "" then
							do until rsApp.eof
						      app_empno = rsApp("app_empno")

							  'Sql = "SELECT * FROM emp_master where emp_no = '"&app_empno&"'"
							  objBuilder.Append "SELECT emtt.emp_company, emtt.emp_name, emtt.emp_bonbu, emtt.emp_saupbu, "
							  objBuilder.Append "	emtt.emp_team, emtt.emp_org_code, emtt.emp_org_name, "
							  objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_name "
							  objBuilder.Append "FROM emp_master AS emtt "
							  objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
							  objBuilder.Append "WHERE emtt.emp_no = '"&app_empno&"' "

							  Set rs_emp = DBConn.Execute(objBuilder.ToString())
							  objBuilder.Clear()

							  if not Rs_emp.eof then
                                   emp_company = rs_emp("org_company")
								   emp_name = rs_emp("emp_name")
								   emp_job = rs_emp("emp_job")
                                   emp_bonbu = rs_emp("org_bonbu")
                                   emp_saupbu = rs_emp("org_saupbu")
                                   emp_team = rs_emp("org_team")
                                   emp_org_code = rs_emp("emp_org_code")
                                   emp_org_name = rs_emp("org_name")
							  end if
							  rs_emp.close() : Set rs_emp = Nothing
						%>
							<tr>
                              <td><%=rsApp("app_empno")%>&nbsp;</td>
                              <td><%=emp_name%>(<%=emp_job%>)&nbsp;</td>
                              <td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td><%=rsApp("app_date")%>&nbsp;</td>
                        <% if rsApp("app_id") = "����߷�" then %>
						      <td class="left">(����)<%=rsApp("app_id_type")%>&nbsp;</td>
                              <td class="left">&nbsp;</td>
                              <td class="left"><%=rsApp("app_reward")%>&nbsp;</td>
                        <%    elseif rsApp("app_id") = "¡��߷�" then %>
                              <td class="left">(¡��)<%=rsApp("app_id_type")%>&nbsp;</td>
                              <td class="left"><%=rsApp("app_start_date")%>��<%=rsApp("app_finish_date")%>&nbsp;</td>
                              <td class="left"><%=rsApp("app_comment")%>&nbsp;</td>
                        <% end if %>
                              <td class="left"><%=rsApp("app_to_grade")%>-<%=rsApp("app_to_position")%>(<%=rsApp("app_to_company")%>&nbsp;<%=rsApp("app_to_org")%>(<%=rsApp("app_to_orgcode")%>)</td>


                        <% if user_id = "900002" Or user_id = "102592" then %>
                              <td >
                              <a href="#" onClick="pop_Window('/insa/insa_reward_punish_add.asp?app_empno=<%=rsApp("app_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')">���</a></td>
							  <td><a href="#" onClick="pop_Window('/insa/insa_reward_punish_add.asp?app_empno=<%=rsApp("app_empno")%>&app_seq=<%=rsApp("app_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')">����</a></td>
                         <% if insa_grade = "0" then %>
                              <td>
                              <a href="#" onClick="reward_punish_del('<%=rsApp("app_empno")%>', '<%=rsApp("app_seq")%>', '<%=emp_name%>');return false;">����</a></td>
                         <%     else %>
                              <td>&nbsp;</td>
                         <% end if %>
                         <% end if %>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
							</tr>
						<%
							rsApp.movenext()
						loop
						rsApp.close() : Set rsApp = Nothing

						end if
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                  <div id="paging">
                        <a href = "<%=be_pg%>?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="<%=be_pg%>insa_reward_punish_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="<%=be_pg%>insa_reward_punish_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="<%=be_pg%>insa_reward_punish_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[����]</a> <a href="<%=be_pg%>insa_reward_punish_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<% if user_id = "900002" or user_id = "102592" then %>
                    <a href="#" onClick="pop_Window('/insa/insa_reward_punish_add.asp?family_empno=<%=view_condi%>&emp_name=<%=emp_name%>','insa_reward_punish_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">������׵��</a>
                    <% end if %>
					</div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="family_empno" value="<%=family_empno%>" ID="Hidden1">
                  <input type="hidden" name="family_seq" value="<%=family_seq%>" ID="Hidden1">
                  <input type="hidden" name="family_name" value="<%=emp_name%>" ID="Hidden1">
			</form>
		</div>
	</div>
	</body>
</html>

