<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim u_type, org_code, view_condi
Dim code_last
Dim org_level, org_date, org_empno, org_empname
Dim org_company, org_bonbu, org_saupbu, org_team, org_reside_place
Dim org_reside_company, org_cost_group, org_cost_center, owner_org
Dim owner_orgname, owner_empno, owner_empname, org_table_org
Dim tel_ddd, tel_no1, tel_no2, org_sido, org_gugun, org_dong, org_addr
Dim org_end_date, org_reg_date, org_reg_user, org_mod_date, org_mod_user
Dim title_line

u_type = Request("u_type")
org_code = Request("org_code")
view_condi = Request("view_condi")

code_last = ""
org_level = ""
org_name = ""
org_date = ""
org_end_date = ""
org_empno = ""
org_empname = ""
org_company = ""
org_bonbu = ""
org_saupbu = ""
org_team = ""
org_reside_place = ""
org_reside_company = ""
org_cost_group = ""
org_cost_center = ""
owner_org = ""
owner_orgname = ""
owner_empno = ""
owner_empname = ""
org_table_org = 0
tel_ddd = ""
tel_no1 = ""
tel_no2 = ""
org_sido = ""
org_gugun = ""
org_dong = ""
org_addr = ""
org_end_date = ""
org_reg_date = ""
org_reg_user = ""
org_mod_date = ""
org_mod_user = ""

'Set Rs_memb = Server.CreateObject("ADODB.Recordset")

'Set Rs_org = Server.CreateObject("ADODB.Recordset")
'Set Rs_tra = Server.CreateObject("ADODB.Recordset")
'Set Rs_owner = Server.CreateObject("ADODB.Recordset")
'Set Rs_max = Server.CreateObject("ADODB.Recordset")

title_line = " ���� ��� "

'���� ������ ���
If u_type = "U" Then
	Dim rs

	'Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	objBuilder.Append "SELECT org_level, org_name, org_date, org_end_date, org_empno, "
	objBuilder.Append "org_emp_name, org_company, org_bonbu, org_saupbu, org_team, "
	objBuilder.Append "org_reside_place, org_reside_company, org_cost_group, org_cost_center, "
	objBuilder.Append "org_owner_org, org_owner_empno, org_owner_empname, org_table_org, "
	objBuilder.Append "org_tel_ddd, org_tel_no1, org_tel_no2, org_sido, org_gugun, "
	objBuilder.Append "org_dong, org_addr, org_end_date, org_reg_date, "
	objBuilder.Append "org_reg_user, org_mod_date, org_mod_user, "
	objBuilder.Append "(SELECT org_name FROM emp_org_mst "
	objBuilder.Append "	WHERE org_code = eomt.org_owner_org) AS owner_orgname "
	objBuilder.Append "FROM emp_org_mst AS eomt "
	objBuilder.Append "WHERE org_code = '"&org_code&"'"

	Set rs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    org_level = rs("org_level")
    org_name = rs("org_name")
    org_date = rs("org_date")
	org_end_date = rs("org_end_date")
    org_empno = rs("org_empno")
    org_empname = rs("org_emp_name")
    org_company = rs("org_company")
    org_bonbu = rs("org_bonbu")
    org_saupbu = rs("org_saupbu")
    org_team = rs("org_team")
	org_reside_place = rs("org_reside_place")
	org_reside_company = rs("org_reside_company")
	org_cost_group = rs("org_cost_group")
	org_cost_center = rs("org_cost_center")
    owner_org = rs("org_owner_org")
    owner_empno = rs("org_owner_empno")
    owner_empname = rs("org_owner_empname")

	If rs("org_table_org") = "" Or IsNull(rs("org_table_org")) Then
		org_table_org = 0
	Else
		org_table_org = rs("org_table_org")
	End If

    tel_ddd = rs("org_tel_ddd")
    tel_no1 = rs("org_tel_no1")
    tel_no2 = rs("org_tel_no2")
	org_sido = rs("org_sido")
    org_gugun = rs("org_gugun")
    org_dong = rs("org_dong")
    org_addr = rs("org_addr")
    org_end_date = rs("org_end_date")
    org_reg_date = rs("org_reg_date")
	org_reg_user = rs("org_reg_user")
    org_mod_date = rs("org_mod_date")
    org_mod_user = rs("org_mod_user")
	owner_orgname = rs("owner_orgname")

	rs.Close() : Set rs = Nothing

	'Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
	'Set rs_owner=DbConn.Execute(Sql)

    'owner_orgname = rs_owner("org_name")
	'rs_owner.close()

	title_line = " ���� ���� "

Else
	Dim rs_max
	Dim max_seq

    'sql="select max(org_code) as max_seq from emp_org_mst "
	objBuilder.Append "SELECT MAX(org_code) AS max_seq FROM emp_org_mst "

	Set rs_max = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If IsNull(rs_max("max_seq")) Then
		code_last = "0001"
	Else
		max_seq = "000" + CStr((Int(rs_max("max_seq")) + 1))
		code_last = Right(max_seq, 4)
	End If

    rs_max.Close() : Set rs_max = Nothing

    org_code = code_last
End If
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
				return "0 1";
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=org_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=org_end_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

     		function chkfrm(){
				if(document.frm.org_code.value ==""){
					alert('�����ڵ带 �Է��ϼ���');
					frm.org_code.focus();
					return false;
				}

				if(document.frm.org_name.value ==""){
					alert('�������� �Է��ϼ���');
					frm.org_name.focus();
					return false;
				}else{
					if(document.frm.org_name.value.indexOf('�ѻ�') > -1){
					  alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
						frm.org_name.focus();
						return false;
					}
				}

				if(document.frm.org_date.value ==""){
					alert('������������ �Է��ϼ���');
					frm.org_date.focus();
					return false;
				}

				if($('#org_level').val() === "����" && ($('#org_name').val() !== '�濵����' && $('#org_name').val() !== '���������')){
					if(document.frm.org_empno.value ==""){
						alert('���������� �Է��ϼ���');
						frm.org_empno.focus();
						return false;
					}

					if(document.frm.org_empname.value ==""){
						alert('�����强���� �Է��ϼ���');
						frm.org_empname.focus();
						return false;
					}
				}

//				if(document.frm.org_cost_group.value ==""){
//					alert('��뼾Ÿ�׷��� �����ϼ���');
//					frm.org_cost_group.focus();
//					return false;}

				if(document.frm.org_cost_center.value ==""){
					alert('��뱸���� �����ϼ���');
					frm.org_cost_center.focus();
					return false;
				}

				if(document.frm.org_level.value !="ȸ��"){
					if(document.frm.owner_org.value ==""){
						alert('���������� �Է��ϼ���');
						frm.owner_org.focus();
						return false;
					}
				}

				if(document.frm.org_level.value =="����ó"){
					if(document.frm.org_cost_center.value =="����������"){
						if(document.frm.org_reside_place.value ==""){
							alert('����ó�� �Է��ϼ���');
							frm.org_reside_place.focus();
							return false;
						}else{
							if(document.frm.org_reside_place.value.indexOf('�ѻ�') > -1){
								alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
								frm.org_reside_place.focus();
								return false;
							}
  						}
					}
				}

				if(document.frm.org_cost_center.value =="����������"){
					if(document.frm.org_reside_company.value ==""){
						alert('����óȸ�縦 �Է��ϼ���');
						frm.org_reside_company.focus();
						return false;
					}else{
						if (document.frm.org_cost_center.value.indexOf('�ѻ�') > -1){
						  alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
								frm.org_cost_center.focus();
								return false;
						}
  					}
				}

				if(document.frm.org_level.value =="����ó"){
					if(document.frm.org_reside_company.value ==""){
						alert('����ó ȸ�縦 �����ϼ���');
						frm.org_reside_company.focus();
						return false;
					}
				}

				if(document.frm.org_level.value =="����ó"){
					if(document.frm.org_cost_center.value !="����������"){
						alert('��뱸�п� ���������� �����ϼ���');
						frm.org_cost_center.focus();
						return false;
					}
				}
//				if(document.frm.org_cost_group.value =="") {
//					alert('����óȸ��(�ŷ�ó)�� �׷��̸��� �����ϴ�.');
//					frm.org_reside_company.focus();
//					return false;}

				{
					a=confirm('�Է��Ͻðڽ��ϱ�?')
					if (a==true) {
						return true;
					}
					return false;
				}
			}

			function num_chk(txtObj){
				org_to = parseInt(document.frm.org_table_org.value.replace(/,/g,""));

				org_to = String(org_to);
				num_len = org_to.length;
				sil_len = num_len;
				org_to = String(org_to);

				if(org_to.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) org_to = org_to.substr(0,num_len -3) + "," + org_to.substr(num_len -3,3);
				if(sil_len > 6) org_to = org_to.substr(0,num_len -6) + "," + org_to.substr(num_len -6,3) + "," + org_to.substr(num_len -2,3);

				document.frm.org_table_org.value = org_to;

			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">

			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org_reg_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
						</colgroup>
						<tbody>
                            <tr>
                                <th class="first" style="background:#FFFFE6">ȸ��</th>
                                <td colspan="7" class="left" bgcolor="#FFFFE6">
					            <input name="view_condi" type="text" id="view_condi" size="20" value="<%=view_condi%>" readonly="true">
                                &nbsp;&nbsp;�� ����ó�� ��뱸���� ���������� �ΰ�� ������ �Է��� �ϼž� �մϴ�.
                                </td>
                            </tr>
							<tr>
								<th class="first">�����ڵ�</th>
                                <td class="left"><%=org_code%><input name="org_code" type="hidden" value="<%=org_code%>"></td>
                                <th>����&nbsp;����</th>
                                <td class="left">
                             <%
							 	Dim rs_etc
								'Sql="select * from emp_etc_code where emp_etc_type = '01' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name "
								objBuilder.Append "FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '01' "
								objBuilder.Append "ORDER BY emp_etc_code ASC "

								Set rs_etc = Server.CreateObject("ADODB.Recordset")
								rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
 							 %>
                                <select name="org_level" id="org_level" style="width:150px" value="<%=org_level%>">
                             <%
								Do Until rs_etc.EOF
 			  				 %>
                                <option value='<%=rs_etc("emp_etc_name")%>' <%If org_level = rs_etc("emp_etc_name") Then %>selected<% End If %>><%=rs_etc("emp_etc_name")%></option>
                 			<%
									rs_etc.MoveNext()
								Loop

								rs_etc.Close()
							%>
            					</select>
            					</td>
                                <th>������</th>
                                <td class="left"><input name="org_name" type="text" id="org_name" style="width:150px" value="<%=org_name%>" name="������" onKeyUp="checklength(this,20)"></td>
                                <th>����������</th>
                                <td class="left">
                                <input name="org_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;" value="<%=org_date%>" >
              					</td>
                             </tr>
                             <tr>
								<th class="first">��������</th>
                                <td class="left"><input name="org_empno" type="text" id="org_empno" size="7" readonly="true" value="<%=org_empno%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_emp_select.asp?gubun=<%="orgemp"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">������ã��</a>
                                </td>
                                <th>�����强��</th>
                                <td class="left">
                                <input name="org_empname" type="text" id="org_empname" size="10" readonly="true" value="<%=org_empname%>">
                                </td>
                                <th>�Ҽ�</th>
                                <td colspan="3" class="left">
                                <input name="org_company" type="text" id="org_company" style="width:100px" readonly="true" value="<%=org_company%>">
              					<input name="org_bonbu" type="text" id="org_bonbu" style="width:120px" readonly="true" value="<%=org_bonbu%>">
              					<input name="org_saupbu" type="text" id="org_saupbu" style="width:120px" readonly="true" value="<%=org_saupbu%>">
              					<input name="org_team" type="text" id="org_team" style="width:120px" readonly="true" value="<%=org_team%>">
                                </td>
                             </tr>
							<tr>
								<th class="first">���������ڵ�</th>
                                <td class="left"><input name="owner_org" type="text" id="owner_org" size="4" readonly="true" value="<%=owner_org%>">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="owner"%>&mg_level=<%=org_level%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=850,height=400')">��������ã��</a>
                                </td>
                                <th>����������</th>
                                <td class="left">
                                <input name="owner_orgname" type="text" id="owner_orgname" size="20" readonly="true" value="<%=owner_orgname%>"></td>
                                <th>����������</th>
                                <td class="left">
                                <input name="owner_empno" type="text" id="owner_empno" size="7" readonly="true" value="<%=owner_empno%>"></td>
                                <th>�����������</th>
                                <td class="left">
                                <input name="owner_empname" type="text" id="owner_empname" size="20" readonly="true" value="<%=owner_empname%>"></td>
                             </tr>
                             <tr>
								<th class="first">��ǥ��ȭ</th>
                                <td class="left"><input name="tel_ddd" type="text" id="tel_ddd" size="3" maxlength="3" value="<%=tel_ddd%>" >
								  -
                                    <input name="tel_no1" type="text" id="tel_no1" size="4" maxlength="4" value="<%=tel_no1%>" >
                                    -
                                <input name="tel_no2" type="text" id="tel_no2" size="4" maxlength="4" value="<%=tel_no2%>" ></td>
                                <th>���������</th>
                                <td class="left">
                                <input name="org_end_date" type="text" size="10" readonly="true" id="datepicker1" style="width:70px;" value="<%=org_end_date%>" >
              					</td>
                                <th>����ó</th>
                                <td class="left"><input name="org_reside_place" type="text" id="org_reside_place" style="width:150px" value="<%=org_reside_place%>">
                                </td>
                                <th class="first">����ó ȸ��</th>
								<td class="left"><input name="org_reside_company" type="text" id="org_reside_company" style="width:120px" readonly="true" value="<%=org_reside_company%>">
								<a href="#" class="btnType03" onClick="pop_Window('insa_trade_search.asp?gubun=<%="1"%>','tradesearch','scrollbars=yes,width=600,height=400')">ã��</a>
            					</td>
                             </tr>
                             <tr>
								<th class="first">�ּ�</th>
								<td colspan="5" class="left">
                                <input name="org_sido" type="text" id="org_sido" style="width:100px" readonly="true" value="<%=org_sido%>">
              					<input name="org_gugun" type="text" id="org_gugun" style="width:150px" readonly="true" value="<%=org_gugun%>">
              					<input name="org_dong" type="text" id="org_dong" style="width:150px" readonly="true" value="<%=org_dong%>">
              					<input name="org_addr" type="text" id="org_addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=org_addr%>">
              					<input name="org_zip" type="hidden" id="org_zip" value="">
                                <a href="#" class="btnType03" onClick="pop_Window('zipcode_search.asp?gubun=<%="org"%>','org_zip_select','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                                </td>
                                <th>��뼾Ÿ�׷�</th>
                                <td class="left"><input name="org_cost_group" type="text" id="org_cost_group" style="width:150px" readonly="true" value="<%=org_cost_group%>">

            					</td>
                              </tr>
                              <tr>
								<th class="first">�����ο�(T.O)</th>
								<td class="left">
                                <input name="org_table_org" type="text" id="org_table_org" style="width:90px;text-align:right" value="<%=formatnumber(org_table_org,0)%>" onKeyUp="num_chk(this);">
            					</td>
                                <th>�Է�����</th>
                                <td class="left">
                                <input name="org_reg_date" type="text" id="org_reg_date" style="width:150px" readonly="true" value="<%=org_reg_date%>">
                                </td>
                                <th>��������</th>
                                <td class="left">
                                <input name="org_mod_date" type="text" id="org_mod_date" style="width:150px" readonly="true" value="<%=org_mod_date%>">
                                </td>
                                <th>��뱸��</th>
                                <td class="left">
                              <%
								'Sql="select * from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
								objBuilder.Append "SELECT emp_etc_name "
								objBuilder.Append "FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '70' "
								objBuilder.Append "ORDER BY emp_etc_code ASC "

								rs_etc.Open objBuilder.ToString(), DBConn, 1
								objBuilder.Clear()
							  %>
								<select name="org_cost_center" id="org_cost_center" style="width:90px">
                                    <option value="" <% If org_cost_center = "" Then %>selected<% End If %>>����</option>
                			  <%
								Do Until rs_etc.EOF
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If org_cost_center = rs_etc("emp_etc_name") Then %>selected<% end If %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.Movenext()
								Loop

								rs_etc.Close() : Set rs_etc = Nothing
							  %>
                			     </select>
                                </td>
                              </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="mg_level" value="<%=org_level%>" ID="Hidden1">
				</form>
		</div>
	</div>
	</body>
</html>

