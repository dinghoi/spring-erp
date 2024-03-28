<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim car_no, u_type, title_line, view_condi
Dim car_old_no, car_name, car_year, oil_kind
Dim car_owner, insurance_company, insurance_date, insurance_amt
Dim buy_gubun, rental_company, car_reg_date, car_use_dept, car_company
Dim car_use, owner_emp_no, start_date, last_km, car_status, car_comment
Dim last_check_date, end_date, owner_emp_name, emp_grade, emp_org_name
Dim rsCar, org_level, emp_org_code

u_type = f_Request("u_type")
car_no = f_Request("car_no")

title_line = "���� ���"
view_condi = ""

'car_name = ""
'car_year = ""
'oil_kind = ""
'car_owner = ""
'insurance_company = ""
'insurance_date = ""
'insurance_amt = 0
'buy_gubun = "����"
'rental_company = ""
'car_reg_date = ""
'car_use_dept = ""
'car_company = ""
'car_use = ""
'owner_emp_no = ""
'owner_emp_name = ""
'emp_name = ""
'emp_grade = ""
'start_date = ""
'end_date = ""
'last_km = 0
'last_check_date = ""
'car_status = ""
'car_comment = ""

If u_type = "U" Then
	'sql = "select * from car_info where car_no = '" + car_no + "'"
	objBuilder.Append "SELECT cait.car_no, cait.car_name, cait.car_year, cait.oil_kind, cait.car_owner, "
	objBuilder.Append "	cait.insurance_company, cait.insurance_date, cait.insurance_amt, cait.buy_gubun, "
	objBuilder.Append "	cait.rental_company, cait.car_reg_date, cait.car_company, "
	objBuilder.Append "	cait.car_use, cait.owner_emp_no, cait.owner_emp_name, cait.start_date, end_date, "
	objBuilder.Append "	cait.last_km, cait.car_status, cait.car_comment, "
	objBuilder.Append "	IF(cait.last_check_date = '1900-01-01' OR cait.last_check_date = NULL, '', "
	objBuilder.Append "		cait.last_check_date) AS 'last_check_date', "
	objBuilder.Append "	IF(cait.end_date = '1900-01-01' OR cait.end_date = NULL, '', "
	objBuilder.Append "		cait.end_date) AS 'end_date', "
	objBuilder.Append "	IF(cait.car_use_dept = '' OR cait.car_use_dept = NULL, emtt.emp_org_name, "
	objBuilder.Append "		cait.car_use_dept) AS 'car_use_dept', "
	objBuilder.Append "	cait.owner_emp_no, "
	objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_org_name, emtt.emp_org_code "
	objBuilder.Append "FROM car_info AS cait "
	objBuilder.Append "INNER JOIN emp_master AS emtt ON cait.owner_emp_no = emtt.emp_no "
	objBuilder.Append "WHERE cait.car_no = '"&car_no&"' "

	Set rsCar = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    car_no = rsCar("car_no")
	car_old_no = rsCar("car_no")
    car_name = rsCar("car_name")

	car_year = rsCar("car_year")
    oil_kind = rsCar("oil_kind")
    car_owner = rsCar("car_owner")
    insurance_company = rsCar("insurance_company")
    insurance_date = rsCar("insurance_date")
    insurance_amt = rsCar("insurance_amt")
    buy_gubun = rsCar("buy_gubun")
    rental_company = rsCar("rental_company")
    car_reg_date = rsCar("car_reg_date")
    car_use_dept = rsCar("car_use_dept")
    car_company = rsCar("car_company")
    car_use = rsCar("car_use")
    owner_emp_no = rsCar("owner_emp_no")

    start_date = rsCar("start_date")
    last_km = rsCar("last_km")

    car_status = rsCar("car_status")
    car_comment = rsCar("car_comment")
	last_check_date = rsCar("last_check_date")
    end_date = rsCar("end_date")

	owner_emp_name = rsCar("emp_name")
	emp_grade = rsCar("emp_grade")
	emp_org_name = rsCar("emp_org_name")
	emp_org_code = rsCar("emp_org_code")

	rsCar.Close() : Set rsCar = Nothing

	title_line = "���� ����"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			/*$(document).ready(function(){
				update_view();
			});*/

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=car_reg_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
			});

			$(function(){
				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});

			$(function(){
				$( "#datepicker3" ).datepicker();
				$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
			});

			function frmcheck(type){
				if(formcheck(document.frm) && chkfrm(type)){
					document.frm.submit();
				}0
			}

			function chkfrm(type){
				if(document.frm.car_no.value =="" ){
					alert('������ȣ�� �Է��ϼ���');
					frm.car_no.focus();
					return false;
				}

				if(document.frm.car_name.value ==""){
					alert('������ �Է��ϼ���');
					frm.car_name.focus();
					return false;
				}

				if(document.frm.oil_kind.value ==""){
					alert('������ �����ϼ���');
					frm.oil_kind.focus();
					return false;
				}

				if(document.frm.car_owner.value ==""){
					alert('�����ڸ� �����ϼ���');
					frm.car_owner.focus();
					return false;
				}

				if(document.frm.car_reg_date.value ==""){
					alert('����������� �Է��ϼ���');
					frm.car_reg_date.focus();
					return false;
				}

				if(document.frm.owner_emp_no.value =="" ){
					alert('�����˻��� �ϼ���');
					frm.emp_name.focus();
					return false;
				}

				var message;

				if(type === 'U') message = "���� �Ͻðڽ��ϱ�?";
				else message = "���� �Ͻðڽ��ϱ�?";

				if(!confirm(message)) return false;
				else return true;
				/*
				{
					a=confirm('�Է��Ͻðڽ��ϱ�?');

					if (a==true) {
						return true;
					}
					return false;
				}*/
			}

			/*function update_view(){
				var c = document.frm.u_type.value;

				if(c == 'U'){
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}*/

			function num_chk(txtObj){
				lst_km = parseInt(document.frm.last_km.value.replace(/,/g,""));
				lst_km = String(lst_km);
				num_len = lst_km.length;
				sil_len = num_len;
				lst_km = String(lst_km);

				if(lst_km.substr(0,1) == "-") sil_len = num_len - 1;

				if(sil_len > 3){
					lst_km = lst_km.substr(0,num_len -3) + "," + lst_km.substr(num_len -3,3);
				}

				if(sil_len > 6){
					lst_km = lst_km.substr(0,num_len -6) + "," + lst_km.substr(num_len -6,3) + "," + lst_km.substr(num_len -2,3);
				}

				document.frm.last_km.value = lst_km;

				if(txtObj.value.length >= 2){
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}

				if(txtObj.value.length<5){
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}

				var num = txtObj.value;

				if(num == "--" ||  num == "." ){
					num = "";
				}

				if(num != "" ){
					temp=new String(num);

					if(temp.length<1) return "";

					// ����ó��
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";

					// �Ҽ�������ó��
					dpoint=temp.search(/\./);

					if(dpoint>0){
						// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
						dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
						temp=temp.substr(0,dpoint);
					}else dpointVa="";

					// �����ܹ̿��� ����
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);

					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);

					if(temp.length<4) return minus+temp+dpointVa;

					buf="";

					while(true){
						if(temp.length<3){
							buf=temp+buf;
							break;
						}

						buf=","+temp.substr(temp.length-3)+buf;
						temp=temp.substr(0, temp.length-3);
					}

					if(buf.substr(0,1)==",") buf=buf.substr(1);

					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";
			}

			function delcheck(){
				a = confirm('���� �����Ͻðڽ��ϱ�?');

				if(a==true){
					document.frm.method = "post";
					document.frm.action = "/insa/insa_car_info_del_ok.asp";
					document.frm.submit();

					return true;
				}
				return false;
			}

			//���� �˻�[����ȣ_20210721]
			function carOrgSearch(level, condi){
				var url = '/insa/insa_org_select.asp';
				var pop_name = '���� �˻�';
				var features = 'scrollbars=yes,width=850,height=400';
				var param;

				param = '?gubun=car&mg_level='+level+'&view_condi='+condi;

				url += param;

				pop_Window(url, pop_name, features);
			}

			//������ ���� �˻�[����ȣ_20210721]
			function carEmpSearch(condi){
				var url = '/insa/insa_emp_select.asp';
				var pop_name = '���� �˻�';
				var features = 'scrollbars=yes,width=600,height=400';
				var param;

				param = '?gubun=car&view_condi='+condi;

				url += param;

				pop_Window(url, pop_name, features);
			}
        </script>
	</head>
	<BODY>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_car_info_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
                                <th class="first">������ȣ</th>
								<td class="left">
									<input name="car_no" type="text" value="<%=car_no%>" style="width:150px" onKeyUp="checklength(this,20)" <%If u_type = "U" Then%>readonly<%End If%>>
								</td>
								<th>����</th>
								<td class="left">
									<input name="car_name" type="text" value="<%=car_name%>" style="width:150px" onKeyUp="checklength(this,30)">
								</td>
							</tr>
                           	<tr>
								<th class="first">��������</th>
								<td colspan="3" class="left">
									<input name="car_year" type="text" value="<%=car_year%>" style="width:70px" id="datepicker3">
								</td>
							</tr>
							<tr>
								<th class="first">����</th>
								<td class="left">
									<select name="oil_kind" id="oil_kind" style="width:150px">
										<option value="">����</option>
										<option value="�ֹ���" <%If oil_kind = "�ֹ���" Then %>selected<%End If %>>�ֹ���</option>
										<option value="����" <%If oil_kind = "����" then %>selected<%End If %>>����</option>
										<option value="����" <%If oil_kind = "����" then %>selected<%End If %>>����</option>
									</select>
                                </td>
								<th>����</th>
								<td class="left">
									<select name="car_owner" id="car_owner" style="width:150px">
										<option value="">����</option>
										<option value="ȸ��" <%If car_owner = "ȸ��" Then %>selected<%End If %>>ȸ��</option>
										<option value="����" <%If car_owner = "����" Then %>selected<%End If %>>����</option>
									</select>
								</td>
							</tr>
							<tr>
								<th class="first">���ű���</th>
								<td class="left">
									<input type="radio" name="buy_gubun" value="����" <%if buy_gubun = "����" Then %>checked<%End If %> style="width:40px" id="Radio1">����
									<input type="radio" name="buy_gubun" value="����" <%if buy_gubun = "����" Then %>checked<%End If %> style="width:40px" id="Radio2">����
									<input type="radio" name="buy_gubun" value="��Ʈ" <%if buy_gubun = "��Ʈ" Then %>checked<%End If %> style="width:40px" id="Radio2">��Ʈ
                                </td>
								<th>��Ʈȸ��</th>
                                <td class="left">
									<input name="rental_company" type="text" value="<%=rental_company%>" style="width:150px" onKeyUp="checklength(this,30)">
								</td>
							</tr>
							<tr>
								<th class="first">�Ҽ�ȸ��</th>
								<td class="left">
								<%
								Call SelectEmpOrgList("car_company", "car_company", "width:150px;", car_company)
								%>
                                </td>
								<th>���������</th>
								<td class="left"><input name="car_reg_date" type="text" value="<%=car_reg_date%>" style="width:70px" id="datepicker"></td>
							</tr>
							<tr>
								<th class="first">�뵵</th>
								<td class="left">
									<input name="car_use" type="text" value="<%=car_use%>" style="width:150px" onKeyUp="checklength(this,10)">
								</td>
								<th>���μ�</th>
								<td class="left">
									<input name="car_use_dept" type="text" id="car_use_dept" style="width:80px" value="<%=car_use_dept%>" readonly="true">
									<a href="#" class="btnType03" onClick="carOrgSearch('<%=org_level%>', '<%=view_condi%>');">
									�μ�ã��</a>
                                </td>
							</tr>
							<tr>
								<th class="first">������</th>
								<td colspan="3" class="left">
									<input name="emp_name" type="text" id="emp_name" style="width:80px" value="<%=owner_emp_name%>" readonly="true">
									<input name="emp_grade" type="text" id="emp_grade" style="width:80px" value="<%=emp_grade%>" readonly="true">
									<input name="owner_emp_no" type="text" id="owner_emp_no" style="width:80px" value="<%=owner_emp_no%>" readonly="true">
								<%If u_type = "" Then %>
									<a href="#" class="btnType03" onclick="carEmpSearch('<%=view_condi%>');">�����˻�</a>
								<%End If %>
									<input name="emp_company" type="hidden" id="emp_company" value="<%=emp_company%>">
									<input name="emp_org_code" type="hidden" id="emp_org_code" value="<%=emp_org_code%>">
									<input name="emp_org_name" type="hidden" id="emp_org_name" value="<%=emp_org_name%>">
								</td>
							</tr>
							<tr>
								<th class="first">��������</th>
								<td class="left">
									<input name="car_status" type="text" value="<%=car_status%>" style="width:150px" onKeyUp="checklength(this,20);">
								 </td>
								<th>��������</th>
								<td class="left">
									<input name="car_comment" type="text" value="<%=car_comment%>" style="width:170px" onKeyUp="checklength(this,50);">
								</td>
							</tr>
                        	<tr>
								<th class="first">������km</th>
								<td class="left">
									<input name="last_km" type="text" id="last_km" style="width:70px;text-align:right" value="<%=FormatNumber(last_km, 0)%>" onKeyUp="num_chk(this);">
								</td>
								<th>�����˻���</th>
                                <td class="left">
									<input name="last_check_date" type="text" value="<%=last_check_date%>" style="width:70px" id="datepicker1">
								</td>
							</tr>
                        	<tr>
								<th class="first">ó������</th>
								<td colspan="3" class="left">
									<input name="end_date" type="text" value="<%=end_date%>" style="width:70px" id="datepicker2">
								</td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01">
						<input type="button" value="<%If u_type = "U" Then%>����<%Else%>����<%End If%>" onclick="javascript:frmcheck('<%=u_type%>');" />
					</span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:toclose();"></span>
				<%If u_type = "U" And InsaCarDelYn = "Y" Then%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:delcheck();"></span>
				<%End If%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" />
                <input type="hidden" name="car_old_no" value="<%=car_old_no%>" />
			</form>
		</div>
	</body>
</html>

