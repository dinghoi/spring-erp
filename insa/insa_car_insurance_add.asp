<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### �۾� ����
'===================================================
' ����ȣ_20210721 :
'	- �ű� ������ �ۼ� �� �ڵ� ����
'	- ������ ���� �������� �߰��� �����ϰ� �ۼ�, ���� ���� �������� nkp���� �������� ����(����:�λ� ������ ����)

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
Dim car_no, rsCarInfo
Dim owner_emp_name, owner_emp_no, last_check_date, end_date
Dim car_year, car_name, car_reg_date, car_use_dept, oil_kind
Dim car_owner, title_line

car_no = f_Request("car_no")

title_line = "���� ������"

objBuilder.Append "SELECT cait.owner_emp_no, "
objBuilder.Append "	IFNULL(cait.owner_emp_name, emtt.emp_name) AS 'owner_emp_name', "
objBuilder.Append "	cait.last_check_date, cait.end_date, cait.car_year, cait.car_name, "
objBuilder.Append "	cait.car_reg_date, cait.oil_kind, cait.car_owner, cait.car_use_dept "
objBuilder.Append "FROM car_info AS cait "
objBuilder.Append "INNER JOIN emp_master AS emtt ON cait.owner_emp_no = emtt.emp_no "
objBuilder.Append "WHERE cait.car_no = '"&car_no&"' "

Set rsCarInfo = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

owner_emp_name = ""
owner_emp_no = rsCarInfo("owner_emp_no")

If rsCarInfo("last_check_date") = "1900-01-01"  Then
	 last_check_date = ""
Else
	 last_check_date = rsCarInfo("last_check_date")
End If

If rsCarInfo("end_date") = "1900-01-01" Then
	 end_date = ""
Else
	 end_date = rsCarInfo("end_date")
End If

If rsCarInfo("car_year") = "1900-01-01" Then
	 car_year = ""
Else
	 car_year = rsCarInfo("car_year")
End If

car_name = rsCarInfo("car_name")
car_reg_date = rsCarInfo("car_reg_date")
car_use_dept = rsCarInfo("car_use_dept")
oil_kind = rsCarInfo("oil_kind")
car_owner = rsCarInfo("car_owner")

rsCarInfo.Close() : Set rsCarInfo = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			$(function(){
				$("#datepicker").datepicker();
				$("#datepicker").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker").datepicker("setDate", "");
			});

			$(function(){
				$("#datepicker1").datepicker();
				$("#datepicker1").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker1").datepicker("setDate", "");
			});

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.ins_date.value ==""){
					alert('���谡������ �����ϼ���');
					frm.ins_date.focus();
					return false;
				}

				if(document.frm.ins_company.value =="" ){
					alert('����ȸ�縦 �Է��ϼ���');
					frm.ins_company.focus();
					return false;
				}

				if(document.frm.ins_amount.value ==""){
					alert('���谡�Ծ��� �Է��ϼ���');
					frm.ins_amount.focus();
					return false;
				}

				if(document.frm.ins_last_date.value ==""){
					alert('���踸������ �����ϼ���');
					frm.ins_last_date.focus();
					return false;
				}

				if(document.frm.ins_date.value > document.frm.ins_last_date.value){
					alert('���踸������ ���谡���Ϻ��� �����ϴ�');
					frm.ins_last_date.focus();
					return false;
				}

				if(document.frm.ins_man1.value ==""){
					alert('����1 �������� �Է��ϼ���');
					frm.ins_man1.focus();
					return false;
				}

				if(document.frm.ins_man2.value ==""){
					alert('����2 �������� �Է��ϼ���');
					frm.ins_man2.focus();
					return false;
				}

				if(document.frm.ins_object.value =="" ){
					alert('�빰 �������� �Է��ϼ���');
					frm.ins_object.focus();
					return false;
				}

				if(document.frm.ins_self.value =="" ){
					alert('�ڱ� �������� �Է��ϼ���');
					frm.ins_self.focus();
					return false;
				}

				if(document.frm.ins_injury.value =="" ){
					alert('������ �������� �Է��ϼ���');
					frm.ins_injury.focus();
					return false;
				}

				if(document.frm.ins_self_car.value =="" ){
					alert('���� �������� �Է��ϼ���');
					frm.ins_self_car.focus();
					return false;
				}

				if(!confirm('�Է� �Ͻðڽ��ϱ�?'))return false;
				else return true;
			}

			function num_chk(txtObj){
				ins_amt = parseInt(document.frm.ins_amount.value.replace(/,/g,""));
				ins_amt = String(ins_amt);
				num_len = ins_amt.length;
				sil_len = num_len;
				ins_amt = String(ins_amt);
				if(ins_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) ins_amt = ins_amt.substr(0,num_len -3) + "," + ins_amt.substr(num_len -3,3);
				if(sil_len > 6) ins_amt = ins_amt.substr(0,num_len -6) + "," + ins_amt.substr(num_len -6,3) + "," + ins_amt.substr(num_len -2,3);

				document.frm.ins_amount.value = ins_amt;

				if(txtObj.value.length >= 2){
					if(txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}

				if (txtObj.value.length<5){
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}

				var num = txtObj.value;

				if(num == "--" ||  num == "." ) num = "";
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
						if(temp.length<3) { buf=temp+buf; break; }

						buf=","+temp.substr(temp.length-3)+buf;
						temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);

					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";
			}
			/*
			function update_view(){
				var c = document.frm.u_type.value;

				if(c == 'U'){
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			*/
        </script>
	</head>
	<!--<body onload="update_view()">-->
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_car_insurance_save.asp" method="post" name="frm">
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
								<th class="first" style="background:#FFFFE6">������ȣ</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_no%>&nbsp;
                                <input name="car_no" type="hidden" value="<%=car_no%>" style="width:150px" readonly="true"></td>
								<th style="background:#FFFFE6">����</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_name%>&nbsp;
                                <input name="car_name" type="hidden" value="<%=car_name%>" style="width:150px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first" style="background:#FFFFE6">��������</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_year%>&nbsp;
                                <input name="car_year" type="hidden" value="<%=car_year%>" style="width:70px" readonly="true"></td>
                                <th style="background:#FFFFE6">���������</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_reg_date%>&nbsp;
                                <input name="car_reg_date" type="hidden" value="<%=car_reg_date%>" style="width:70px" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first">���谡����</th>
								<td class="left"><input name="ins_date" type="text" value="" style="width:70px" id="datepicker"></td>
                                <th class="first">���踸����</th>
								<td class="left"><input name="ins_last_date" type="text" value="" style="width:70px" id="datepicker1"></td>
							</tr>
							<tr>
								<th class="first">����ȸ��</th>
								<td class="left">
                                <input name="ins_company" type="text" value="" style="width:150px" onKeyUp="checklength(this,30)">
                                </td>
								<th>���谡�Ծ�</th>
								<td class="left">
                                <input name="ins_amount" type="text" id="ins_amount" style="width:80px;text-align:right" value="0" onKeyUp="num_chk(this);">
                                </td>
							</tr>
							<tr>
								<th class="first">����1</th>
								<td class="left">
                                <input name="ins_man1" type="text" value="" style="width:150px" onKeyUp="checklength(this,30)"></td>
								<th>����2</th>
                                <td class="left">
								<input name="ins_man2" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">�빰</th>
								<td class="left">
                                <input name="ins_object" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>�ڱ�</th>
                                <td class="left">
								<input name="ins_self" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">������</th>
								<td class="left">
                                <input name="ins_injury" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>����</th>
                                <td class="left">
								<input name="ins_self_car" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">����</th>
								<td class="left">
                                <input name="ins_age" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>����⵿</th>
                                <td class="left">
								<input name="ins_scramble" type="text" value="" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                             <tr>
								<th class="first">��೻��<br>��������</th>
                                <td class="left">
									<input type="radio" name="ins_contract_yn" value="Y" checked />����
              						<input name="ins_contract_yn" type="radio" value="N" />������
                                </td>
                                <th>���</th>
								<td class="left"><input name="ins_comment" type="text" value="" style="width:200px" /></td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01">
						<input type="button" value="����" onclick="javascript:frmcheck();" />
					</span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:toclose();"></span>
                </div>
				<!--<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="ins_old_date" value="<%=ins_old_date%>" ID="Hidden1">-->
			</form>
		</div>
	</body>
</html>

