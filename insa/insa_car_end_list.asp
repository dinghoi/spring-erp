<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### �۾� ����
'===================================================
' ����ȣ_20210722 :
'	- �ű� ������ �ۼ� �� �ڵ� ����
'	- AS���, ���·� ���� ó��(���� ��� ��ɸ� ������ ���� ���� ������ ����, ��� �������� �Ϲ� ���� ���� �����)

'===================================================
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
Dim be_pg, page, owner_view, field_check, field_view
Dim pgsize, start_page, stpage, title_line, str_param
Dim base_sql, owner_sql, field_sql, rsCount, total_record
Dim total_page, order_sql, rsCar

page = f_Request("page")
owner_view = f_Request("owner_view")
field_check = f_Request("field_check")
field_view = f_Request("field_view")

title_line = "ó�� ���� ����"
be_pg = "/insa/insa_car_end_list.asp"

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

str_param = "&owner_view="&owner_view&"&field_check="&field_check&"&field_view="&field_view

If owner_view = "" Then
	owner_view = "T"
	field_check = "total"
End If

If field_check = "total" Then
	field_view = ""
End If

If owner_view = "C" Then
	owner_sql = "WHERE cait.car_owner = 'ȸ��' "
ElseIf owner_view = "P" Then
	owner_sql = "WHERE cait.car_owner = '����' "
Else
  	owner_sql = "WHERE (cait.car_owner = '����' OR cait.car_owner = 'ȸ��') "
End If

field_sql = "AND (cait.end_date <> '' AND cait.end_date <> '1900-01-01') "
If field_check <> "total" Then
	field_sql = field_sql & "AND (" & field_check & " LIKE '%" & field_view & "%') "
End If

'List Count
objBuilder.Append "SELECT COUNT(*) FROM car_info AS cait "
objBuilder.Append owner_sql & field_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

order_sql = " ORDER BY cait.car_owner DESC, cait.car_no DESC"

objBuilder.Append "SELECT cait.owner_emp_no, "
objBuilder.Append "	IFNULL(cait.owner_emp_name, emtt.emp_name) AS 'emp_name', "
objBuilder.Append "	cait.last_check_date, cait.end_date, cait.car_year, cait.car_no, cait.car_name, "
objBuilder.Append "	cait.car_reg_date, cait.oil_kind, cait.car_owner, cait.buy_gubun, cait.rental_company, "
objBuilder.Append "	cait.insurance_amt, cait.insurance_date, cait.last_km, cait.car_use_dept "
objBuilder.Append "FROM car_info AS cait "
objBuilder.Append "INNER JOIN emp_master AS emtt ON cait.owner_emp_no = emtt.emp_no "
objBuilder.Append owner_sql & field_sql & order_sql & " LIMIT "& stpage & "," & pgsize

Set rsCar = Server.CreateObject("ADODB.RecordSet")
rsCar.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.field_check.value == ""){
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			//���� �ٿ�ε�[����ȣ_20210721]
			function carEndInfoExcel(o_view, f_check, f_view){
				var url = '/insa/excel/insa_excel_car_end_info.asp';
				var param;

				param = '?owner_view='+o_view+'&field_check='+f_check+'&field_view='+f_view;

				location.href = url + param;
			}

			//���� ���� �˾�[����ȣ_20210720]
			function carInfoView(car_no, car_name, car_year, car_reg_date, oil_kind){
				var url = '/insa/insa_car_info_view.asp';
				var pop_name = '���� ����';
				var features = 'scrollbars=yes,width=900,height=600';
				var param;

				param = '?car_no='+car_no+'&car_name='+car_name+'&car_year='+car_year+'&car_reg_date='+car_reg_date;
				param += '&oil_kind='+oil_kind;

				url += param;

				pop_Window(url, pop_name, features);
			}

			//���� ���谡�� ��Ȳ �˾�[����ȣ_20210720]
			function carInsView(car_no, car_name, car_year, car_reg_date){
				var url = '/insa/insa_car_ins_view.asp';
				var pop_name = '���� ���谡�� ��Ȳ';
				var features = 'scrollbars=yes,width=1200,height=600';
				var param;

				param = '?car_no='+car_no+'&car_name='+car_name+'&car_year='+car_year+'&car_reg_date='+car_reg_date;

				url += param;

				pop_Window(url, pop_name, features);
			}

			//���� ������ ��Ȳ �˾�[����ȣ_20210721]
			function carDrvUserView(car_no, car_name, car_year, car_reg_date){
				var url = '/insa/insa_car_drvuser_view.asp';
				var pop_name = '���� ������ ��Ȳ';
				var features = 'scrollbars=yes,width=750,height=600';
				var param;

				param = '?car_no='+car_no+'&car_name='+car_name+'&car_year='+car_year+'&car_reg_date='+car_reg_date;

				url += param;

				pop_Window(url, pop_name, features);
			}

			//���� ���� ��Ȳ �˾�[����ȣ_20210721]
			function carDriveView(car_no, car_name, car_year, car_reg_date){
				var url = '/insa/insa_car_drv_view.asp';
				var pop_name = '���� ���� ��Ȳ';
				var features = 'scrollbars=yes,width=1250,height=600';
				var param;

				param = '?car_no='+car_no+'&car_name='+car_name+'&car_year='+car_year+'&car_reg_date='+car_reg_date;

				url += param;

				pop_Window(url, pop_name, features);
			}

			//���� ���� ���/����[����ȣ_20210721]
			function carInfoInit(car_no, type){
				var url = '/insa/insa_car_info_add.asp';
				var pop_name = '���� ���';
				var features = 'scrollbars=yes,width=750,height=450';
				var param;

				param = '?car_no='+car_no+'&u_type='+type;

				url += param;

				pop_Window(url, pop_name, features);
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
                                <label>
									<input name="owner_view" type="radio" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px">�Ѱ�
									<input name="owner_view" type="radio" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px">ȸ��
									<input name="owner_view" type="radio" value="P" <%If owner_view = "P" Then %>checked<%End If %> style="width:25px">����
                                </label>
                                <label>
									<strong>�ʵ�����</strong>
									<select name="field_check" id="field_check" style="width:100px">
									  <option value="total" <%If field_check = "total" Then %>selected<%End If %>>��ü</option>
									  <option value="buy_gubun" <%If field_check = "buy_gubun" Then %>selected<%End If %>>���ű���</option>
									  <option value="owner_emp_name" <%If field_check = "owner_emp_name" Then %>selected<%End If %>>������</option>
									  <option value="oil_kind" <%If field_check = "oil_kind" Then %>selected<%End If %>>����</option>
									  <option value="car_no" <%If field_check = "car_no" Then %>selected<%End If %>>������ȣ</option>
									</select>
									<input name="field_view" type="text" value="<%=field_view%>" style="width:100px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="5%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">������ȣ</th>
								<th scope="col">����/����</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">����<br>����</th>
								<th scope="col">���������</th>
								<th scope="col">ó������</th>
								<th scope="col">�����</th>
								<th scope="col">����Ⱓ</th>
								<th scope="col">������</th>
								<th scope="col">����KM</th>
								<th scope="col">�����˻���</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim owner_emp_name, owner_emp_no, last_check_date, end_date, car_year
						Dim car_no, car_name, car_reg_date, oil_kind

						Do Until rsCar.EOF
							owner_emp_name = rsCar("emp_name")
							owner_emp_no = rsCar("owner_emp_no")
							car_no = rsCar("car_no")
							car_name = rsCar("car_name")
							car_reg_date = rsCar("car_reg_date")
							oil_kind = rsCar("oil_kind")

							If rsCar("last_check_date") = "1900-01-01"  Then
	                            last_check_date = ""
							Else
								last_check_date = rsCar("last_check_date")
	                        End If

	                        If rsCar("end_date") = "1900-01-01" Then
								end_date = ""
							Else
							    end_date = rsCar("end_date")
	                        End If

							If rsCar("car_year") = "1900-01-01" Then
								car_year = ""
							Else
								car_year = rsCar("car_year")
	                        End If
						%>
							<tr>
								<td class="first">
									<a href="#" onclick="carInfoView('<%=car_no%>', '<%=car_name%>', '<%=car_year%>', '<%=car_reg_date%>', '<%=oil_kind%>');"><%=rsCar("car_no")%></a>&nbsp;
                                </td>
								<td class="left"><%=rsCar("car_name")%>(<%=car_year%>)</td>
								<td><%=rsCar("oil_kind")%></td>
								<td><%=rsCar("car_owner")%></td>
								<td><%=rsCar("buy_gubun")%>&nbsp;<%=rsCar("rental_company")%></td>
								<td><%=rsCar("car_reg_date")%>&nbsp;</td>
								<td><%=rsCar("end_date")%>&nbsp;</td>
                                <td class="right">
									<a href="#" onclick="carInsView('<%=car_no%>', '<%=car_name%>', '<%=car_year%>', '<%=car_reg_date%>' );"><%=FormatNumber(rsCar("insurance_amt"),0)%></a>
									&nbsp;
								</td>
                                <td><%=rsCar("insurance_date")%>&nbsp;</td>
                                <td>
									<a href="#" onclick="carDrvUserView('<%=car_no%>', '<%=car_name%>', '<%=car_year%>', '<%=car_reg_date%>' );"><%=owner_emp_name%>(<%=rsCar("owner_emp_no")%>)</a>
									&nbsp;
								</td>
                                <td class="right">
									<a href="#" onclick="carDriveView('<%=car_no%>', '<%=car_name%>', '<%=car_year%>', '<%=car_reg_date%>');"><%=FormatNumber(rsCar("last_km"), 0)%></a>
									&nbsp;
								</td>
								<td><%=last_check_date%>&nbsp;</td>
								<td>
									<a href="#" onclick="carInfoInit('<%=car_no%>', 'U');">����</a>
								</td>
							</tr>
						<%
							rsCar.MoveNext()
						Loop
						rsCar.Close() : Set rsCar = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
					<a href="#" onclick="carEndInfoExcel('<%=owner_view%>', '<%=field_check%>', '<%=field_view%>');" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
					<%
					'page navigator[����ȣ_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)
					%>
					</td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->