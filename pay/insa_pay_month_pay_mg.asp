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
Dim month_tab(24, 2)
Dim quarter_tab(8, 2)
Dim year_tab(3, 2)

Dim be_pg, view_condi, owner_view, condi, pmg_yymm, to_date
Dim page, curr_dd, from_date, datYear, datMonth
Dim datLastDay, exec_LastDay, give_date
Dim st_in_date, curr_mm, i, j, cal_quarter, cal_month, view_month
Dim cal_year, pgsize, start_page, stpage, title_line
Dim rsCount, totRecord, total_page, rs_etc
Dim emp_payend_date, emp_payend_yn, emp_payend
Dim whereSql, pg_url

Dim rsPay, arrPay
Dim emp_grade, emp_position, emp_first_date, emp_in_date, emp_org_name
Dim pmg_base_pay, pmg_give_total, de_deduct_total, pmg_emp_no, de_emp_no
Dim dt_ck, pmg_curr_pay, emp_org_code, emp_name
Dim rs_emp

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")
owner_view = f_Request("owner_view")
pmg_yymm = f_Request("pmg_yymm")
to_date = f_Request("to_date")

be_pg = "/pay/insa_pay_month_pay_mg.asp"

If view_condi = "" Then
	view_condi = "���̿�"
	condi = ""
	owner_view = "C"
	curr_dd = CStr(DatePart("d", Now()))
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	pmg_yymm = Mid(CStr(from_date), 1, 4)&Mid(CStr(from_date), 6, 2)

	'�ſ� ���� ���ϱ�
	datYear = Mid(CStr(pmg_yymm), 1, 4)
	datMonth = Mid(CStr(pmg_yymm), 5, 2)

	If datMonth = 4 Or datMonth = 6 Or datMonth = 9 Or datMonth = 11 Then  '4�� 6�� 9�� 11���̸� �������� 30��
		datLastDay = 30
	ElseIf datMonth = 2 And Not (datYear Mod 4) = 0 Then  '2���̰�  �⵵�� 4�� ���� ���� 0�� �ƴϸ� 28��
		datLastDay = 28
	ElseIf datMonth = 2 And (datYear Mod 4) = 0 Then '���� ���
		If (datYear Mod 100) = 0 Then
			If (datYear Mod 400) = 0 Then
				datLastDay=29
			Else
				datLastDay=28
			End If
		Else
			datLastDay=29
		End If
	Else
		datLastDay=31
	End If

	exec_LastDay = datLastDay
'   to_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + cstr(exec_LastDay)

   to_date = ""
End If

give_date = to_date '������

'��� �Ի����� 15�� �����̸� ��� �޿������
'st_es_date = Mid(CStr(pmg_yymm), 1, 4)&"-"&Mid(CStr(pmg_yymm), 5, 2)&"-"&"01"
st_in_date = Mid(CStr(pmg_yymm), 1, 4)&"-"&Mid(CStr(pmg_yymm), 5, 2)&"-"&"16"

' �ֱ�3���⵵ ���̺�� ����
year_tab(3, 1) = Mid(Now(), 1, 4)
year_tab(3, 2) = CStr(year_tab(3, 1))&"��"
year_tab(2, 1) = CInt(Mid(Now(), 1, 4)) - 1
year_tab(2, 2) = CStr(year_tab(2, 1))&"��"
year_tab(1, 1) = CInt(Mid(Now(), 1, 4)) - 2
year_tab(1, 2) = CStr(year_tab(1, 1))&"��"

' �б� ���̺� ����
curr_mm = Mid(Now(), 6, 2)

If curr_mm > 0 And curr_mm < 4 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"1"
End If

If curr_mm > 3 And curr_mm < 7 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"2"
End If

If curr_mm > 6 And curr_mm < 10 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"3"
End If

If curr_mm > 9 And curr_mm < 13 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4))&"4"
End If

quarter_tab(8, 2) = CStr(Mid(quarter_tab(8, 1), 1, 4))&"�� "&CStr(Mid(quarter_tab(8, 1), 5, 1))&"/4�б�"

For i = 7 To 1 Step -1
	cal_quarter = CInt(quarter_tab(i+1, 1)) - 1

	If CStr(Mid(cal_quarter, 5, 1)) = "0" Then
		quarter_tab(i, 1) = CStr(CInt(Mid(cal_quarter, 1, 4)) - 1)&"4"
	Else
		quarter_tab(i, 1) = cal_quarter
	End If

	quarter_tab(i, 2) = CStr(Mid(quarter_tab(i, 1), 1, 4))&"�� "&CStr(Mid(quarter_tab(i, 1), 5, 1))&"/4�б�"
Next

' ��� ���̺����
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()), 1, 4)&Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month, 1, 4)&"�� "&Mid(cal_month, 5, 2)&"��"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = CStr(Int(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(Int(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year&"12"
	End If

	view_month = Mid(cal_month, 1, 4)&"�� "&Mid(cal_month, 5, 2)&"��"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next

'�޿����� ���� ��ȸ
objBuilder.Append "SELECT emp_payend_date, emp_payend_yn FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_code = '9999';"

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_payend_date = rs_etc("emp_payend_date")
emp_payend_yn = rs_etc("emp_payend_yn")

rs_etc.Close() : Set rs_etc = Nothing

If pmg_yymm > emp_payend_date Then
       emp_payend = "N"
Else
	   emp_payend = "Y"
End If

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&owner_view="&owner_view&"&condi="&condi&"&pmg_yymm="&pmg_yymm&"&to_date="&to_date

If condi = "" Then
	whereSql = "AND emtt.emp_no < '900000' "
Else
	If owner_view = "C" Then
		whereSql = "AND emtt.emp_name LIKE '%"&condi&"%' "
	Else
		whereSql = "AND emtt.emp_no = '"&condi&"' "
	End If
End If

'ī��Ʈ ��ȸ
objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt "
objBuilder.Append "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date >= '"&st_in_date&"') "
objBuilder.Append "	AND emtt.emp_in_date < '"&st_in_date&"' "
objBuilder.Append "	AND emtt.emp_company = '"&view_condi&"' "
objBuilder.Append "	AND emtt.emp_pay_id <> '5' "
objBuilder.Append whereSql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

totRecord = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

title_line = " �޿��ڷ� �Է� "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�޿����� �ý���</title>
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

		    /*$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%'=from_date%>" );
			});*/

			$(function(){
				$("#datepicker1").datepicker();
				$("#datepicker1").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker1").datepicker("setDate", "<%=to_date%>");
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

			function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				var result = confirm('�����Ͻðڽ��ϱ�?');

				if(result == true){
					return true;
				}
				return false;
			}

			function pay_month_del(val, val2, val3, val4){
				if(!confirm("���� �����Ͻðڽ��ϱ� ?")) return;

				var frm = document.frm;

				document.frm.in_empno1.value = val;
				document.frm.in_name1.value = val2;
				document.frm.pmg_yymm1.value = val3;
				document.frm.view_condi1.value = val4;

				document.frm.action = "/pay/insa_pay_month_del.asp";
				document.frm.submit();
			}

			function pay_month_tax_cal(val, val2, val3, val4, val5){
				if(!confirm("�޿� ���ݰ��ó���� �Ͻðڽ��ϱ�?")) return;

				var frm = document.frm;

				document.frm.pmg_yymm1.value = document.getElementById(val).value;
				document.frm.view_condi1.value = document.getElementById(val2).value;
				document.frm.in_empno1.value = val3;
				document.frm.in_name1.value = val4;
				document.frm.owner_view1.value = val5;

				document.frm.action = "/pay/insa_pay_month_tax_calcu.asp";
				document.frm.submit();
            }
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/pay/insa_pay_month_pay_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>�� �˻���</dt>
                        <dd>
                            <p>
                             <strong>ȸ�� : </strong>
                              <%
							  	Dim rs_org
								' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = 'ȸ��' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

								Set rs_org = DBConn.Execute(objBuilder.ToString())
	                            objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px;">
							  <%
								Do Until rs_org.EOF
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") Then %>selected<%End If %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.MoveNext()
								Loop
								rs_org.Close() : Set rs_org = Nothing
							  %>
            					</select>
                                </label>
                                <label>
								<strong>�ͼӳ�� : </strong>
                                <select name="pmg_yymm" id="pmg_yymm" value="<%=pmg_yymm%>" style="width:90px;">
                                    <%For i = 24 to 1 Step -1%>
                                    <option value="<%=month_tab(i, 1)%>" <%If pmg_yymm = month_tab(i, 1) Then %>selected<%End If%>><%=month_tab(i, 2)%></option>
                                    <%Next	%>
                                 </select>
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px;" id="datepicker1">
								</label>
								<label>
                                <input name="owner_view" type="radio" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px">���
                                <input name="owner_view" type="radio" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px">����
                                </label>
							<strong>���� : </strong>
								<label>
        						<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px;text-align:left;">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="3%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�����Ի���</th>
                                <th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col">�⺻��</th>
                                <th scope="col">���޾װ�</th>
                                <th scope="col">�����װ�</th>
                                <th scope="col">�������޾�</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <th scope="col">�޿�</th>
                                <th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						'�޿� ���� ��ȸ
						objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_grade, emtt.emp_position, emtt.emp_first_date, "
						objBuilder.Append "	emtt.emp_in_date, emtt.emp_org_name, "
						objBuilder.Append "	pmgt.pmg_base_pay, pmgt.pmg_give_total, "
						objBuilder.Append "	pmdt.de_deduct_total, "
						objBuilder.Append "	pmgt.pmg_emp_no, pmdt.de_emp_no, emtt.emp_org_code, emtt.emp_company "
						objBUilder.Append "FROM emp_master AS emtt "
						objBUilder.Append "LEFT OUTER JOIN pay_month_give AS pmgt ON emtt.emp_no = pmgt.pmg_emp_no "
						objBuilder.Append "	AND emtt.emp_company = pmgt.pmg_company AND pmgt.pmg_id = '1' "
						objBuilder.Append "	AND pmgt.pmg_yymm = '"&pmg_yymm&"' "
						objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
						objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company AND pmdt.de_id = '1' "
						objBuilder.Append "	AND pmdt.de_yymm = '"&pmg_yymm&"' "
						objBuilder.Append "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01' OR emtt.emp_end_date >= '"&st_in_date&"') "
						objBuilder.Append "	AND emtt.emp_company = '"&view_condi&"' "
						objBuilder.Append "	AND emp_pay_id <> '5' AND emp_in_date < '"&st_in_date&"' "
						objBuilder.Append whereSql
						objBuilder.Append "ORDER BY emp_in_date, emp_no ASC "
						objBuilder.Append "LIMIT "& stpage & "," &pgsize&";"

						Set rsPay = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsPay.EOF Then
							arrPay = rsPay.getRows()
						End If
						rsPay.Close() : Set rsPay = Nothing

						If IsArray(arrPay) Then
							For i = LBound(arrPay) To UBound(arrPay, 2)
								emp_no = arrPay(0, i)
								emp_name = arrPay(1, i)
								emp_grade = arrPay(2, i)
								emp_position = arrPay(3, i)
								emp_first_date = arrPay(4, i)
								emp_in_date = arrPay(5, i)
								emp_org_name = arrPay(6, i)
								pmg_base_pay = CLng(f_toString(arrPay(7, i), 0))
								pmg_give_total = CLng(f_toString(arrPay(8, i), 0))
								de_deduct_total = CLng(f_toString(arrPay(9, i), 0))
								pmg_emp_no = f_toString(arrPay(10, i), "")
								de_emp_no = f_toString(arrPay(11, i), "")
								emp_org_code = arrPay(12, i)
								emp_company = arrPay(13, i)

								dt_ck = "1"

								If pmg_emp_no = "" Then
									dt_ck = "0"
								Else
									dt_ck = "1"
								End If

								pmg_curr_pay = pmg_give_total - de_deduct_total
	           			%>
							<tr>
								<td class="first"><%=emp_no%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=emp_no%>','�λ� ���� ī��','scrollbars=yes,width=1250,height=650')"><%=emp_name%></a>
								</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=emp_first_date%>&nbsp;</td>
                                <td><%=emp_in_date%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_base_pay, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_give_total, 0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(de_deduct_total,0)%>&nbsp;</td>
                                <td class="right">
								<%If pmg_curr_pay > 0 Then%>
									<a href="#" onClick="pop_Window('/person/insa_pay_person_view.asp?emp_no=<%=emp_no%>&pmg_yymm=<%=pmg_yymm%>&pmg_company=<%=emp_company%>','�޿� �󼼳���','scrollbars=yes,width=750,height=700')"><%=FormatNumber(pmg_curr_pay, 0)%></a>&nbsp;
								<%
								Else
									Response.Write "0"
								End If%>
								</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(emp_org_code)
								%>
								</td>
								<td>
								<%If emp_payend = "N" Then%>
									<a href="#" onClick="pop_Window('/pay/insa_pay_month_give_add.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&pmg_yymm=<%=pmg_yymm%>&view_condi=<%=view_condi%>&u_type=<%If dt_ck <> "0" Then %>U<%End If%>','�޿� ����/���� �Է�','scrollbars=yes,width=750,height=700')">�Է�</a>
								<%End If%>
								</td>
								<td>
								<%If emp_payend = "N" And dt_ck = "1"  Then %>
									<a href="#" onClick="pay_month_del('<%=emp_no%>', '<%=emp_name%>', '<%=pmg_yymm%>', '<%=view_condi%>');return false;">����</a>
								<%End If %>
								</td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='14' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
					<td width="15%">
					<!--�޿�������Ȳ�� �ߺ� ������� �ּ� ó��[����ȣ_20220405]<div class="btnCenter">
                    <% 'insa_excel_pay_month_ledger %>
						<a href="/insa_excel_pay_transe_list.asp?view_condi=<%'=view_condi%>&pmg_yymm=<%'=pmg_yymm%>&to_date=<%'=to_date%>&owner_view=<%'=owner_view%>" class="btnType04">�����ٿ�ε�</a>
					</div>
					-->
                  	</td>
				    <td>
                    <%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, totRecord, pgsize)
					%>
                    </td>
					<td width="15%">
						<div class="btnRight">
						<%
						Dim v_emp_no, v_emp_company, v_emp_name

						If emp_payend = "N" Then
							If owner_view = "T" Then
								  v_emp_no = condi

								  objBuilder.Append "SELECT emp_name, emp_company FROM emp_master WHERE emp_no = '"&v_emp_no&"';"

								  Set rs_emp = DbConn.Execute(objBuilder.ToString())
								  objBuilder.Clear()

								  If Not rs_emp.EOF Then
									   v_emp_company = rs_emp("emp_company")
									   v_emp_name = rs_emp("emp_name")
								  End If
								  rs_emp.Close() : Set rs_emp = Nothing
						%>
							<a href="#" onClick="pop_Window('/pay/insa_pay_month_give_add.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&pmg_yymm=<%=pmg_yymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>','�޿� ���� �Է�','scrollbars=yes,width=750,height=700')" class="btnType04">�޿������Է�</a>

						<!--
						��� �̻������ �ּ� ó��[����ȣ_20220405]
						<a href="#" onClick="pay_month_tax_cal('pmg_yymm','view_condi','<%'=emp_no%>','<%'=emp_name%>','<%'=owner_view%>');return false;" class="btnType04">�޿� ���ݰ�� ó��</a>
						-->
						<%
							'Else
								'If condi = "" Then
						%>
							<!--��� �̻����� �ּ� ó��[����ȣ_20220405]
							<a href="#" onClick="pay_month_tax_cal('pmg_yymm','view_condi','in_empno','in_name','<%'=owner_view%>');return false;" class="btnType04">�޿� ���ݰ�� �ϰ�ó��</a>
							-->
						<%
								'End If
							End If
						End If
						DBConn.Close() : Set DBConn = Nothing
						%>
						</div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="view_condi1" value="<%=view_condi%>"/>
                  <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>"/>
                  <input type="hidden" name="in_empno1" value="<%=emp_no%>"/>
                  <input type="hidden" name="in_name1" value="<%=emp_name%>"/>
                  <input type="hidden" name="owner_view1" value="<%=owner_view%>"/>
        	</form>
		</div>
	</div>
	</body>
</html>