<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_emp_wetax_report.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "���̿��������"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
	
	tax_meals_no = 0	
	tax_car_no = 0	
	tax_meals_yes = 0	
	tax_car_yes = 0	
	
end if

' ��� ���̺����
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	  pay_count = pay_count + 1
							  
	  pmg_date = rs("pmg_date")
	  pmg_base_pay = rs("pmg_base_pay")
	  pmg_meals_pay = rs("pmg_meals_pay")
	  pmg_postage_pay = rs("pmg_postage_pay")
	  pmg_re_pay = rs("pmg_re_pay")
	  pmg_overtime_pay = rs("pmg_overtime_pay")
	  pmg_car_pay = rs("pmg_car_pay")
	  pmg_position_pay = rs("pmg_position_pay")
  	  pmg_custom_pay = rs("pmg_custom_pay")
	  pmg_job_pay = rs("pmg_job_pay")
	  pmg_job_support = rs("pmg_job_support")
	  pmg_jisa_pay = rs("pmg_jisa_pay")
	  pmg_long_pay = rs("pmg_long_pay")
	  pmg_disabled_pay = rs("pmg_disabled_pay")

	  meals_pay = pmg_meals_pay
	  car_pay = pmg_car_pay
	  meals_tax_pay = 0
	  meals_taxno_pay = 0
	  car_tax_pay = 0
	  car_taxno_pay = 0
	  
	  if  meals_pay > 100000 then
	         meals_tax_pay = meals_pay - 100000
	         tax_meals_yes = tax_meals_yes + (meals_pay - 100000)
			 meals_taxno_pay = 100000
			 tax_meals_no= tax_meals_no + 100000
		  else	 
		     meals_taxno_pay = meals_pay
			 tax_meals_no= tax_meals_no + meals_pay
	  end if
  	  if car_pay > 200000 then
	         car_tax_pay = car_pay - 200000
			 tax_car_yes = tax_car_yes + (car_pay - 200000)
			 car_taxno_pay = 200000
			 tax_car_no =  tax_car_no + 200000
		 else
			 tax_car_no =  tax_car_no + car_pay
			 car_taxno_pay = car_pay
	  end if
	  
	  pmg_tax_yes = 0
	  pmg_tax_no = 0
	  
	  pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	  pmg_tax_no = meals_taxno_pay + car_taxno_pay
	  
	  sum_tax_yes = sum_tax_yes + pmg_tax_yes
	  sum_tax_no = sum_tax_no + pmg_tax_no

	rs.movenext()
loop
rs.close()

pmg_date = curr_date '�׽�Ʈ

sum_give_tot = sum_tax_yes + sum_tax_no

month_person_pay = int(sum_tax_yes / pay_count) '�Ű�� ������޿���
deduct_14 = month_person_pay * (pay_count - pay_count) '������
income_pay15 = sum_tax_yes - deduct_14 '�����ǥ
income_tax16 = int(income_pay15 * (0.5 / 100)) '���⼼��
add_tax1 = 0
add_tax2 = 0
add_tax17 = 0
tax_hap = income_tax16 + add_tax17

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = " �������һ���Ҽ�(���漼) "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "5 1";
			}
		</script>
		<script type="text/javascript">
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
					alert ("�Ҽ��� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_tax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_emp_wetax_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>�� �˻���</dt>
                        <dd>
                            <p>
                             <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = 'ȸ��' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
								<strong>�ͼӳ�� : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <h3 class="stit">*�������� �ֹμ�&nbsp;&nbsp;</h3>
				<div class="gView">
                    <table width="175%" border="0" cellpadding="0" cellspacing="0">
				        <tr>
                            <td width="50%" class="left">&nbsp;&nbsp;&nbsp;&nbsp;�ͼӳ��:&nbsp;<%=mid(pmg_yymm,1,4)%>��&nbsp;<%=mid(pmg_yymm,5,2)%>����</td>
                            <td width="50%" class="right">�޿�������:&nbsp;<%=pmg_date%></td>
                        </tr>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
                            <col width="20%" >
						</colgroup>
						<thead>
							<tr>
				                <th rowspan="2" class="first" scope="col">����</th>
                                <th rowspan="2" scope="col">8.������ο�</th>
				                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">����ǥ�ؾ�</th>
			                </tr>
                            <tr>
							    <th scope="col" style=" border-left:1px solid #e3e3e3;">10.�������ܱ޿�</th>
								<th scope="col">11.�����޿�</th>  
								<th scope="col">9.�����ޱ޿���</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first" style="background:#f8f8f8;">��������</td>
                                <td class="right"><%=formatnumber(pay_count,0)%>&nbsp;��&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_tax_no,0)%>&nbsp;��&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_tax_yes,0)%>&nbsp;��&nbsp;</td>
                                <td class="right"><%=formatnumber(sum_give_tot,0)%>&nbsp;��&nbsp;</td>
							</tr>
						</tbody>
					</table>
                <h3 class="stit">�߼ұ�� ��������� �ش�Ǵ� �߼ұ���� ������(���漼�� ��84����5�� �ش��ϴ°��)</h3>    
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="40%" >
                            <col width="30%" >
                            <col width="30%" >
						</colgroup>
						<thead>
                            <tr>
							    <th class="first" scope="col">12.�������� ����� ��������</th>
								<th scope="col">13.�Ű�� ������޿���(11/8)</th>  
								<th scope="col">14.������(13*(8-12))</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                                <td class="right"><%=formatnumber(pay_count,0)%>&nbsp;��&nbsp;</td>
                                <td class="right"><%=formatnumber(month_person_pay,0)%>&nbsp;��&nbsp;</td>
                                <td class="right"><%=formatnumber(deduct_14,0)%>&nbsp;��&nbsp;</td>
							</tr>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							    <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">15.�����ǥ(11-14)</th>  
								<td class="right"><%=formatnumber(income_pay15,0)%>&nbsp;��&nbsp;</td>
								<th scope="col" style=" border-bottom:1px solid #e3e3e3;">16.���⼼��(15*0.5%)</th> 
                                <td class="right"><%=formatnumber(income_tax16,0)%>&nbsp;��&nbsp;</td>
							</tr>
                            <tr>
							    <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">���κҼ��ǰ��꼼</th>  
								<td class="right"><%=formatnumber(add_tax1,0)%>&nbsp;��&nbsp;</td>
								<th scope="col" style=" border-bottom:1px solid #e3e3e3;">�Ű�Ҽ��ǰ��꼼</th> 
                                <td class="right"><%=formatnumber(add_tax2,0)%>&nbsp;��&nbsp;</td>
							</tr>
                            <tr>
							    <th class="first" scope="col">17.���꼼</th>  
								<td class="right"><%=formatnumber(add_tax17,0)%>&nbsp;��&nbsp;</td>
								<th scope="col">�Ű����հ�(16+17)</th> 
                                <td class="right"><%=formatnumber(tax_hap,0)%>&nbsp;��&nbsp;</td>
							</tr>
						</thead>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_empwetax_report.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_emp_wetax_print.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>','insa_pay_emp_wetax_pop','scrollbars=yes,width=1250,height=600')" class="btnType04">���μ�</a>
					</div>                  
                    </td> 
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

