<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim abc,filenm
dim month_tab(24,2)
Set abc = Server.CreateObject("ABCUpload4.XForm")
abc.AbsolutePath = True
abc.Overwrite = true
abc.MaxUploadSize = 1024*1024*50

pay_company = abc("pay_company")
pay_month   = abc("pay_month")
give_date   = abc("give_date")
file_type   = abc("file_type")

if ck_sw = "y" then
	pay_company = request("pay_company")
	pay_month=request("pay_month")
end if

if pay_company = "" then
	ck_sw = "y"
  else
  	ck_sw = "n"
end if
	
if pay_company = "" then
    pay_company = "���̿��������"
    curr_dd = cstr(datepart("d",now))
    give_date = mid(cstr(now()),1,10)
    from_date = mid(cstr(now()-curr_dd+1),1,10)
    pay_month = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
end if
	
' ��� ���̺����
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
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
	
	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")	
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_org = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
	Set Rs_give = Server.CreateObject("ADODB.Recordset")
  Set Rs_dct = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	If ck_sw = "n" Then
		Set filenm = abc("att_file")(1)
		
		path = Server.MapPath ("/pay_file")
		filename = filenm.safeFileName
		fileType = mid(filename,inStrRev(filename,".")+1)
		file_name = pay_company + "_" + pay_month + "_�޿�" + give_date
		
'		save_path = path & "\" & filename
		save_path = path & "\" & file_name&"."&fileType

		if fileType = "xls" or fileType = "xlk" then
			file_type = "Y"
			filenm.save save_path
		
			objFile = save_path
	'		objFile = Request.form("att_file")
	'		objFile = SERVER.MapPath("att_file")
	'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
	'		response.write(objFile)
			
			cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
			rs.Open "select * from [1:10000]",cn,"0"
				
			rowcount=-1
			xgr = rs.getrows
			rowcount = ubound(xgr,2)
			fldcount = rs.fields.count
			tot_cnt = rowcount + 1
		else
			objFile = "none"
			rowcount=-1
			file_type = "N"
		end if		  
	else
		objFile = "none"
		rowcount=-1
	end if
	title_line = "�޿� �ڷ� ���ε�"

etc_code = "9999"	
sql = "select * from emp_etc_code where emp_etc_code = '" + etc_code + "'"
Rs_etc.Open Sql, Dbconn, 1
'Response.write Sql&"<br>"
emp_payend_date = Rs_etc("emp_payend_date")
emp_payend_yn = Rs_etc("emp_payend_yn")

Rs_etc.close()

'Response.write pay_month & "<br>"
'Response.write emp_payend_date & "<br>"


if pay_month > emp_payend_date then
	emp_payend = "N"
else   
	emp_payend = "Y"
end if   	
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
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
            
            // �˻� ��ư Ŭ��!!
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.pay_company.value == "") {
					alert ("ȸ�縦 �����ϼ���");
					return false;
				}	
				if (document.frm.pay_month.value == "") {
					alert ("�ͼӳ���� �����ϼ���");
					return false;
				}	
				if (document.frm.att_file.value == "") {
					alert ("���ε� ���� ������ �����ϼ���");
					return false;
				}	
				return true;
            }
            
            // �޿� upload ��ư Ŭ��!!
			function frm1check () {
				if (chkfrm1()) {
					document.frm1.submit ();
				}
			}
			
            function chkfrm1() 
            {
				if (confirm('DB�� ���ε� �Ͻðڽ��ϱ�?')==true) {
					return true;
				}
				return false;
			}
			
            function pay_month_updel(val, val2) 
            {
				if (!confirm("�޿� Upload�ڷḦ ���� �Ͻðڽ��ϱ� ?")) return;
                var frm = document.frm;
                
                document.frm.pay_month1.value   = document.getElementById(val).value;
                document.frm.pay_company1.value = document.getElementById(val2).value;
                    
                document.frm.action = "insa_pay_month_up_del.asp";
                document.frm.submit();
            }	
		</script>

</head>
<body>
	<div id="wrap">			
	<!--#include virtual = "/include/insa_pay_header.asp" -->
	<!--#include virtual = "/include/insa_pay_menu.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_month_up.asp" method="post" name="frm" enctype="multipart/form-data">
					<fieldset class="srch">
						<legend>��ȸ����</legend>
						<dl>
							<dt>���ε峻��</dt>
							<dd>
								<p>
									<label>
										<strong>ȸ��: </strong>
										<%
                                        ' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
                                        Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
                                        rs_org.Open Sql, Dbconn, 1	
                                        %>
                                        <select name="pay_company" id="pay_company" type="text" style="width:110px">
                                            <option value="">����</option>
                                            <% 
                                            do until rs_org.eof 
                                                %>
                                                <option value='<%=rs_org("org_name")%>' <%If pay_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                                                <%
                                                rs_org.movenext()  
                                            loop 
                                            rs_org.Close()
                                            %>
                                        </select>
                                    </label>
                                    <label>
                                        <strong>�ͼӳ��: </strong>
                                        <select name="pay_month" id="pay_month" type="text" value="<%=pay_month%>" style="width:90px">
                                            <%	for i = 24 to 1 step -1	%>
                                            <option value="<%=month_tab(i,1)%>" <%If pay_month = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                            <%	next	%>
                                        </select>
                                    </label>
                                    <br>
                                    <label>
                                        <strong>���ε�����: </strong>
                                        <input name="att_file" type="file" id="att_file" size="100" value="<%=att_file%>" style="text-align:left"> 
                                    </label>

                                    <input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
                                    <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                                </p>
                            </dd>
						</dl>
					</fieldset>
					<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
							<colgroup>
								<col width="3%" >
								<col width="3%" >
								<col width="4%" >
								<col width="4%" >
								<col width="7%" >
								<col width="7%" >
								<col width="4%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="6%" >
								<col width="5%" >
								<col width="5%" >
								<col width="6%" >
                                <col width="6%" >
                                <col width="5%" >
                                <col width="*" >
                                <col width="6%" >
                                <col width="8%" >
							</colgroup>
							<thead>
								<tr>
									<th class="first" scope="col">�Ǽ�</th>
									<th scope="col">���</th>
									<th scope="col">���</th>
									<th scope="col">����</th>
									<th scope="col">�⺻��</th>
									<th scope="col">�Ĵ�</th>
									<th scope="col">����<br>����</th>
									<th scope="col">��ź�</th>
									<th scope="col">�ұ�</th>
									<th scope="col">����</th>
									<th scope="col">����<br>����</th>
									<th scope="col">��å</th>
									<th scope="col">��<br>����</th>
									<th scope="col">����<br>����</th>
                                    <th scope="col">����<br>���</th>
                                    <th scope="col">������<br>�ٹ�</th>
                                    <th scope="col">�ټ�</th>
                                    <th scope="col">�����</th>
                                    <th scope="col">����<br>�װ�</th>
								</tr>
							</thead>
							<tbody>
								<%
                                tot_emp          = 0
                                tot_name         = 0
                                tot_bank         = 0
                                tot_err          = 0

                                tot_base_pay     = 0
                                tot_meals_pay    = 0
                                tot_research_pay = 0
                                tot_postage_pay  = 0
                                tot_re_pay       = 0
                                tot_overtime_pay = 0
                                tot_car_pay      = 0
                                tot_position_pay = 0
                                tot_custom_pay   = 0
                                tot_job_pay      = 0
                                tot_job_support  = 0
                                tot_jisa_pay     = 0
                                tot_long_pay     = 0
                                tot_disabled_pay = 0
                                tot_family_pay   = 0
                                tot_school_pay   = 0
                                tot_qual_pay     = 0
                                tot_other_pay1   = 0
                                tot_other_pay2   = 0
                                tot_other_pay3   = 0
                                tot_tax_yes      = 0
                                tot_tax_no       = 0
                                tot_tax_reduced  = 0	
                                tot_give_total   = 0
                                        
                                if rowcount > -1 then
                                    for i=0 to rowcount
                                    if xgr(1,i) = "" or isnull(xgr(1,i)) then
                                        exit for
                                    end if
                                    
                                    ' ���üũ 				
                                    emp_sw = "Y"
                                    emp_no = xgr(3,i)
                                    Sql = "select * from emp_master where emp_no = '"&xgr(3,i)&"'"
                                    Set rs_emp = DbConn.Execute(Sql)
                                    'Response.write Sql & "<br>"
                                    if rs_emp.eof then
                                        tot_emp = tot_emp + 1
                                        tot_err = tot_err + 1
                                        emp_sw = "N"
                                        emp_name =""
                                    else
                                        emp_name = rs_emp("emp_name")	  
                                    end if
                                    name_sw = "Y"
                                    
        '							if xgr(4,i) <> emp_name then
        '							    tot_name = tot_name + 1
        '								tot_err = tot_err + 1
        '								name_sw = "N"
        '								emp_name = xgr(4,i)	
        '							end if

                                    ' �������üũ
                                    bank_sw = "Y"
                                    Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
                                    Set rs_bnk = DbConn.Execute(SQL)
                                    if  rs_bnk.eof then
                                        tot_bank = tot_bank + 1
                                        tot_err = tot_err + 1
                                        bank_sw = "N"
                                    end if
                                    rs_bnk.close()	 

                                    ' �����׸�
                                    pmg_base_pay	  = toString(xgr(12,i),"0")	'�⺻��
                                    pmg_meals_pay	  = toString(xgr(13,i),"0")	'�Ĵ�
                                    pmg_research_pay  = toString(xgr(14,i),"0")	'�������� (�ű��߰�)
                                    pmg_postage_pay	  = toString(xgr(15,i),"0")	'��ź�
                                    pmg_re_pay		  = toString(xgr(16,i),"0")	'�ұޱ޿�
                                    pmg_overtime_pay  = toString(xgr(17,i),"0")	'����ٷμ���
                                    pmg_car_pay		  = toString(xgr(18,i),"0")	'����������
                                    pmg_position_pay  = toString(xgr(19,i),"0")	'��å����
                                    pmg_custom_pay	  = toString(xgr(20,i),"0")	'����������
                                    pmg_job_pay		  = toString(xgr(21,i),"0")	'����������
                                    pmg_job_support	  = toString(xgr(22,i),"0")	'���������
                                    pmg_jisa_pay	  = toString(xgr(23,i),"0")	'������ٹ���
                                    pmg_long_pay	  = toString(xgr(24,i),"0")	'�ټӼ���
                                    pmg_disabled_pay  = toString(xgr(25,i),"0")	'����μ���
                                    pmg_family_pay 	  = 0
                                    pmg_school_pay 	  = 0
                                    pmg_qual_pay 	  = 0
                                    pmg_other_pay1 	  = 0
                                    pmg_other_pay2 	  = 0
                                    pmg_other_pay3 	  = 0
                                    pmg_tax_yes 	  = 0
                                    pmg_tax_no 		  = 0
                                    pmg_tax_reduced   = 0	
                                    
                                    pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_research_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
                                    'pmg_give_total = xgr(25,i)	
                                    
                                    ' �����׸�			
                                    de_nps_amt			= toString(xgr(27,i),"0")
                                    de_nhis_amt			= toString(xgr(28,i),"0")
                                    de_epi_amt			= toString(xgr(29,i),"0")
                                    de_longcare_amt		= toString(xgr(30,i),"0")
                                    de_income_tax		= toString(xgr(31,i),"0")
                                    de_wetax			= toString(xgr(32,i),"0")
                                    de_year_incom_tax	= toString(xgr(33,i),"0")
                                    de_year_wetax		= toString(xgr(34,i),"0")
                                    de_year_incom_tax2	= toString(xgr(35,i),"0")
                                    de_year_wetax2		= toString(xgr(36,i),"0")
                                    de_other_amt1		= toString(xgr(37,i),"0")
                                    de_special_tax		= 0
                                    de_saving_amt		= 0
                                    de_sawo_amt			= toString(xgr(38,i),"0")
                                    de_johab_amt		= 0
                                    de_school_amt		= toString(xgr(39,i),"0")
                                    de_nhis_bla_amt		= toString(xgr(40,i),"0")
                                    de_long_bla_amt		= toString(xgr(41,i),"0")
                                    de_hyubjo_amt		= toString(xgr(42,i),"0")
                                    
                                    de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2 + de_other_amt1 + de_sawo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt + de_hyubjo_amt
                                    'de_deduct_total = xgr(38,i)

                                    sql = "select * from pay_month_give where pmg_yymm = '"&pay_month&"' and pmg_id = '1' and pmg_emp_no = '"&emp_no&"'"
                                    set Rs_give=dbconn.execute(sql)
                                    'Response.write sql&"<br>"
                                    if Rs_give.eof or Rs_give.bof then
                                        reg_sw = "N"
                                    else
                                        reg_sw = "Y"
                                    end if
                                    
                                    tot_base_pay 		= tot_base_pay     + pmg_base_pay
                                    tot_meals_pay 		= tot_meals_pay    + pmg_meals_pay
                                    tot_research_pay 	= tot_research_pay + pmg_research_pay
                                    tot_postage_pay 	= tot_postage_pay  + pmg_postage_pay
                                    tot_re_pay 			= tot_re_pay       + pmg_re_pay
                                    tot_overtime_pay 	= tot_overtime_pay + pmg_overtime_pay
                                    tot_car_pay 		= tot_car_pay      + pmg_car_pay
                                    tot_position_pay 	= tot_position_pay + pmg_position_pay
                                    tot_custom_pay 		= tot_custom_pay   + pmg_custom_pay
                                    tot_job_pay 		= tot_job_pay      + pmg_job_pay
                                    tot_job_support 	= tot_job_support  + pmg_job_support
                                    tot_jisa_pay 		= tot_jisa_pay     + pmg_jisa_pay
                                    tot_long_pay 		= tot_long_pay     + pmg_long_pay
                                    tot_disabled_pay 	= tot_disabled_pay + pmg_disabled_pay
                                    tot_family_pay 		= tot_family_pay   + pmg_family_pay
                                    tot_school_pay 		= tot_school_pay   + pmg_school_pay
                                    tot_qual_pay 		= tot_qual_pay     + pmg_qual_pay
                                    tot_other_pay1 		= tot_other_pay1   + pmg_other_pay1
                                    tot_other_pay2 		= tot_other_pay2   + pmg_other_pay2
                                    tot_other_pay3 		= tot_other_pay3   + pmg_other_pay3
                                    tot_tax_yes 		= tot_tax_yes      + pmg_tax_yes
                                    tot_tax_no 			= tot_tax_no       + pmg_tax_no
                                    tot_tax_reduced 	= tot_tax_reduced  + pmg_tax_reduced
                                    tot_give_total 		= tot_give_total   + pmg_give_total
                                    
                                    
                                    if reg_sw = "N" then 
                                        reg_flag = "No"
                                        bgcolor0=""
                                    else
                                        reg_flag = "Yes"
                                        bgcolor0="#FFCCFF"
                                    end if
                                    
                                    if emp_sw = "Y" then
                                        bgcolor1=""
                                    else
                                        bgcolor1="#FFCCFF"
                                    end if
                                    
                                    if name_sw = "Y" then
                                        bgcolor2=""
                                    else
                                        bgcolor2="#FFCCFF"
                                    end if
                                    
                                    if bank_sw = "Y" then
                                        bgcolor3=""
                                    else
                                        bgcolor3="#FFCCFF"
                                    end if
                                    %>
                                    <tr>
                                        <td class="first"><%=i+1%></td>
                                        <td bgcolor="<%=bgcolor%>"><%=reg_flag%></td>                            
                                        <td bgcolor="<%=bgcolor%>"><%=emp_no%></td>
                                        <td bgcolor="<%=bgcolor%>"><%=emp_name%></td>
                                        <td class="right" bgcolor="<%=bgcolor%>"><%=formatnumber(pmg_base_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_meals_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_research_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_postage_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_re_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_overtime_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_car_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_position_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_custom_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_job_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_job_support,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_jisa_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_long_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_disabled_pay,0)%></td>
                                        <td class="right"><%=formatnumber(pmg_give_total,0)%></td>
                                    </tr>
                                    <%
                                    next
                                end if
								%>
								<tr>
									<th class="first">����</th>
									<th title="�޿����¹̵�ϰǼ�"><%=formatnumber(tot_bank,0)%></th>
									<th title="�����̵�ϰǼ�"><%=formatnumber(tot_emp,0)%></th>
									<th><%=formatnumber(tot_name,0)%></th>
									<th class="right"><%=formatnumber(tot_base_pay,0)%></th>
									<th class="right"><%=formatnumber(tot_meals_pay,0)%></th>
									<th class="right"><%=formatnumber(tot_research_pay,0)%></th>								
									<th class="right"><%=formatnumber(tot_postage_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_re_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_overtime_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_car_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_position_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_custom_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_job_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_job_support,0)%></th>
                                    <th class="right"><%=formatnumber(tot_jisa_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_long_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_disabled_pay,0)%></th>
                                    <th class="right"><%=formatnumber(tot_give_total,0)%></th>
								</tr>
							</tbody>
						</table>
					</div>
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  	<tr>
                        <td width="15%"><div class="btnCenter"></div></td>
                        <td>
                            <div class="btnRight"><a href="#" onClick="pay_month_updel('pay_month','pay_company');return false;" class="btnType04">�޿� Upload ����</a></div>
                        </td> 
                    </tr>
                    </table>
                    <input type="hidden" name="pay_company1" value="<%=pay_company%>" ID="Hidden1">
                    <input type="hidden" name="pay_month1" value="<%=pay_month%>" ID="Hidden1">             

				</form>
				
				<%
                if emp_payend = "N" then 
                    if tot_cnt <> 0 and tot_err = 0 then 
                    %>
                        <form action="insa_pay_month_up_ok.asp" method="post" name="frm1">
                            <br>
                            <div align=center>
                                <span class="btnType01"><input type="button" value="�޿��ڷ� Upload" onclick="javascript:frm1check();"NAME="Button1"></span>
                            </div>
                            <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                            <input name="pmg_yymm" type="hidden" id="pmg_yymm" value="<%=pay_month%>">
                            <input name="pmg_date" type="hidden" id="pmg_date" value="<%=give_date%>">
                            <input name="pmg_company" type="hidden" id="pmg_company" value="<%=pay_company%>">
                            <br>
                        </form>
				    <%
					end if
			   	end if 
			  %>
			</div>				
		</div>
	</body>
</html>

