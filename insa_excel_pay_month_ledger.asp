<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '������

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " �޿�����(" + view_condi + ")"

savefilename = title_line +".xls"
'savefilename = "�Ի��� ��Ȳ -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

	sum_base_pay = 0
	sum_meals_pay = 0
	sum_postage_pay = 0
	sum_re_pay = 0
	sum_overtime_pay = 0
	sum_car_pay = 0
	sum_position_pay = 0
	sum_custom_pay = 0
	sum_job_pay = 0
	sum_job_support = 0
	sum_jisa_pay = 0
	sum_long_pay = 0
	sum_disabled_pay = 0
	sum_family_pay = 0
	sum_school_pay = 0
	sum_qual_pay = 0
	sum_other_pay1 = 0
	sum_other_pay2 = 0
	sum_other_pay3 = 0
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
    sum_nps_amt = 0
    sum_nhis_amt = 0
    sum_epi_amt = 0
    sum_longcare_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_year_incom_tax2 = 0
    sum_year_wetax2 = 0
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

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

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="2" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��������</div></td>
    <td colspan="7" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">�⺻�޿� �� ������</div></td>
    <td colspan="6" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">���� �� �������޾�</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">��  ��</div></td>
    <td><div align="center" class="style1">�⺻��</div></td>
    <td><div align="center" class="style1">�Ĵ�</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">��ź�</div></td>
    <td><div align="center" class="style1">�ұޱ޿�</div></td>
    <td><div align="center" class="style1">����ٷμ���</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">���ο���</div></td>
    <td><div align="center" class="style1">�ǰ�����</div></td>
    <td><div align="center" class="style1">��뺸��</div></td>
    <td><div align="center" class="style1">����纸���</div></td>
    <td><div align="center" class="style1">�ҵ漼</div></td>
    <td><div align="center" class="style1">����ҵ漼</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">�Ի���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><div align="center" class="style1">����</div></td>
    <td><div align="center" class="style1">��å����</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">����������</div></td>
    <td><div align="center" class="style1">���������</div></td>
    <td><div align="center" class="style1">������ٹ���</div></td>
    <td><div align="center" class="style1">�ټӼ���</div></td>
    <td><div align="center" class="style1">����μ���</div></td>
    <td><div align="center" class="style1">��Ÿ����</div></td>
    <td><div align="center" class="style1">���ȸ ȸ��</div></td>
    <td><div align="center" class="style1">���ڱݻ�ȯ</div></td>
    <td><div align="center" class="style1">�ǰ����������</div></td>
    <td><div align="center" class="style1">����纸�������</div></td>
    <td><div align="center" class="style1">�����հ�</div></td>  
  </tr>
  <tr>
    <td style=" border-bottom:2px solid #515254; background:#f8f8f8;"><div align="center" class="style1">�����</div></td>
    <td style=" border-bottom:2px solid #515254; background:#f8f8f8;"><div align="center" class="style1">�μ�</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">&nbsp;</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">�����հ�</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">������</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">��������ҵ漼</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">�����������漼</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">����������ҵ漼</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">�������������漼</div></td>
    <td style=" border-bottom:2px solid #515254;"><div align="center" class="style1">�������޾�</div></td>      
 </tr>

    <%
		do until rs.eof 
		
		  emp_no = rs("pmg_emp_no")
		  pmg_give_tot = rs("pmg_give_total")
		  pay_count = pay_count + 1
					  
		  sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
	      sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
	      sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
	      sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
	      sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
	      sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
          sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
	      sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
	      sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
	      sum_job_support = sum_job_support + int(rs("pmg_job_support"))
	      sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
	      sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
	      sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
	      sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_name")%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_base_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_meals_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_postage_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_re_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_overtime_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_car_pay"),0)%></div></td>
    
    <%
          Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
          Set rs_emp = DbConn.Execute(SQL)
		  if not rs_emp.eof then
				emp_in_date = rs_emp("emp_in_date")
				emp_end_date = rs_emp("emp_end_date")
	         else
				emp_end_date = ""
				emp_in_date = ""
          end if
          rs_emp.close()
		  if emp_end_date = "1999-01-01" then emp_end_date = "" end if
    %>    
    
    <%
	      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
          Set Rs_dct = DbConn.Execute(SQL)
		  if not Rs_dct.eof then
				de_nps_amt = int(Rs_dct("de_nps_amt"))
                de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                de_epi_amt = int(Rs_dct("de_epi_amt"))
                de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                de_income_tax = int(Rs_dct("de_income_tax"))
                de_wetax = int(Rs_dct("de_wetax"))
				de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                de_year_wetax = int(Rs_dct("de_year_wetax"))
				de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
                de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
                de_other_amt1 = int(Rs_dct("de_other_amt1"))
                de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                de_school_amt = int(Rs_dct("de_school_amt"))
                de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
                de_deduct_tot = int(Rs_dct("de_deduct_total"))	
             else
				de_nps_amt = 0
                de_nhis_amt = 0
                de_epi_amt = 0
                de_longcare_amt = 0
                de_income_tax = 0
                de_wetax = 0
				de_year_incom_tax = 0
                de_year_wetax = 0
				de_year_incom_tax2 = 0
                de_year_wetax2 = 0
                de_other_amt1 = 0
                de_sawo_amt = 0
                de_hyubjo_amt = 0
                de_school_amt = 0
                de_nhis_bla_amt = 0
                de_long_bla_amt = 0
                de_deduct_tot = 0
           end if
           Rs_dct.close()
		   
		   pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
	 	   sum_nps_amt = sum_nps_amt + de_nps_amt
           sum_nhis_amt = sum_nhis_amt + de_nhis_amt
           sum_epi_amt = sum_epi_amt + de_epi_amt
		   sum_longcare_amt = sum_longcare_amt + de_longcare_amt
           sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           sum_year_wetax = sum_year_wetax + de_year_wetax
		   sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
           sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
           sum_other_amt1 = sum_other_amt1 + de_other_amt1
           sum_sawo_amt = sum_sawo_amt + de_sawo_amt
           sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
           sum_school_amt = sum_school_amt + de_school_amt
           sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
           sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
		   sum_deduct_tot = sum_deduct_tot + de_deduct_tot
							  
    %>    
    
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
  </tr>
  <tr>    
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_grade")%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_position_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_custom_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_support"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_jisa_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_long_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_disabled_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_other_amt1,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_school_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_long_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
  </tr>    
  <tr>    
    <td width="110" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=emp_end_date%></div></td>
    <td width="110" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1"><%=rs("pmg_org_name")%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(rs("pmg_give_total"),0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_hyubjo_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax2,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(de_year_wetax2,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
  </tr>        
	<%
	    Rs.MoveNext()
	loop
	
	sum_curr_pay = sum_give_tot - sum_deduct_tot
	
	%>
    
  <tr>    
    <th rowspan="3" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="center" class="style1">�Ѱ�</div></th>
    <th rowspan="3" width="110" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(pay_count,0)%>&nbsp;��</div></th>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_re_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sumpmg_car_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nps_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nhis_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_longcare_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
  </tr>
  <tr>    
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_position_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_job_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_job_support,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_long_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_other_amt1,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_sawo_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_school_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_nhis_bla_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_long_bla_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
  </tr>
  <tr>    
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1">&nbsp;</div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_give_tot,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_hyubjo_amt,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax2,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax2,0)%></div></td>
    <td width="100" style=" border-bottom:1px solid #e3e3e3; background:#fff0f5;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
