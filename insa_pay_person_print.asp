<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(10,10)

curr_date = mid(cstr(now()),1,10)

pmg_emp_no = request("emp_no")
pmg_emp_name = request("emp_name")
pmg_yymm = request("pmg_yymm")
pmg_date = request("pmg_date")
pmg_company = request("pmg_company")
pmg_org_code = request("pmg_org_code")
pmg_org_name = request("pmg_org_name")
pmg_grade = request("pmg_grade")
pmg_position = request("pmg_position")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")

DbConn.Open dbconnect

	Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+pmg_emp_no+"') and (pmg_company = '"+pmg_company+"')"
	set rs = dbconn.execute(sql)

    pmg_yymm = rs("pmg_yymm")
	pmg_emp_no = rs("pmg_emp_no")
    pmg_company = rs("pmg_company")
	pmg_date = rs("pmg_date")
	pmg_emp_name = rs("pmg_emp_name")
	pmg_org_code = rs("pmg_org_code")
	pmg_org_name = rs("pmg_org_name")
	pmg_grade = rs("pmg_grade")
	pmg_position = rs("pmg_position")

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
	pmg_family_pay = rs("pmg_family_pay")
	pmg_school_pay = rs("pmg_school_pay")
	pmg_qual_pay = rs("pmg_qual_pay")
	pmg_other_pay1 = rs("pmg_other_pay1")
	pmg_other_pay2 = rs("pmg_other_pay2")
	pmg_other_pay3 = rs("pmg_other_pay3")
	pmg_tax_yes = rs("pmg_tax_yes")
	pmg_tax_no = rs("pmg_tax_no")
	pmg_tax_reduced = rs("pmg_tax_reduced")
	pmg_give_tot = rs("pmg_give_total")

	rs.close()

	Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+pmg_emp_no+"') and (de_company = '"+pmg_company+"')"
    Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then
           de_nps_amt = Rs_dct("de_nps_amt")
           de_nhis_amt = Rs_dct("de_nhis_amt")
           de_epi_amt = Rs_dct("de_epi_amt")
		   de_longcare_amt = Rs_dct("de_longcare_amt")
           de_income_tax = Rs_dct("de_income_tax")
           de_wetax = Rs_dct("de_wetax")
		   de_year_incom_tax = Rs_dct("de_year_incom_tax")
           de_year_wetax = Rs_dct("de_year_wetax")
		   de_year_incom_tax2 = Rs_dct("de_year_incom_tax2")
           de_year_wetax2 = Rs_dct("de_year_wetax2")
           de_other_amt1 = Rs_dct("de_other_amt1")
           de_sawo_amt = Rs_dct("de_sawo_amt")
           de_hyubjo_amt = Rs_dct("de_hyubjo_amt")
           de_school_amt = Rs_dct("de_school_amt")
           de_nhis_bla_amt = Rs_dct("de_nhis_bla_amt")
           de_long_bla_amt = Rs_dct("de_long_bla_amt")
		   de_deduct_tot = Rs_dct("de_deduct_total")
	   else
		   de_deduct_tot = 0
    end if
    Rs_dct.close()

pay_curr_amt = pmg_give_tot - de_deduct_tot

	sql = "select * from emp_master where emp_no='" +pmg_emp_no+ "'"
		Set Rs_emp = DbConn.Execute(SQL)
		if not Rs_emp.eof then
			emp_in_date = Rs_emp("emp_in_date")
		  else
			emp_in_date = ""
		end if
		Rs_emp.close()

    Sql = "SELECT * FROM pay_bank_account where emp_no = '"+pmg_emp_no+"'"
    Set rs_bnk = DbConn.Execute(SQL)
    if not rs_bnk.eof then
           bank_name = rs_bnk("bank_name")
           account_no = rs_bnk("account_no")
		   account_holder = rs_bnk("account_holder")
	   else
           bank_name = ""
		   account_no = ""
		   account_holder = ""
    end if
    rs_bnk.close()

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)

main_title = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " �޿�����"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>�λ�޿� �ý���</title>
        <script src="/java/common.js" type="text/javascript"></script>
        <script type="text/javascript">
	function printWindow(){
//		viewOff("button");
		factory.printing.header = ""; //�Ӹ��� ����
		factory.printing.footer = ""; //������ ����
		factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
		factory.printing.leftMargin = 13; //���� ���� ����
		factory.printing.topMargin = 25; //���� ���� ����
		factory.printing.rightMargin = 13; //�����P ���� ����
		factory.printing.bottomMargin = 15; //�ٴ� ���� ����
//		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
//		factory.printing.printer = ""; //������ �� ������ �̸�
//		factory.printing.paperSize = "A4"; //��������
//		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
//		factory.printing.collate = true; //������� ����ϱ�
//		factory.printing.copies = "1"; //�μ��� �ż�
//		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
//		factory.printing.Printer(true); //����ϱ�
		factory.printing.Preview(); //�����츦 ���ؼ� ���
		factory.printing.Print(false); //�����츦 ���ؼ� ���
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}

</script>
<title>���� �޿�����</title>
<style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14C {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14R {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style14L {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
</style>
<style media="print">
.noprint     { display: none }
</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div class="noprint">
<p><a href="#" onClick="printWindow()"><img src="image/printer.jpg" width="39" height="36" border="0" alt="����ϱ�" /></a></p>
</div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>

<table width="690" cellpadding="0" cellspacing="0">
  <tr>
     <td colspan="3" align="center" class="style32BC"><%=main_title%></td>
  </tr>
  <tr>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
  </tr>
</table>
<table width="690" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�����ȣ</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_no%></span></td>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">��� ��</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_emp_name%></span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�� ��</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_grade%></span></td>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�Ի�����</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=emp_in_date%></span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">�Ҽ�</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=pmg_org_name%>(<%=pmg_org_code%>)</span></td>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14BC">���¹�ȣ</span></td>
    <td width="30%" height="30" align="left"><span class="style14C">&nbsp;&nbsp;<%=account_no%><br>&nbsp;&nbsp;(<%=bank_name%>-<%=account_holder%>)</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">���޳���</span></td>
    <td width="30%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14BC">���޾�</span></td>
    <td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">��������</span></td>
    <td width="30%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14BC">������</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">�⺻��</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_base_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">���ο���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_nps_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">�Ĵ�</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_meals_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">�ǰ�����</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_nhis_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center"><span class="style14C">��ź�</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_postage_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">��뺸��</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_epi_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">�ұޱ޿�</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_re_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">����纸��</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_longcare_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <tr>
    <td width="20%" height="30" align="center"><span class="style14C">����ٷμ���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_overtime_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">�ҵ漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_income_tax,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_car_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">����ҵ漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_wetax,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">��å����</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_position_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">��Ÿ����</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_other_amt1,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_custom_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">���ȸ ȸ��</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_sawo_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">����������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_job_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_hyubjo_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">���������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_job_support,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">���ڱݴ���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_school_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">������ٹ���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_jisa_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">�ǰ����������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_nhis_bla_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
    <tr>
    <td width="20%" height="30" align="center"><span class="style14C">�ټӼ���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_long_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">����纸�������</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_long_bla_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">����μ���</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(pmg_disabled_pay,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">��������ҵ漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_year_incom_tax,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
    <td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">�����������漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_year_wetax,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
    <td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">����������ҵ漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_year_incom_tax2,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
    <td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center"><span class="style14C">�������������漼</span></td>
    <td width="30%" height="30" align="right"><span class="style14C"><%=formatnumber(de_year_wetax2,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center"><span class="style14C">&nbsp;</span></td>
    <td width="30%" height="30" align="right"><span class="style14C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center" bgcolor="#E0FFFF"><span class="style14C">�����װ�</span></td>
    <td width="30%" height="30" align="right" bgcolor="#E0FFFF"><span class="style14C"><%=formatnumber(de_deduct_tot,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="20%" height="30" align="center" bgcolor="#FFFFE6"><span class="style14C">���޾װ�</span></td>
    <td width="30%" height="30" align="right" bgcolor="#FFFFE6"><span class="style14C"><%=formatnumber(pmg_give_tot,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
    <td width="20%" height="30" align="center" bgcolor="#BFBFFF"><span class="style14C">�������޾�</span></td>
    <td width="30%" height="30" align="right" bgcolor="#BFBFFF"><span class="style14C"><%=formatnumber(pay_curr_amt,0)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
</table>
<table width="690" cellpadding="0" cellspacing="0">
  <tr>
     <td width="50%" height="30" align="left" class="style1">�� ������ ��� ����帳�ϴ�</td>
  <% if pmg_company = "���̿��������" then %>
	 <td width="50%" height="30" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=60 height=60 alt="" align=right><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̿��������<br /></td>
  <% end if %>
  <% if pmg_company = "�޵�" then %>
	 <td width="50%" height="30" align="right" valign="middle" width="100%"><img src="image/k-hudis001.png" width=60 height=60 alt="" align=right><font style="font-size:14px"><br><br>�ֽ�ȸ�� �޵�<br /></td>
  <% end if %>
  <% if pmg_company = "���̳�Ʈ����" then %>
	 <td width="50%" height="30" align="right" valign="middle" width="100%"><img src="image/k-net001.png" width=60 height=60 alt="" align=right><font style="font-size:14px"><br><br>���̳�Ʈ���� �ֽ�ȸ��<br /></td>
  <% end if %>
  <% if pmg_company = "����������ġ" then %>
	 <td width="50%" height="30" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=60 height=60 alt="" align=right><font style="font-size:14px"><br><br>�ֽ�ȸ�� ����������ġ<br /></td>
  <% end if %>
  <% if pmg_company = "�ڸ��Ƶ𿣾�" then %>
	 <td width="50%" height="30" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=60 height=60 alt="" align=right><font style="font-size:14px"><br><br>�ֽ�ȸ�� ����������ġ<br /></td>
  <% end if %>
  </tr>
</table>
</p>

</body>
</html>
