<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_agree = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

sql = "select * from emp_master where emp_no = '" + emp_no  + "'"
Rs.Open Sql, Dbconn, 1

agree_empno = rs("emp_no")
agree_emp_type = rs("emp_type")
agree_empname = rs("emp_name")
agree_company = rs("emp_company")
agree_org_name = rs("emp_org_name")
agree_grade = rs("emp_grade")
agree_job = rs("emp_job")
agree_position = rs("emp_position")
agree_jikmu = rs("emp_jikmu")
agree_person1 = rs("emp_person1")
agree_person2 = rs("emp_person2")
agree_sido = rs("emp_sido")
agree_gugun = rs("emp_gugun")
agree_dong = rs("emp_dong")
agree_addr = rs("emp_addr")
agree_tel_ddd = rs("emp_tel_ddd")
agree_tel_no1 = rs("emp_tel_no1")
agree_tel_no2 = rs("emp_tel_no2")

emp_in_date = mid(cstr(rs("emp_in_date")),1,10)
emp_in_year = mid(cstr(rs("emp_in_date")),1,4)
emp_in_month = mid(cstr(rs("emp_in_date")),6,2)
emp_in_day = mid(cstr(rs("emp_in_date")),9,2)

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)

year_cnt = datediff("yyyy", rs("emp_in_date"), curr_date)
mon_cnt = datediff("m", rs("emp_in_date"), curr_date)
day_cnt = datediff("d", rs("emp_in_date"), curr_date)
'rs.close()
'response.write(year_cnt)
'response.write(mon_cnt)
'response.write(day_cnt)
emp_no = "100173"

Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&curr_year&"'"
Set rs_year = DbConn.Execute(SQL)
if not rs_year.eof then
       incom_base_pay = rs_year("incom_base_pay")
       incom_overtime_pay = rs_year("incom_overtime_pay")
	   incom_meals_pay = rs_year("incom_meals_pay")
       incom_severance_pay = rs_year("incom_severance_pay")
	   incom_total_pay = rs_year("incom_total_pay")
	   incom_first3_percent = rs_year("incom_first3_percent")
   else
       incom_base_pay = 0
       incom_overtime_pay = 0
       incom_meals_pay = 0
	   incom_severance_pay = 0
       incom_total_pay = 0
	   incom_first3_percent = 0
end if
rs_year.close()

' �ݾ��� �ѱ۷� ��ȯ....
'amt = "21345000"
amt = incom_total_pay
Dim unit1(10)
Dim unit2(2)
Dim unit3(2)

unit1(0) = ""
unit1(1) = "��"
unit1(2) = "��"
unit1(3) = "��"
unit1(4) = "��"
unit1(5) = "��"
unit1(6) = "��"
unit1(7) = "ĥ"
unit1(8) = "��"
unit1(9) = "��"

unit2(0) = "��"
unit2(1) = "��"
unit2(2) = "õ"

unit3(0) = "��"
unit3(1) = "��"
unit3(2) = "��"
 
vamt = Replace(amt, ",", "")
xchk = IsNumeric(vamt)

If xchk = True Then
    total = Len(CStr(CDbl(amt)))
    vamt = CDbl(amt)
    rt_amt = ""
    For i = 1 To total
        num = Mid(vamt, i, 1)
        temp1 = (total - i) + 1
        rt_amt = rt_amt & unit1(num)
 
        If num <> 0 And i <> total Then
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 0 Then rt_amt = rt_amt & unit2(2)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 3 Then rt_amt = rt_amt & unit2(1)
            If Len(Left(vamt, (total - i) + 1)) Mod 4 = 2 Then rt_amt = rt_amt & unit2(0)
        End If
 
        If temp1 = 5 And Right(rt_amt, 1) <> unit3(2) And Right(rt_amt, 1) <> unit3(1) Then rt_amt = rt_amt & unit3(0)
        If temp1 = 9 And Right(rt_amt, 1) <> unit3(2) Then rt_amt = rt_amt & unit3(1)
        If temp1 = 13 Then rt_amt = rt_amt & unit3(2)
 
    Next
 
    rt_amt = rt_amt & "��"
 
    'msgbox
    'response.write  "input : " & amt & vbCr & "output : " & rt_amt
End If


seq_last = ""
agree_year = curr_year
agree_id = "�����ٷΰ�༭"       

    sql="select max(agree_seq) as max_seq from emp_agree where agree_empno = '"&emp_no&"' and agree_year = '"&agree_year&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		seq_last = "001"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		seq_last = right(max_seq,3)
	end if
    rs_max.close()

agree_seq = seq_last

main_title = cstr(agree_year) + "�� "  + " ���� �ٷΰ�༭"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>���ξ���-�λ�</title>
        <script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" id='dummy'></script>
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
	
    function year_income_agree(val, val2, val3) {
            if (!confirm("�����ٷΰ���� �����Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm; 
			
			alert (val);
			alert (val2);
			alert (val3);
            
			document.frm.action = "insa_year_income_agree_save.asp?emp_no=" + val;
            document.frm.submit();
			
			<%
			'var scpt= document.getElementById('dummy');
			'alert (scpt);
			'scpt.src='insa_year_income_agree_save.asp?emp_no='+val;
			'document.submit();
			%>
    }	
	
</script>
<title>���� �����ٷΰ�༭</title>
<style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
		.style3 {font-size:14px;color: #666666}
-->
</style>
<style media="print"> 
.noprint     { display: none }
</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <form action="insa_year_income_vow.asp" method="post" name="frm">
    <div align=center class="noprint">
     <p>
        <% '<a href="javascript:printW();"><img src="image/b_print.gif" border="0" alt="����ϱ�" /></a> %>
        <a href="#" onClick="year_income_agree('<%=agree_empno%>','<%=agree_empname%>','<%=curr_year%>');return false;" style="border-width:0px;"><img src="image/b_agree2.jpg" border="0" alt="�����ϱ�" /></a>
        <a href="#" onClick="printWindow()"><img src="image/b_print.gif" border="0" alt="����ϱ�" /></a>
        <a href="javascript:goBefore();"><img src="image/b_close.gif" border="0" alt="�ݱ�" /></a> 
     </p>
    </div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>    
   
<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td colspan="3" align="center" class="style32BC"><%=main_title%></td>
  </tr>
  <tr>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
  </tr>
</table>
<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="10%" height="30" rowspan="3" align="center" bgcolor="#eaeaea"><span class="style14BC">�����</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">���ü��</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;(��)���̿��������</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">��ǥ��</span></td>
    <td width="25%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;�� ����</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">������</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">&nbsp;&nbsp;����� ��õ�� ���������2�� 18 �븢��ũ��Ÿ��1�� 6��</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea" style=" border-bottom:2px solid #515254;"><span class="style14BC">����</span></td>
    <td colspan="3" height="30" align="center" style=" border-bottom:2px solid #515254;"><span class="style14C">&nbsp;&nbsp;��ǻ�� ���� � �� ��Ű���</span></td>
  </tr>
  <tr>
    <td width="10%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">�ٷ���</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_empname%>&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ֹε�Ϲ�ȣ</span></td>
    <td width="25%" height="30" align="center"><span class="style14C"><%=agree_person1%>-*******&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ּ�</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C"><%=agree_sido%>&nbsp;<%=agree_gugun%>&nbsp;<%=agree_dong%>&nbsp;<%=agree_addr%></span></td>
  </tr>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     <br/>&nbsp;&nbsp;����� ����ڿ� �ٷ��ڴ� ���� ������ �������� �����ǻ�� ������ ���� �ٷΰ���� ü���ϰ� ������<br/> ������ �����ϱ� ���Ͽ� �̸� ������ ���� �� ���� �����Ѵ�.<br/><br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="center" class="style3">
     <br/> -&nbsp;&nbsp; ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp; - <br/><br/></td>
  </tr>
</table>

<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ٷ����</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">�����μ� ������</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ٹ�����</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">������</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_no%>&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_grade%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�Ի���</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=emp_in_year%>��&nbsp;<%=emp_in_month%>��&nbsp;<%=emp_in_day%>��</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����ó</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=agree_tel_ddd%>-<%=agree_tel_no1%>-<%=agree_tel_no2%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ٷαⰣ</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">�ٷαⰣ�� ���� ������ ����</span></td>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
<% if emp_in_year = curr_year then %>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>1. �ӱ����� ���Ⱓ :&nbsp;</strong><%=emp_in_year%>��&nbsp;<%=emp_in_month%>��&nbsp;<%=emp_in_day%>�� ~ </strong><%=mid(cstr(now()),1,4)%>��&nbsp;12��&nbsp;31��</td>
<%    else %>    
     <td width="100%" height="30" align="left" class="style3"><br/><strong>1. �ӱ����� ���Ⱓ :&nbsp;</strong><%=mid(cstr(now()),1,4)%>��&nbsp;01��&nbsp;01�� ~ </strong><%=mid(cstr(now()),1,4)%>��&nbsp;12��&nbsp;31��</td>
<%  end if %>    
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;��� �Ⱓ ���̶� �ʿ��ϴٰ� ������ ��� ��ȣ �����Ͽ� �� �޿����� ������ �� �ִ�.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>2. �ӱݳ���</strong></td>
  </tr>
</table>

<table width="690" border="1px" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�Ѽ��ɾ�</span></td>
    <td colspan="3" height="30" align="center"><span class="style14C">�ϱ�&nbsp;:&nbsp;<%=rt_amt%>&nbsp;&nbsp;&nbsp;(\:<%=formatnumber(incom_total_pay,0)%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">����</span></td>
    <td width="35%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ش�ݾ�</span></td>
    <td width="15%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">���ޱ���</span></td>
    <td width="35%" height="30" rowspan="2" align="center" bgcolor="#eaeaea"><span class="style14BC">���</span></td>
  </tr>
  <tr>
    <td width="35%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�� ������</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�⺻��</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_base_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">-</span></td>
    <td width="35%" height="30" align="center"><span class="style14C">&nbsp;&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����ٷμ���</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_overtime_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">���������ӱ�</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�Ĵ�</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_meals_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">����������</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">������</span></td>
    <td width="35%" height="30" align="right"><span class="style14C">\<%=formatnumber(incom_severance_pay,0)%>&nbsp;&nbsp;</span></td>
    <td width="15%" height="30" align="center" bgcolor="#eaeaea"><span class="style14BC">�ٹ��ð�������</span></td>
    <td width="35%" height="30" align="center"><span class="style14C"><%=pmg_emp_name%>&nbsp;</span></td>
  </tr>
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;�� ����� ���� �� ���� �� ���� ���濡 ���Ͽ� ��ź�, ������, �ټӼ���� �߰� �����Ѵ�.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;�衰���� �� ���������� ��� ���Ͽ��� ���ϱ����� ����Ͽ� ��� ���� �����Ѵ�<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;�� ���������� 1�� �̻� �ٹ��� �ڿ� ���Ͽ� �����Ѵ�.<br/></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>3. �����Ⱓ</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;���� �����Ⱓ�� �ű� ä���Ϸκ��� 3������ �ϸ�, �����Ⱓ �� �Ǵ� ���� �� �������ϴٰ� ����������<br/>&nbsp;&nbsp;&nbsp;&nbsp;������ ��쿡�� ����ä���� �ź��� �� �ִ�. ���Ⱓ�� �޿��� ��2���� �������� �ұ��ϰ� �� �޿�����<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=incom_first3_percent%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%�� �����Ѵ�.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>4. ������</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;���� �Ի��Ϸκ��� 1�� �̻� �ټ� �� ���� �� �ٷα��ع� �� ��ü���࿡ ���� �������ݿ� ���ԵǾ�<br/>&nbsp;&nbsp;&nbsp;&nbsp;�����Ѵ�.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>5. �ٷνð�</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;�� �ٷ��� : �� 5��, 1�� 8�ð�, �� 40�ð��� �������� �Ѵ�.<br/>&nbsp;&nbsp;&nbsp;&nbsp;�� �ٹ��ð� : ���� 9�ú��� ���� 6�ø� �������� �ϸ� �ްԽð��� 12�ú��� 13�ñ����̸硰���������� ��<br/>&nbsp;&nbsp;&nbsp;&nbsp;����ٹ��� �����Ѵ�.<br/>&nbsp;&nbsp;&nbsp;&nbsp;�� �� �ӱݾ� �� ����ٷμ����� �ٷ��� ����, �ӱݰ���� ���Ǽ� �� ����ڰ��� ���� ���Ͽ� �ٹ���<br/>&nbsp;&nbsp;&nbsp;&nbsp;�翬�� �߻� �����ϴ� ����������(����ٷε�)�� ���Ե� ������������ �ӱ����� �Ѵ�.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>6. ������ �ٷα��ع��� ������</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>7. �޿�����</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;���ɿ��� ���ϴ� ����,������ �ٷ��ڿ� ������ ������ �޿����� ������ �� �ִ�.</td>
  </tr> 
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>8. ��Ÿ����</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;���޼���, ����������, ��ź�, ���������� ���� ��Ÿ ������ �ش������ ��å�� ���� �������� �� ��<br/>&nbsp;&nbsp;&nbsp;&nbsp;������ �ش� ������ ���� ���� ���� �� ���� ��� �� �����Ѵ�.</td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>9. ��Ÿ�ٷλ���</strong></td>
  </tr>
  <tr>
     <td width="100%" height="30" align="left" class="style3">
     &nbsp;&nbsp;&nbsp;&nbsp;��õ��� �ƴ��� �ٷ������� �뵿�������, ȸ�� ������ �� �����ʿ� ������.</td>
  </tr>  
  <tr>
     <td width="100%" height="30" align="left" class="style3"><br/><strong>&nbsp;&nbsp;&nbsp;&nbsp;���� ���� ���� ���� �����ٷΰ�࿡ �����մϴ�.</strong></td>
  </tr>   
  <tr>
     <td width="100%" height="30" align="right" class="style3"><br/><%=mid(cstr(now()),1,4)%>��&nbsp;<%=mid(cstr(now()),6,2)%>��&nbsp;<%=mid(cstr(now()),9,2)%>��<br/><br/></td>
  </tr>  
</table>

<table width="690" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td width="50%" height="30" align="left" class="style3"><strong>��:&nbsp;(��)���̿��������</strong></td>
     <td width="50%" height="30" align="right" class="style3"><strong>��ǥ�̻�&nbsp;&nbsp;&nbsp;�� ����&nbsp;&nbsp;(��)</strong></td>
  </tr>
  <tr>
     <td width="50%" height="30" align="left" class="style3"><strong>��:&nbsp;<%=agree_grade%></strong></td>
     <td width="50%" height="30" align="right" class="style3"><strong>��&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;<%=agree_empname%>&nbsp;&nbsp;(��)</strong></td>
  </tr>
</table>
</p>	

 <%         
' 		sql = "insert into emp_agree(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment) values "
'		sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"')"
		
'		dbconn.execute(sql)
		
 %>     
  </form>
</body>
</html>
