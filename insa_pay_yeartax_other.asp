<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_other.asp"

y_final=Request("y_final")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

inc_yyyy = cint(mid(now(),1,4)) - 1

for i = 1 to 10
    family_tab(i,1) = ""
	family_tab(i,2) = ""
	family_tab(i,3) = ""
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_hous = Server.CreateObject("ADODB.Recordset")
Set rs_houm = Server.CreateObject("ADODB.Recordset")
Set rs_othe = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")
emp_person = cstr(rs_emp("emp_person1")) + cstr(rs_emp("emp_person2"))	
rs_emp.close()	

Sql = "select * from pay_yeartax_other where o_year = '"&inc_yyyy&"' and o_emp_no = '"&emp_no&"'"
rs_othe.Open Sql, Dbconn, 1
Set rs_othe = DbConn.Execute(SQL)
if not rs_othe.eof then
       u_type = "U"
       o_nps = rs_othe("o_nps")
	   o_nhis = rs_othe("o_nhis")
	   o_sosang = rs_othe("o_sosang")
	   o_chul2012 = rs_othe("o_chul2012")
	   o_chul2013 = rs_othe("o_chul2013")
	   o_chul2014 = rs_othe("o_chul2014")
	   o_woori = rs_othe("o_woori")
	   o_goyoung = rs_othe("o_goyoung")
	   o_chul_hap = o_chul2008 + o_chul2009
   else
       u_type = ""
       o_nps = 0
	   o_nhis = 0
	   o_sosang = 0
	   o_chul2012 = 0
	   o_chul2013 = 0
	   o_chul2014 = 0
	   o_woori = 0
	   o_goyoung = 0
	   o_chul_hap = 0
end if
rs_othe.close()	

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "����" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

title_line = "�������� - �׹��ǰ���(��Ÿ����) "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}		
			
			function chkfrm() {
//				if(document.frm.emp_ename.value =="") {
//					alert('���������� �Է��ϼ���');
//					frm.emp_ename.focus();
//					return false;}
					
				a=confirm('����Ͻðڽ��ϱ�?');
				if (a==true) {
					return true;
				}
				return false;
			} 
			
			function num_chk(txtObj){
				oo_nps = parseInt(document.frm.o_nps.value.replace(/,/g,""));	
				oo_nhis = parseInt(document.frm.o_nhis.value.replace(/,/g,""));	
				oo_sosang = parseInt(document.frm.o_sosang.value.replace(/,/g,""));	
				oo_chul2012 = parseInt(document.frm.o_chul2012.value.replace(/,/g,""));	
				oo_chul2013 = parseInt(document.frm.o_chul2013.value.replace(/,/g,""));	
				oo_chul2014 = parseInt(document.frm.o_chul2014.value.replace(/,/g,""));	
				oo_woori = parseInt(document.frm.o_woori.value.replace(/,/g,""));	
				oo_goyoung = parseInt(document.frm.o_goyoung.value.replace(/,/g,""));	
		
		        oo_chul_hap = oo_chul2012 + oo_chul2013 + oo_chul2014;
				
				oo_nps = String(oo_nps);
				num_len = oo_nps.length;
				sil_len = num_len;
				oo_nps = String(oo_nps);
				if (oo_nps.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_nps = oo_nps.substr(0,num_len -3) + "," + oo_nps.substr(num_len -3,3);
				if (sil_len > 6) oo_nps = oo_nps.substr(0,num_len -6) + "," + oo_nps.substr(num_len -6,3) + "," + oo_nps.substr(num_len -2,3);
				document.frm.o_nps.value = oo_nps;
				
				oo_nhis = String(oo_nhis);
				num_len = oo_nhis.length;
				sil_len = num_len;
				oo_nhis = String(oo_nhis);
				if (oo_nhis.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_nhis = oo_nhis.substr(0,num_len -3) + "," + oo_nhis.substr(num_len -3,3);
				if (sil_len > 6) oo_nhis = oo_nhis.substr(0,num_len -6) + "," + oo_nhis.substr(num_len -6,3) + "," + oo_nhis.substr(num_len -2,3);
				document.frm.o_nhis.value = oo_nhis;
				
				oo_sosang = String(oo_sosang);
				num_len = oo_sosang.length;
				sil_len = num_len;
				oo_sosang = String(oo_sosang);
				if (oo_sosang.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_sosang = oo_sosang.substr(0,num_len -3) + "," + oo_sosang.substr(num_len -3,3);
				if (sil_len > 6) oo_sosang = oo_sosang.substr(0,num_len -6) + "," + oo_sosang.substr(num_len -6,3) + "," + oo_sosang.substr(num_len -2,3);
				document.frm.o_sosang.value = oo_sosang;
				
				oo_chul2012 = String(oo_chul2012);
				num_len = oo_chul2012.length;
				sil_len = num_len;
				oo_chul2012 = String(oo_chul2012);
				if (oo_chul2012.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_chul2012 = oo_chul2012.substr(0,num_len -3) + "," + oo_chul2012.substr(num_len -3,3);
				if (sil_len > 6) oo_chul2012 = oo_chul2012.substr(0,num_len -6) + "," + oo_chul2012.substr(num_len -6,3) + "," + oo_chul2012.substr(num_len -2,3);
				document.frm.o_chul2012.value = oo_chul2012;
				
				oo_chul2013 = String(oo_chul2013);
				num_len = oo_chul2013.length;
				sil_len = num_len;
				oo_chul2013 = String(oo_chul2013);
				if (oo_chul2013.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_chul2013 = oo_chul2013.substr(0,num_len -3) + "," + oo_chul2013.substr(num_len -3,3);
				if (sil_len > 6) oo_chul2013 = oo_chul2013.substr(0,num_len -6) + "," + oo_chul2013.substr(num_len -6,3) + "," + oo_chul2013.substr(num_len -2,3);
				document.frm.o_chul2013.value = oo_chul2013;
				
				oo_chul2014 = String(oo_chul2014);
				num_len = oo_chul2014.length;
				sil_len = num_len;
				oo_chul2014 = String(oo_chul2014);
				if (oo_chul2014.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_chul2014 = oo_chul2014.substr(0,num_len -3) + "," + oo_chul2014.substr(num_len -3,3);
				if (sil_len > 6) oo_chul2014 = oo_chul2014.substr(0,num_len -6) + "," + oo_chul2014.substr(num_len -6,3) + "," + oo_chul2014.substr(num_len -2,3);
				document.frm.o_chul2014.value = oo_chul2014;
				
				oo_woori = String(oo_woori);
				num_len = oo_woori.length;
				sil_len = num_len;
				oo_woori = String(oo_woori);
				if (oo_woori.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_woori = oo_woori.substr(0,num_len -3) + "," + oo_woori.substr(num_len -3,3);
				if (sil_len > 6) oo_woori = oo_woori.substr(0,num_len -6) + "," + oo_woori.substr(num_len -6,3) + "," + oo_woori.substr(num_len -2,3);
				document.frm.o_woori.value = oo_woori;
				
				oo_goyoung = String(oo_goyoung);
				num_len = oo_goyoung.length;
				sil_len = num_len;
				oo_goyoung = String(oo_goyoung);
				if (oo_goyoung.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_goyoung = oo_goyoung.substr(0,num_len -3) + "," + oo_goyoung.substr(num_len -3,3);
				if (sil_len > 6) oo_goyoung = oo_goyoung.substr(0,num_len -6) + "," + oo_goyoung.substr(num_len -6,3) + "," + oo_goyoung.substr(num_len -2,3);
				document.frm.o_goyoung.value = oo_goyoung;
				
				oo_chul_hap = String(oo_chul_hap);
				num_len = oo_chul_hap.length;
				sil_len = num_len;
				oo_chul_hap = String(oo_chul_hap);
				if (oo_chul_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) oo_chul_hap = oo_chul_hap.substr(0,num_len -3) + "," + oo_chul_hap.substr(num_len -3,3);
				if (sil_len > 6) oo_chul_hap = oo_chul_hap.substr(0,num_len -6) + "," + oo_chul_hap.substr(num_len -6,3) + "," + oo_chul_hap.substr(num_len -2,3);
				document.frm.o_chul_hap.value = oo_chul_hap;
			}		
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_other_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="8%" >
                            <col width="9%" >
                            <col width="24%" >
                            <col width="25%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (�Ի���:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">�Ҽ�<input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true"></th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_company%> - <%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (�ͼӳ⵵:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                              </td>
						    </tr>
                             <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">����</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3;">�����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���ⱸ��</th>
                              <th>�ݾ�</th>
                              <th colspan="2">�������</th>
						    </tr>
                            <tr>
							  <th rowspan="9">��Ÿ����</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3;">���������� ���ο��ݺ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><input name="o_nps" type="text" id="o_nps" style="width:90px;text-align:right" value="<%=formatnumber(o_nps,0)%>" onKeyUp="num_chk(this);"></td>
                              <td colspan="2" class="left">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">���������� �ǰ������</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><input name="o_nhis" type="text" id="o_nhis" style="width:90px;text-align:right" value="<%=formatnumber(o_nhis,0)%>" onKeyUp="num_chk(this);"></td>
                              <td colspan="2" class="left">&nbsp;</td>
						    </tr>
                            <tr>
                              <th colspan="3" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�ұ�� �һ���� �����α� �ҵ����</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">���Աݾ�</th>
                              <td class="right"><input name="o_sosang" type="text" id="o_sosang" style="width:90px;text-align:right" value="<%=formatnumber(o_sosang,0)%>" onKeyUp="num_chk(this);"></td>
                              <td colspan="2" class="left">����Ⱓ�� 1�� �̻��� �ұ�� �һ���� ��ǥ�ڰ� ���ѿ������� ������ ���.
                              </td>
						    </tr>
                            <tr>
                              <th rowspan="4" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������� ���ڰ���</th>
                              <th colspan="2" style="background:#f8f8f8; border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">2012�� �������ں�</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">�������ڱݾ�</th>
                              <td class="right"><input name="o_chul2012" type="text" id="o_chul2012" style="width:90px;text-align:right" value="<%=formatnumber(o_chul2012,0)%>" onKeyUp="num_chk(this);"></td>
                              <td rowspan="3" colspan="2" class="left">�߼ұ��â���������� �Ǵ� ��ó��� � ���� �Ǵ� ���� �� 2���� �Ǵ� ���� ���ϴ� ���������� �ϳ��� ���������� �����Ͽ� ����
                              </td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2013�� �������ں�</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�������ڱݾ�</th>
                              <td class="right"><input name="o_chul2013" type="text" id="o_chul2013" style="width:90px;text-align:right" value="<%=formatnumber(o_chul2013,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">2014������ �������ں�</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�������ڱݾ�</th>
                              <td class="right"><input name="o_chul2014" type="text" id="o_chul2014" style="width:90px;text-align:right" value="<%=formatnumber(o_chul2014,0)%>" onKeyUp="num_chk(this);"></td>
						    </tr>
                            <tr>
                              <th colspan="2" style=" border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�������� ���ڰ��� ��</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td class="right"><input name="o_chul_hap" type="text" id="o_chul_hap" style="width:90px;text-align:right" value="<%=formatnumber(o_chul_hap,0)%>" readonly="true"></td>
						    </tr>
                            <tr>
                              <th colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">�츮���� �⿬�� �ҵ����</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�⿬�ݾ�</th>
                              <td class="right"><input name="o_woori" type="text" id="o_woori" style="width:90px;text-align:right" value="<%=formatnumber(o_woori,0)%>" onKeyUp="num_chk(this);"></td>
                              <td colspan="2" class="left">�츮�������տ��� �ڻ��ָ� ����ϱ� ���Ͽ� �츮�������տ� �⿬�� �ݾ�
                              </td>
						    </tr>
                            <tr>
                              <th colspan="3" style="border-left:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">��������߼ұ�� �ٷ��� �ҵ����</th>
                              <th style="background:#f8f8f8; border-bottom:1px solid #e3e3e3;">�ӱݻ谨��</th>
                              <td class="right"><input name="o_goyoung" type="text" id="o_goyoung" style="width:90px;text-align:right" value="<%=formatnumber(o_goyoung,0)%>" onKeyUp="num_chk(this);"></td>
                              <td colspan="2" class="left">2011����� ������� �߼ұ�� ��ñٷ��� �ӱݻ谨���� 50% ����(�� 1,000�����ѵ�)
                              </td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">�� �ұ�� �һ���� �����α� : �����α� ���Ծ�(�� 300���� �ѵ�).<br>
                �� ������������ : ���ڱݾ��� 10%(�ҵ�ݾ��� 30%�ѵ�).<br>
                �� �츮�������� �⿬�� : �⿬�ݾ� 400���� �ѵ�.<br>
                �� ��������߼ұ�� �ٷ��� �ҵ���� : (�������������� �ش� �ٷ��� ���� �ӱ��Ѿ� - �ش� ���������� �ش�ٷ��� ���� �ӱ��Ѿ�)*50%.</h3>
				</div>
                <br>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  <td width="100%">
                    <div align=center>
              <% if y_final <> "Y" then  %>                      
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
			  <%   end if  %>      				   
                    </div>
				  </td>	
                  </tr>
				</table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="emp_person" value="<%=emp_person%>" ID="Hidden1">                 
			</form>
		</div>				
	</div>        				
	</body>
</html>

