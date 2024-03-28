<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim family_tab(10,3)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_deduction.asp"

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
Set rs_dedu = Server.CreateObject("ADODB.Recordset")
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

old_de_id = ""

Sql = "select * from pay_yeartax_deduction where de_year = '"&inc_yyyy&"' and de_emp_no = '"&emp_no&"'"
rs_dedu.Open Sql, Dbconn, 1
Set rs_dedu = DbConn.Execute(SQL)
if not rs_dedu.eof then
       u_type = "U"
       de_id = rs_dedu("de_id")
	   old_de_id = rs_dedu("de_id")
	   young_fdate = rs_dedu("young_fdate")
	   young_ldate = rs_dedu("young_ldate")
	   de_person_no = rs_dedu("de_person_no")
	   de_wonchen = rs_dedu("de_wonchen")
	   de_tax_s = rs_dedu("de_tax_s")
	   de_tax_w = rs_dedu("de_tax_w")
	   de_tax_nation = rs_dedu("de_tax_nation")
	   de_tax_date = rs_dedu("de_tax_date")
	   de_report_date = rs_dedu("de_report_date")
	   de_office = rs_dedu("de_office")
	   de_stay = rs_dedu("de_stay")
	   de_position = rs_dedu("de_position")
	   if young_fdate = "1900-01-01" then
	       young_fdate = ""
	   end if
	   if young_ldate = "1900-01-01" then
	       young_ldate = ""
	   end if
	   if de_tax_date = "1900-01-01" then
	       de_tax_date = ""
	   end if
	   if de_report_date = "1900-01-01" then
	       de_report_date = ""
	   end if
   else
       u_type = ""
       de_id = ""
	   young_fdate = ""
	   young_ldate = ""
	   de_person_no = ""
	   de_wonchen = 0
	   de_tax_s = 0
	   de_tax_w = 0
	   de_tax_nation = ""
	   de_tax_date = ""
	   de_report_date = ""
	   de_office = ""
	   de_stay = ""
	   de_position = ""
end if
rs_dedu.close()	

sql = "select * from pay_yeartax_family where f_year = '"&inc_yyyy&"' and f_emp_no = '"&emp_no&"' ORDER BY f_emp_no,f_pseq,f_person_no ASC"
rs_fami.Open Sql, Dbconn, 1
Set rs_fami = DbConn.Execute(SQL)
i = 0
do until rs_fami.eof
   if rs_fami("f_rel") = "본인" or rs_fami("f_wife") = "Y" or rs_fami("f_age20") = "Y" or rs_fami("f_age60") = "Y" or rs_fami("f_old") = "Y" then
		  i = i + 1
		  family_tab(i,1) = rs_fami("f_rel")
	      family_tab(i,2) = rs_fami("f_family_name")
	      family_tab(i,3) = rs_fami("f_person_no")
	end if
	rs_fami.MoveNext()
loop
rs_fami.close()

title_line = "연말정산 - 세액감면/공제 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=young_fdate%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=young_ldate%>" );
			});	
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=de_tax_date%>" );
			});	
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=de_report_date%>" );
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
//					alert('영문성명을 입력하세요');
//					frm.emp_ename.focus();
//					return false;}
					
				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			} 
			
			function num_chk(txtObj){
				d_wonchen = parseInt(document.frm.de_wonchen.value.replace(/,/g,""));	
				d_tax_s = parseInt(document.frm.de_tax_s.value.replace(/,/g,""));	
				d_tax_w = parseInt(document.frm.de_tax_w.value.replace(/,/g,""));	
		
				d_wonchen = String(d_wonchen);
				num_len = d_wonchen.length;
				sil_len = num_len;
				d_wonchen = String(d_wonchen);
				if (d_wonchen.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) d_wonchen = d_wonchen.substr(0,num_len -3) + "," + d_wonchen.substr(num_len -3,3);
				if (sil_len > 6) d_wonchen = d_wonchen.substr(0,num_len -6) + "," + d_wonchen.substr(num_len -6,3) + "," + d_wonchen.substr(num_len -2,3);
				document.frm.de_wonchen.value = d_wonchen;
				
				d_tax_s = String(d_tax_s);
				num_len = d_tax_s.length;
				sil_len = num_len;
				d_tax_s = String(d_tax_s);
				if (d_tax_s.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) d_tax_s = d_tax_s.substr(0,num_len -3) + "," + d_tax_s.substr(num_len -3,3);
				if (sil_len > 6) d_tax_s = d_tax_s.substr(0,num_len -6) + "," + d_tax_s.substr(num_len -6,3) + "," + d_tax_s.substr(num_len -2,3);
				document.frm.de_tax_s.value = d_tax_s;
				
				d_tax_w = String(d_tax_w);
				num_len = d_tax_w.length;
				sil_len = num_len;
				d_tax_w = String(d_tax_w);
				if (d_tax_w.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) d_tax_w = d_tax_w.substr(0,num_len -3) + "," + d_tax_w.substr(num_len -3,3);
				if (sil_len > 6) d_tax_w = d_tax_w.substr(0,num_len -6) + "," + d_tax_w.substr(num_len -6,3) + "," + d_tax_w.substr(num_len -2,3);
				document.frm.de_tax_w.value = d_tax_w;
				
			}		
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_deduction_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">성명(<%=emp_no%><input name="emp_no" type="hidden" value="<%=emp_no%>" style="width:40px" readonly="true">)</th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_name%>
                                <input name="emp_name" type="hidden" value="<%=emp_name%>" style="width:50px" readonly="true">
                                (입사일:<%=emp_in_date%>
                                <input name="emp_in_date" type="hidden" value="<%=emp_in_date%>" style="width:70px" readonly="true">)
                              </td>
							  <th style=" border-bottom:1px solid #e3e3e3;">소속<input name="emp_company" type="hidden" value="<%=emp_company%>" style="width:90px" readonly="true"></th>
							  <td colspan="3" class="left" style=" border-bottom:1px solid #e3e3e3;"><%=emp_company%> - <%=emp_org_name%>
                                <input name="emp_org_name" type="hidden" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                - <%=emp_grade%>
                                <input name="emp_grade" type="hidden" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                - <%=emp_position%>
                                <input name="emp_position" type="hidden" value="<%=emp_position%>" style="width:70px" readonly="true">
                                (귀속년도:
                                <input name="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:40px; text-align:center" readonly="true">)
                              </td>
						    </tr>
                             <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">구분</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3;">지출명세</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">지출구분</th>
                              <th>금액</th>
                              <th colspan="2">공제요건</th>
						    </tr>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">세액감면</th>
                              <th colspan="3" style=" border-bottom:1px solid #e3e3e3;">중소기업 취업 청년 감면</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">취업일</th>
                              <td class="left"><input name="young_fdate" type="text" value="<%=young_fdate%>" style="width:90px;text-align:center" id="datepicker" readonly="true"></td>
                              <th style=" border-top:1px solid #e3e3e3; border-bottom:1px solid #e3e3e3;">감면기간 종료일</th>
                              <td class="left"><input name="young_ldate" type="text" value="<%=young_ldate%>" style="width:90px;text-align:center" id="datepicker1" readonly="true"></td>
						    </tr>
                            <tr>
							  <th rowspan="6">세액공제</th>
                              <th rowspan="6" colspan="3" style=" border-bottom:1px solid #e3e3e3;">외국납부세액</th>
                              <th style=" border-bottom:1px solid #e3e3e3;">국외원천소득</th>
                              <td class="left"><input name="de_wonchen" type="text" id="de_wonchen" style="width:90px;text-align:right" value="<%=formatnumber(de_wonchen,0)%>" onKeyUp="num_chk(this);"></td>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">납세액(외화)</th>
                              <td class="left"><input name="de_tax_s" type="text" id="de_tax_s" style="width:90px;text-align:right" value="<%=formatnumber(de_tax_s,0)%>" onKeyUp="num_chk(this);"></td>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">납세액(원화)</th>
                              <td class="left"><input name="de_tax_w" type="text" id="de_tax_w" style="width:90px;text-align:right" value="<%=formatnumber(de_tax_w,0)%>" onKeyUp="num_chk(this);"></td>
                              <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <td style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;">&nbsp;</td>
						    </tr>
                            <tr>
                              <th style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">납세국명</th>
                              <td class="left"><input name="de_tax_nation" type="text" id="de_tax_nation" value="<%=de_tax_nation%>" style="width:130px" ></td>
                              <th style="border-left:1px solid #e3e3e3;">납부일</th>
                              <td class="left" style=" border-bottom:1px solid #e3e3e3;"><input name="de_tax_date" type="text" value="<%=de_tax_date%>" style="width:90px;text-align:center" id="datepicker2" readonly="true"></td>
						    </tr>
                            <tr>
                              <th style="border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">신청서제출일</th>
                              <td class="left"><input name="de_report_date" type="text" value="<%=de_report_date%>" style="width:90px;text-align:center" id="datepicker3" readonly="true"></td>
                              <th style=" border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">국외근무처</th>
                              <td class="left" style=" border-bottom:1px solid #e3e3e3;"><input name="de_office" type="text" id="de_office" value="<%=de_office%>" style="width:130px" ></td>
						    </tr>
                            <tr>
                              <th style="border-bottom:1px solid #e3e3e3; border-left:1px solid #e3e3e3;">근무기간</th>
                              <td class="left"><input name="de_stay" type="text" id="de_stay" value="<%=de_stay%>" style="width:90px" >&nbsp;개월</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">직책</th>
                              <td class="left" style=" border-bottom:1px solid #e3e3e3;"><input name="de_position" type="text" id="de_position" value="<%=de_position%>" style="width:130px"></td>
						    </tr>
						</thead>
						<tbody>
					</table>
				<h3 class="stit">※ 중소기업 취업 청연 감면 : .<br>
                ※ 외국납세액 : 국외원천소득에 대해서 외국에서 납부한 세액이 있는 경우 공재대상.<br>
                ※ 외국에서 납부하였거나 납부할 세액.<br>
                ※ 근로소득산출액*(국외근로소득금액/근로소득금액).</h3>

				</div>
                <br>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  <td width="100%">
                    <div align=center>
              <% if y_final <> "Y" then  %>                      
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
			  <%   end if  %>      						    
                    </div>
				  </td>	
                  </tr>
				</table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="emp_person" value="<%=emp_person%>" ID="Hidden1">  
                <input type="hidden" name="de_id" value="<%=de_id%>" ID="Hidden1">    
                <input type="hidden" name="old_de_id" value="<%=old_de_id%>" ID="Hidden1">                 
			</form>
		</div>				
	</div>        				
	</body>
</html>

