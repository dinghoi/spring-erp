<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(6,20)
dim amount_tab(3,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

view_condi=Request("view_condi")
buy_goods_type=Request("goods_type")
buy_no=Request("buy_no")
buy_seq=Request("buy_seq")
buy_date=Request("buy_date")

curr_date = mid(cstr(now()),1,10)
buy_date = curr_date
buy_company = ""
buy_saupbu = ""
buy_org_code = ""
buy_org_name = ""
buy_emp_no = ""
buy_emp_name = ""
buy_bill_collect = "현금"
buy_collect_due_date = curr_date
buy_trade_no = ""
buy_trade_name = ""
buy_trade_person = ""
buy_trade_email = ""
buy_out_method = ""
buy_out_request_date = ""
buy_price = 0
buy_cost = 0
buy_cost_vat = 0
buy_memo = ""
buy_ing = "0"
buy_sign_yn = "N"
buy_sign_no = ""
buy_sign_date = "0000-00-00"

pummok_cnt = 0

path_name = "/met_upload"

for i = 1 to 6
	for j = 1 to 20
		pummok_tab(i,j) = ""
	next
next
for i = 1 to 3
	for j = 1 to 20
		amount_tab(i,j) = 0
	next
next

' response.write(reg_date)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_bg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&user_id&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_no = rs_emp("emp_no")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
		buy_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		buy_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
   else
		emp_name = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
end if
rs_emp.close()


title_line =  " 구매품의 등록 "

if u_type = "U" then

	Sql="select * from met_buy where (buy_no = '"&buy_no&"') and (buy_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"')"
	Set rs=DbConn.Execute(Sql)

	buy_no = rs("buy_no")
	buy_date = rs("buy_date")
	buy_goods_type = rs("buy_goods_type")
	buy_company = rs("buy_company")
	buy_bonbu = rs("buy_bonbu")
	buy_saupbu = rs("buy_saupbu")
	buy_team = rs("buy_team")
	buy_org_code = rs("buy_org_code")
	buy_org_name = rs("buy_org_name")
	buy_emp_no = rs("buy_emp_no")
	buy_emp_name = rs("buy_emp_name")
	buy_bill_collect = rs("buy_bill_collect")
    buy_collect_due_date = rs("buy_collect_due_date")
	buy_trade_no = rs("buy_trade_no")
    buy_trade_name = rs("buy_trade_name")
    buy_trade_person = rs("buy_trade_person")
	buy_trade_email = rs("buy_trade_email")
    buy_out_method = rs("buy_out_method")
    buy_out_request_date = rs("buy_out_request_date")
    buy_price = rs("buy_price")
    buy_cost = rs("buy_cost")
    buy_cost_vat = rs("buy_cost_vat")
    buy_memo = rs("buy_memo")
    buy_ing = rs("buy_ing")
	buy_sign_yn = rs("buy_sign_yn")
	buy_sign_no = rs("buy_sign_no")
	buy_sign_date = rs("buy_sign_date")
	buy_att_file = rs("buy_att_file")

	if buy_out_request_date = "0000-00-00" then
	      buy_out_request_date = ""
	end if
	if buy_collect_due_date = "0000-00-00" then
	      buy_collect_due_date = ""
	end if
	if buy_sign_date = "0000-00-00" then
	      buy_sign_date = ""
	end if
	
	rs.close()

	j = 0
	Sql="select * from met_buy_goods where (bg_no = '"&buy_no&"') and (bg_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("bg_goods_type")
		pummok_tab(2,j) = rs("bg_goods_gubun")
		pummok_tab(3,j) = rs("bg_goods_code")
		pummok_tab(4,j) = rs("bg_goods_name")
		pummok_tab(5,j) = rs("bg_standard")
		pummok_tab(6,j) = rs("bg_seq")
		amount_tab(1,j) = rs("bg_qty")
		amount_tab(2,j) = rs("bg_unit_cost")
		amount_tab(3,j) = rs("bg_buy_amt")
		rs.movenext()
	loop
	pummok_cnt = j
    
	title_line =  " 구매품의 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
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
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=buy_out_request_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=buy_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=bill_due_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=bill_issue_date%>" );
			});	  
			$(function() {    $( "#datepicker4" ).datepicker();
												$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker4" ).datepicker("setDate", "<%=buy_collect_due_date%>" );
			});	  
			$(function() {    $( "#datepicker5" ).datepicker();
												$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker5" ).datepicker("setDate", "<%=collect_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm1()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
						
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function chkfrm1() {
//				if(document.frm.sales_company.value == "") {
//					alert('매출회사를 선택하세요');
//					frm.sales_company.focus();
//					return false;}
				if(document.frm.buy_goods_type.value == "") {
					alert('구매용도를 선택하세요');
					frm.buy_goods_type.focus();
					return false;}
				if(document.frm.trade_name.value == "") {
					alert('구매처를 선택하세요');
					frm.trade_name.focus();
					return false;}
				if(document.frm.trade_no.value == "") {
					alert('구매처를 선택하세요');
					frm.trade_no.focus();
					return false;}
				if(document.frm.trade_person.value == "") {
					alert('구매처 담당자를 입력하세요');
					frm.trade_person.focus();
					return false;}
				if(document.frm.trade_email.value == "") {
					alert('구매처 담당자 이메일을 입력하세요');
					frm.trade_email.focus();
					return false;}

				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.bill_collect[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("대금지금방법을 선택하세요");
					return false;
				}	
						
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function pop_pummok() 
			{ 
				var code_ary = new Array();
//				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value;
				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value;
				code_ary[1] = document.frm.goods_code2.value + "/" + document.frm.goods_standard2.value;
				code_ary[2] = document.frm.goods_code3.value + "/" + document.frm.goods_standard3.value;
				code_ary[3] = document.frm.goods_code4.value + "/" + document.frm.goods_standard4.value;
				code_ary[4] = document.frm.goods_code5.value + "/" + document.frm.goods_standard5.value;
				code_ary[5] = document.frm.goods_code6.value + "/" + document.frm.goods_standard6.value;
				code_ary[6] = document.frm.goods_code7.value + "/" + document.frm.goods_standard7.value;
				code_ary[7] = document.frm.goods_code8.value + "/" + document.frm.goods_standard8.value;
				code_ary[8] = document.frm.goods_code9.value + "/" + document.frm.goods_standard9.value;
				code_ary[9] = document.frm.goods_code10.value + "/" + document.frm.goods_standard10.value;
				goods_type = document.frm.buy_goods_type.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_select.asp?code_ary='+code_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				
				var url = "met_goods_select.asp?code_ary="+code_ary+'&goods_type='+goods_type;
//				var url = "met_goods_select.asp?code_ary="+code_ary;
				pop_Window(url,'품목선택','scrollbars=yes,width=700,height=400');
				
				
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
//				code_ary[0] = document.frm.goods_code1.value + "/" + document.frm.goods_standard1.value + "/" + document.frm.qty1.value + "/" + document.frm.buy_cost1.value + "/" + document.frm.buy_tot1.value;
				code_ary[0] = document.frm.goods_code1.value
				code_ary[1] = document.frm.goods_code2.value
				code_ary[2] = document.frm.goods_code3.value
				code_ary[3] = document.frm.goods_code4.value
				code_ary[4] = document.frm.goods_code5.value
				code_ary[5] = document.frm.goods_code6.value
				code_ary[6] = document.frm.goods_code7.value
				code_ary[7] = document.frm.goods_code8.value
				code_ary[8] = document.frm.goods_code9.value
				code_ary[9] = document.frm.goods_code10.value

				if (document.frm.del_check1.checked == true) {
					del_ary[0] = 'Y'; } 					
				if (document.frm.del_check1.checked == false) {
					del_ary[0] = 'N'; } 
				if (document.frm.del_check2.checked == true) {
					del_ary[1] = 'Y'; } 					
				if (document.frm.del_check2.checked == false) {
					del_ary[1] = 'N'; } 
				if (document.frm.del_check3.checked == true) {
					del_ary[2] = 'Y'; } 					
				if (document.frm.del_check3.checked == false) {
					del_ary[2] = 'N'; } 
				if (document.frm.del_check4.checked == true) {
					del_ary[3] = 'Y'; } 					
				if (document.frm.del_check4.checked == false) {
					del_ary[3] = 'N'; } 
				if (document.frm.del_check5.checked == true) {
					del_ary[4] = 'Y'; } 					
				if (document.frm.del_check5.checked == false) {
					del_ary[4] = 'N'; } 
				if (document.frm.del_check6.checked == true) {
					del_ary[5] = 'Y'; } 					
				if (document.frm.del_check6.checked == false) {
					del_ary[5] = 'N'; } 
				if (document.frm.del_check7.checked == true) {
					del_ary[6] = 'Y'; } 					
				if (document.frm.del_check7.checked == false) {
					del_ary[6] = 'N'; } 
				if (document.frm.del_check8.checked == true) {
					del_ary[7] = 'Y'; } 					
				if (document.frm.del_check8.checked == false) {
					del_ary[7] = 'N'; } 
				if (document.frm.del_check9.checked == true) {
					del_ary[8] = 'Y'; } 					
				if (document.frm.del_check9.checked == false) {
					del_ary[8] = 'N'; } 
				if (document.frm.del_check10.checked == true) {
					del_ary[9] = 'Y'; } 					
				if (document.frm.del_check10.checked == false) {
					del_ary[9] = 'N'; } 
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'', '선택상품삭제', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
//				alert("삭제되었습니다 !!!");
//				NumCal();
				
				var url = "met_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary;				
				pop_Window(url,'선택상품삭제','scrollbars=yes,width=600,height=400');
				alert("삭제되었습니다 !!!");
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();
			var buy_ary = new Array();
			var buy_tot = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				buy_cal = qty_ary[j] * buy_ary[j];
				buy_cal = String(buy_cal);
				num_len = buy_cal.length;
				sil_len = num_len;
				if (buy_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) buy_cal = buy_cal.substr(0,num_len -3) + "," + buy_cal.substr(num_len -3,3);
				if (sil_len > 6) buy_cal = buy_cal.substr(0,num_len -6) + "," + buy_cal.substr(num_len -6,3) + "," + buy_cal.substr(num_len -2,3);
				if (sil_len > 9) buy_cal = buy_cal.substr(0,num_len -9) + "," + buy_cal.substr(num_len -9,3) + "," + buy_cal.substr(num_len -5,3) + "," + buy_cal.substr(num_len -1,3);
				eval("document.frm.buy_tot" + j + ".value = buy_cal");
				
			}

			buy_tot_cost = 0;
			for (j=1;j<21;j++) {
				buy_tot_cost = buy_tot_cost + buy_tot[j];
			}
			
			tot_cal = buy_tot_cost;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost.value = tot_cal");
			
			buy_vat = parseInt(buy_tot_cost * (10 / 100));
			buy_hap = buy_tot_cost + buy_vat;
			
			buy_vat = String(buy_vat);
			num_len = buy_vat.length;
			sil_len = num_len;
			if (buy_vat.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_vat = buy_vat.substr(0,num_len -3) + "," + buy_vat.substr(num_len -3,3);
			if (sil_len > 6) buy_vat = buy_vat.substr(0,num_len -6) + "," + buy_vat.substr(num_len -6,3) + "," + buy_vat.substr(num_len -2,3);
			if (sil_len > 9) buy_vat = buy_vat.substr(0,num_len -9) + "," + buy_vat.substr(num_len -9,3) + "," + buy_vat.substr(num_len -5,3) + "," + buy_vat.substr(num_len -1,3);
			eval("document.frm.buy_tot_cost_vat.value = buy_vat");
			
			buy_hap = String(buy_hap);
			num_len = buy_hap.length;
			sil_len = num_len;
			if (buy_hap.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) buy_hap = buy_hap.substr(0,num_len -3) + "," + buy_hap.substr(num_len -3,3);
			if (sil_len > 6) buy_hap = buy_hap.substr(0,num_len -6) + "," + buy_hap.substr(num_len -6,3) + "," + buy_hap.substr(num_len -2,3);
			if (sil_len > 9) buy_hap = buy_hap.substr(0,num_len -9) + "," + buy_hap.substr(num_len -9,3) + "," + buy_hap.substr(num_len -5,3) + "," + buy_hap.substr(num_len -1,3);
			eval("document.frm.buy_tot_price.value = buy_hap");

			if (txtObj.value.length<1) {
				txtObj.value=txtObj.value.replace(/,/g,"");
				txtObj.value=txtObj.value.replace(/\D/g,"");
			}
			var num = txtObj.value;
			if (num == "--" ||  num == "." ) num = "";
			if (num != "" ) {
				temp=new String(num);
				if(temp.length<1) return "";
							
				// 음수처리
				if(temp.substr(0,1)=="-") minus="-";
					else minus="";
							
				// 소수점이하처리
				dpoint=temp.search(/\./);
						
				if(dpoint>0)
				{
				// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
				dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
				temp=temp.substr(0,dpoint);
				}else dpointVa="";
							
				// 숫자이외문자 삭제
				temp=temp.replace(/\D/g,"");
				zero=temp.search(/[1-9]/);
						
				if(zero==-1) return "";
				else if(zero!=0) temp=temp.substr(zero);
							
				if(temp.length<4) return minus+temp+dpointVa;
				buf="";
				while (true)
				{
				if(temp.length<3) { buf=temp+buf; break; }
					
				buf=","+temp.substr(temp.length-3)+buf;
				temp=temp.substr(0, temp.length-3);
				}
				if(buf.substr(0,1)==",") buf=buf.substr(1);
						
				//return minus+buf+dpointVa;
				txtObj.value = minus+buf+dpointVa;
			}else txtObj.value = "0";						
		}
		function pummok_list_view() {
				pummok_cnt = parseInt(document.frm.pummok_cnt.value);
				for (j=1;j<pummok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.method = "post";
					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_buy_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
		</script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_buy_add_save.asp" enctype="multipart/form-data">
                    <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>구매용도</th>
							  <td class="left">
							<%
                                Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_goods_type" id="buy_goods_type" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until Rs_etc.eof 
                            %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If buy_goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                            <%
                                    Rs_etc.movenext()  
                                loop 
                                Rs_etc.Close()
                            %>
                                </select>
                              </td>
                              <th>구매회사</th>
							  <td class="left">
							<%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_company" id="buy_company" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = buy_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>구매사업부</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="buy_saupbu" id="buy_saupbu" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = buy_saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>구매담장자</th>
							  <td class="left"><%=emp_name%>(<%=emp_no%>)&nbsp;-&nbsp;<%=org_name%></td>
						    </tr>
							<tr>
							  <th>구매처</th>
							  <td class="left"><input name="trade_name" type="text" value="<%=buy_trade_name%>" readonly="true" style="width:120px">
                              <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="buy"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              </td>
							  <th>사업자번호</th>
							  <td class="left"><input name="trade_no" type="text" value="<%=buy_trade_no%>" readonly="true" style="width:150px"></td>
							  <th>구매처<br>담당자</th>
							  <td class="left"><input name="trade_person" type="text" value="<%=buy_trade_person%>" id="trade_person" style="width:120px; ime-mode:active" onKeyUp="checklength(this,20);"></td>
                              <th>이메일</th>
							  <td class="left"><input name="trade_email" type="text" value="<%=buy_trade_email%>" id="trade_email" style="width:180px;" onKeyUp="checklength(this,30);"></td>
						    </tr>
							<tr>
							  <th>구매일자</th>
							  <td class="left"><input name="buy_date" type="text" value="<%=buy_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
							  <th>대금<br>지급방법</th>
							  <td colspan="3" class="left">
                              <input type="radio" name="bill_collect" value="현금" <% if buy_bill_collect = "현금" then %>checked<% end if %> style="width:40px" id="Radio3">현금
  							  <input type="radio" name="bill_collect" value="어음" <% if buy_bill_collect = "어음" then %>checked<% end if %> style="width:40px" id="Radio4">어음
                              <input type="radio" name="bill_collect" value="카드" <% if buy_bill_collect = "카드" then %>checked<% end if %> style="width:40px" id="Radio3">카드
  							  <input type="radio" name="bill_collect" value="외환" <% if buy_bill_collect = "외환" then %>checked<% end if %> style="width:40px" id="Radio4">외환
                              </td>
							  <th>지급예정일</th>
							  <td class="left"><input name="collect_due_date" type="text" value="<%=buy_collect_due_date%>" style="width:80px;text-align:center" id="datepicker4"></td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td class="left" colspan="8" ><textarea name="buy_memo" rows="3" id="textarea"><%=buy_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 구매품의 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="16%" >
							<col width="14%" >
							<col width="11%" >
							<col width="11%" >
							<col width="11%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first; left" colspan="9" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">품목코드선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col">선택</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">수량</th>
								<th scope="col">구입단가</th>
								<th scope="col">구입금액</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" value="<%=pummok_tab(1,i)%>" style="width:70px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:120px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:80px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:140px" readonly="true"></td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:130px"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" value="<%=formatnumber(amount_tab(1,i),0)%>"  style="width:80px;text-align:right" onKeyUp="NumCal(this);"></td>
								<td><input name="buy_cost<%=i%>" type="text" id="buy_cost<%=i%>" value="<%=formatnumber(amount_tab(2,i),0)%>" style="width:90px;text-align:right" onKeyUp="NumCal(this);" ></td>
								<td><input name="buy_tot<%=i%>" type="text" disabled id="buy_tot<%=i%>" value="<%=formatnumber(amount_tab(3,i),0)%>"  style="width:90px;text-align:right" readonly="true"></td>
							</tr>
						<%
							next
						%>
						</tbody>
					</table>                    
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="21%" >
							<col width="13%" >
							<col width="21%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>구입총액</th>
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_price,0)%>" readonly="true"></td>
							  <th>구입금액</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost,0)%>" readonly="true"></td>
							  <th>구입부가세</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost_vat,0)%>" readonly="true"></td>
						    </tr>
							<tr>
							  <th>첨부</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=buy_att_file%>"><%=buy_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>
<br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"></span>
			<% end if	%>        
                    
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="buy_sign_yn" value="<%=buy_sign_yn%>" ID="Hidden1">
                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="Hidden1">
                <input type="hidden" name="emp_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="emp_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="emp_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="emp_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="emp_org_code" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="emp_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                
                <input type="hidden" name="old_buy_no" value="<%=buy_no%>">
                <input type="hidden" name="old_buy_seq" value="<%=buy_seq%>">
				<input type="hidden" name="old_buy_date" value="<%=buy_date%>">
				<input type="hidden" name="old_buy_goods_type" value="<%=buy_goods_type%>">
				<input type="hidden" name="old_att_file" value="<%=buy_att_file%>">
                
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
                </div>
			</div>
		</div>
   	</body>
</html>
