<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

u_type = request("u_type")
slip_no = request("slip_no")

slip_id = "W"
sales_company = ""
sales_saupbu = ""
trade_no = ""
trade_name = ""
trade_person = ""
trade_email = ""
out_method = ""
out_request_date = ""
sales_date = ""
sales_yn = ""

bill_due_date = ""
bill_issue_yn = ""
bill_issue_date = ""
bill_collect = ""
collect_due_date = ""
collect_stat = ""
collect_date = ""
slip_memo = ""

sales_price = 0
sales_cost = 0
sales_vat = 0
buy_price = 0
buy_cost = 0
buy_vat = 0
margin_cost = 0

curr_date = mid(cstr(now()),1,10)

title_line = "대기 전표 등록"
if u_type = "U" then

	Sql="select * from sales_slip where slip_no = '"&slip_no&"'"
	Set rs=DbConn.Execute(Sql)

	slip_id = rs("slip_id")
	sales_company = rs("sales_company")
	sales_saupbu = rs("sales_saupbu")
	trade_no = rs("trade_no")
	trade_name = rs("trade_name")
	trade_person = rs("trade_person")
	trade_email = rs("trade_email")
	out_method = rs("out_method")
	out_request_date = rs("out_request_date")
	sales_date = rs("sales_date")
	sales_yn = rs("sales_yn")
	
	bill_due_date = rs("bill_due_date")
	bill_issue_yn = rs("bill_issue_yn")
	bill_issue_date = rs("bill_issue_date")
	bill_collect = rs("bill_collect")
	collect_due_date = rs("collect_due_date")
	collect_stat = rs("collect_stat")
	collect_date = rs("collect_date")
	slip_memo = rs("slip_memo")
	
	sales_price = rs("sales_price")
	sales_cost = rs("sales_cost")
	sales_vat = rs("sales_cost_vat")
	buy_price = rs("buy_price")
	buy_cost = rs("buy_cost")
	buy_vat = rs("buy_cost_vat")
	margin_cost = rs("margin_cost")

	rs.close()

	title_line = "대기 전표 수정"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=out_request_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=sales_date%>" );
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
												$( "#datepicker4" ).datepicker("setDate", "<%=collect_due_date%>" );
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
				if (chkfrm()) {
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
//				if(document.frm.sales_saupbu.value == "") {
//					alert('매출사업부를 선택하세요');
//					frm.sales_saupbu.focus();
//					return false;}
				if(document.frm.trade_name.value == "") {
					alert('거래처를 선택하세요');
					frm.trade_name.focus();
					return false;}
				if(document.frm.trade_no.value == "") {
					alert('거래처를 선택하세요');
					frm.trade_no.focus();
					return false;}
				if(document.frm.trade_person.value == "") {
					alert('거래처 담당자를 입력하세요');
					frm.trade_person.focus();
					return false;}
				if(document.frm.trade_email.value == "") {
					alert('계산서 메일을 입력하세요');
					frm.trade_email.focus();
					return false;}

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.slip_id[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("전표유형을 선택하세요");
					return false;
				}	

				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.bill_collect[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("수금방법을 선택하세요");
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
				code_ary[0] = document.frm.pummok_code1.value
				code_ary[1] = document.frm.pummok_code2.value
				code_ary[2] = document.frm.pummok_code3.value
				code_ary[3] = document.frm.pummok_code4.value
				code_ary[4] = document.frm.pummok_code5.value
				code_ary[5] = document.frm.pummok_code6.value
				code_ary[6] = document.frm.pummok_code7.value
				code_ary[7] = document.frm.pummok_code8.value
				code_ary[8] = document.frm.pummok_code9.value
				code_ary[9] = document.frm.pummok_code10.value
				var popupW = 600;
				var popupH = 400;
				var left = Math.ceil((window.screen.width - popupW)/2);
				var top = Math.ceil((window.screen.height - popupH)/2);
				window.open('sales_goods_select.asp?code_ary='+code_ary+'', '상품선택', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
				code_ary[0] = document.frm.pummok_code1.value
				code_ary[1] = document.frm.pummok_code2.value
				code_ary[2] = document.frm.pummok_code3.value
				code_ary[3] = document.frm.pummok_code4.value
				code_ary[4] = document.frm.pummok_code5.value
				code_ary[5] = document.frm.pummok_code6.value
				code_ary[6] = document.frm.pummok_code7.value
				code_ary[7] = document.frm.pummok_code8.value
				code_ary[8] = document.frm.pummok_code9.value
				code_ary[9] = document.frm.pummok_code10.value

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
				var popupW = 600;
				var popupH = 400;
				var left = Math.ceil((window.screen.width - popupW)/2);
				var top = Math.ceil((window.screen.height - popupH)/2);
				window.open('sales_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'', '선택상품삭제', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();
			var buy_ary = new Array();
			var sales_ary = new Array();
			var buy_tot = new Array();
			var sales_tot = new Array();
			var margin_tot = new Array();

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				sales_ary[j] = eval("document.frm.sales_cost" + j + ".value").replace(/,/g,"");
				buy_tot[j] = qty_ary[j] * buy_ary[j];
				
				tot_cal = qty_ary[j] * sales_ary[j];
				tot_cal = String(tot_cal);
				num_len = tot_cal.length;
				sil_len = num_len;
				if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
				if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
				if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
				eval("document.frm.sales_tot" + j + ".value = tot_cal");
				sales_tot[j] = qty_ary[j] * sales_ary[j];
				
				tot_cal = sales_ary[j] - buy_ary[j];
				tot_cal = String(tot_cal);
				num_len = tot_cal.length;
				sil_len = num_len;
				if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
				if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
				if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
				eval("document.frm.margin_cost" + j + ".value = tot_cal");

				tot_cal = (sales_ary[j] - buy_ary[j]) * qty_ary[j];
				tot_cal = String(tot_cal);
				num_len = tot_cal.length;
				sil_len = num_len;
				if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
				if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
				if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
				eval("document.frm.margin_tot" + j + ".value = tot_cal");
				margin_tot[j] = (sales_ary[j] - buy_ary[j]) * qty_ary[j];
			}

			buy_tot_cost = 0;
			sales_tot_cost = 0;
			margin_tot_cost = 0;
			for (j=1;j<21;j++) {
				buy_tot_cost = buy_tot_cost + buy_tot[j];
				sales_tot_cost = sales_tot_cost + sales_tot[j];
				margin_tot_cost = margin_tot_cost + margin_tot[j];
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

			tot_cal = sales_tot_cost;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.sales_tot_cost.value = tot_cal");

			tot_cal = margin_tot_cost;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.margin_tot_cost.value = tot_cal");

			tot_cal = margin_tot_cost / sales_tot_cost * 100 ;
			eval("document.frm.margin_per.value = roundXL(tot_cal,2)");

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
			function sales_mod_view() 
			{
				if (document.frm.sales_mod_ck.checked == true) {
					document.getElementById('sales_company').style.display = ''; 
					document.getElementById('sales_org_name').style.display = ''; 
					document.getElementById('sales_mod').style.display = ''; }
				if (document.frm.sales_mod_ck.checked == false) {
					document.getElementById('sales_company').style.display = 'none'; 
					document.getElementById('sales_org_name').style.display = 'none'; 
					document.getElementById('sales_mod').style.display = 'none'; }
			}
        </script>
	</head>
	<body>
		<div id="wrap">			
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="sales_slip_wait_add_save.asp" enctype="multipart/form-data">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>전표유형</th>
							  <td class="left">대기전표</td>
							  <th>매출조직</th>
							  <td colspan="5" class="left">
							  <%=emp_company%>-<%=org_name%>&nbsp;
                              <strong>매출조직변경</strong>
								<input name="sales_mod_ck" type="checkbox" id="sales_mod_ck" value="1"  onClick="sales_mod_view()">
                				<input name="sales_company" id="sales_company" type="text" value="<%=emp_company%>" readonly="true" style="display:none;width:150px">
                				<input name="sales_org_name" id="sales_org_name" type="text" value="<%=org_name%>" readonly="true" style="display:none;width:150px">
                				<input name="sales_bonbu" type="hidden" id="sales_bonbu" value="<%=bonbu%>">
                				<input name="sales_saupbu" type="hidden" id="sales_saupbu" value="<%=saupbu%>">
                				<input name="sales_team" type="hidden" id="sales_team" value="<%=team%>">
                                <a href="#" class="btnType03" onClick="pop_Window('org_search.asp?gubun=<%="영업"%>','org_search_pop','scrollbars=yes,width=600,height=400')" id="sales_mod" style="display:none">조직변경</a>
                              </td>
						    </tr>
							<tr>
							  <th>거래처</th>
							  <td class="left">
                              <input name="trade_name" type="text" value="<%=trade_name%>" readonly="true" style="width:120px">
						      <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="1"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a>
                              </td>
							  <th>사업자번호</th>
							  <td class="left"><input name="trade_no" type="text" value="<%=trade_no%>" readonly="true" style="width:150px"></td>
							  <th>거래처<br>
						      담당자</th>
							  <td class="left"><input name="trade_person" type="text" value="<%=trade_person%>" id="trade_person" style="width:150px; ime-mode:active" onKeyUp="checklength(this,20);"></td>
							  <th>계산서 메일</th>
							  <td class="left"><input name="trade_email" type="text" value="<%=trade_email%>" id="trade_email" style="width:180px" onKeyUp="checklength(this,30);"></td>
						    </tr>
							<tr>
							  <th>제품<br>출고방법</th>
							  <td class="left"><input name="out_method" type="text" id="out_method" style="width:150px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=out_method%>"></td>
							  <th>제품출고<br>요청일</th>
							  <td class="left"><input name="out_request_date" type="text" value="<%=out_request_date%>" style="width:80px;text-align:center" id="datepicker"></td>
							  <th>매출구분</th>
							  <td class="left">
                              <input type="radio" name="sales_yn" value="Y" <% if sales_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio3">매출
  							  <input type="radio" name="sales_yn" value="N" <% if sales_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio4">비매출
                              </td>
							  <th>매출일자</th>
							  <td class="left"><input name="sales_date" type="text" value="<%=sales_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
						    </tr>
							<tr>
							  <th>계산서<br>발행예정일</th>
							  <td class="left"><input name="bill_due_date" type="text" value="<%=bill_due_date%>" style="width:80px;text-align:center" id="datepicker2"></td>
							  <th>계산서<br>발행여부</th>
							  <td class="left">
                              <input type="radio" name="bill_issue_yn" value="Y" <% if bill_issue_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio3">발행
  							  <input type="radio" name="bill_issue_yn" value="N" <% if bill_issue_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio4">미발행
                              </td>
							  <th>계산서발행일</th>
							  <td class="left"><input name="bill_issue_date" type="text" value="<%=bill_issue_date%>" style="width:80px;text-align:center" id="datepicker3"></td>
							  <th>수금상태</th>
							  <td class="left">
                              <input type="radio" name="collect_stat" value="청구" <% if collect_stat = "청구" then %>checked<% end if %> style="width:30px" id="Radio3">청구
  							  <input type="radio" name="collect_stat" value="영수" <% if collect_stat = "영수" then %>checked<% end if %> style="width:30px" id="Radio4">영수
                              </td>
						    </tr>
							<tr>
							  <th>수금방법</th>
							  <td colspan="3" class="left">
                              <input type="radio" name="bill_collect" value="현금" <% if bill_collect = "현금" then %>checked<% end if %> style="width:40px" id="Radio3">현금
  							  <input type="radio" name="bill_collect" value="어음" <% if bill_collect = "어음" then %>checked<% end if %> style="width:40px" id="Radio4">어음
                              <input type="radio" name="bill_collect" value="카드" <% if bill_collect = "카드" then %>checked<% end if %> style="width:40px" id="Radio3">카드
  							  <input type="radio" name="bill_collect" value="외환" <% if bill_collect = "외환" then %>checked<% end if %> style="width:40px" id="Radio4">외환
                              </td>
							  <th>수금예정일</th>
							  <td class="left"><input name="collect_due_date" type="text" value="<%=collect_due_date%>" style="width:80px;text-align:center" id="datepicker4"></td>
							  <th>수금완료일</th>
							  <td class="left"><input name="collect_date" type="text" value="<%=collect_date%>" style="width:80px;text-align:center" id="datepicker5"></td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="7" class="left"><textarea name="slip_memo" rows="3" id="textarea"><%=slip_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
<h3 class="stit">* 계약 세부 내용</h3>
            		<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="12%" >
							<col width="12%" >
							<col width="*" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="10" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">품목코드선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="tot_check" id="tot_check"></th>
								<th scope="col">서비스유형</th>
								<th scope="col">품목</th>
								<th scope="col">규격</th>
								<th scope="col">수량</th>
								<th scope="col">구입단가</th>
								<th scope="col">판매단가</th>
								<th scope="col">판매총액</th>
								<th scope="col">마진단가</th>
								<th scope="col">마진총액</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" style="width:120px" readonly="true">
                                <input name="pummok_code<%=i%>" type="hidden" id="pummok_code<%=i%>" readonly="true">
                                </td>
								<td><input name="pummok<%=i%>" type="text" id="pummok<%=i%>" style="width:120px" readonly="true"></td>
								<td><input name="standard<%=i%>" type="text" id="standard<%=i%>" style="width:170px"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:50px;text-align:right" value="<%=formatnumber(qty,0)%>" onKeyUp="NumCal(this);"></td>
								<td><input name="buy_cost<%=i%>" type="text" id="buy_cost<%=i%>" style="width:100px;text-align:right" value="<%=formatnumber(buy_cost,0)%>" onKeyUp="NumCal(this);" ></td>
								<td><input name="sales_cost<%=i%>" type="text" id="sales_cost<%=i%>" style="width:100px;text-align:right" value="<%=formatnumber(sales_cost,0)%>" onKeyUp="NumCal(this);" ></td>
								<td><input name="sales_tot<%=i%>" type="text" id="sales_tot<%=i%>" style="width:100px;text-align:right" readonly="true" value="<%=formatnumber(sales_tot,0)%>"></td>
								<td><input name="margin_cost<%=i%>" type="text" id="margin_cost<%=i%>" style="width:100px;text-align:right" value="<%=formatnumber(margin_cost,0)%>" readonly="true"></td>
								<td><input name="margin_tot<%=i%>" type="text" id="margin_tot<%=i%>" style="width:100px;text-align:right" value="<%=formatnumber(margin_tot,0)%>" readonly="true"></td>
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
							  <th>매입총액</th>
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_price,0)%>" readonly="true"></td>
							  <th>매입금액</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost,0)%>" readonly="true"></td>
							  <th>매입부가세</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost_vat,0)%>" readonly="true"></td>
						    </tr>
							<tr>
							  <th>매출총액</th>
							  <td class="left"><input name="sales_tot_price" type="text" id="sales_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(sales_tot_price,0)%>" readonly="true"></td>
							  <th>매출금액</th>
							  <td class="left"><input name="sales_tot_cost" type="text" id="sales_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(sales_tot_cost,0)%>" readonly="true"></td>
							  <th>매출부가세</th>
							  <td class="left"><input name="sales_cost_vat" type="text" id="sales_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(sales_cost_vat,0)%>" readonly="true"></td>
						    </tr>
							<tr>
							  <th>마진총액</th>
							  <td class="left"><input name="margin_tot_price" type="text" id="margin_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(margin_tot_price,0)%>" readonly="true"></td>
							  <th>마진금액</th>
							  <td class="left"><input name="margin_tot_cost" type="text" id="margin_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(margin_tot_cost,0)%>" readonly="true"></td>
							  <th>마진비율</th>
							  <td class="left"><input name="margin_per" type="text" id="margin_per" style="width:150px;text-align:right" value="<%=formatnumber(margin_per,2)%>" readonly="true">%</td>
	                      </tr>
							<tr>
							  <th>첨부</th>
							  <td colspan="5" class="left">
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>
<br>
                   		<div align=center>
                            <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();"></span>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    	</div>
					<br>
				</form>
                </div>
			</div>
		</div>
	</body>
</html>

