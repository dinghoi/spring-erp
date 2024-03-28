<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	dim pummok_tab(4,20)
	dim cost_tab(9,40)
	
	u_type = request("u_type")
	slip_id = request("slip_id")
	slip_no = request("slip_no")
	slip_seq = request("slip_seq")
	
	for i = 1 to 4
		for j = 1 to 20
			pummok_tab(i,j) = ""
		next
	next
	for i = 1 to 9
		for j = 1 to 20
			cost_tab(i,j) = 0
		next
	next
	
	Sql="select * from sales_slip where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)
	
	sales_company = rs("sales_company")
	sales_bonbu = rs("sales_bonbu")
	sales_saupbu = rs("sales_saupbu")
	sales_team = rs("sales_team")
	sales_org_name = rs("sales_org_name")
	trade_code = rs("trade_code")
'	trade_no = rs("trade_no")
	trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + right(rs("trade_no"),5)
	trade_name = rs("trade_name")
	trade_person = rs("trade_person")
	trade_email = rs("trade_email")
	trade_person_tel_no = rs("trade_person_tel_no")
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
	view_att_file = rs("att_file")
	view_slip_no = slip_no + "-" + slip_seq
	sign_yn = rs("sign_yn")
	rs.close()

	j = 0
	Sql="select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
'		if rs("srv_type") <> "상품" then
			j = j + 1
			pummok_tab(1,j) = rs("goods_code")
			pummok_tab(2,j) = rs("srv_type")
			pummok_tab(3,j) = rs("pummok")
			pummok_tab(4,j) = rs("standard")
			cost_tab(1,j) = rs("qty")
			cost_tab(2,j) = rs("buy_cost")
			cost_tab(3,j) = rs("sales_cost")
			cost_tab(4,j) = rs("qty") * rs("sales_cost")
			cost_tab(5,j) = rs("margin_cost")
			cost_tab(6,j) = rs("qty") * rs("margin_cost")
			cost_tab(7,j) = rs("order_qty")
			cost_tab(8,j) = int(rs("qty")) - int(rs("order_qty"))
			cost_tab(9,j) = rs("sales_cost")
'		end if		
		rs.movenext()
	loop
	pummok_cnt = j

	title_line = "대기 전표 -> 수주 전표 등록"
	view_slip_id = "수주전표"
	
	path_name = "/sales_file"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
				code_ary[0] = document.frm.pummok_code1.value + "/" + document.frm.standard1.value;
				code_ary[1] = document.frm.pummok_code2.value + "/" + document.frm.standard2.value;
				code_ary[2] = document.frm.pummok_code3.value + "/" + document.frm.standard3.value;
				code_ary[3] = document.frm.pummok_code4.value + "/" + document.frm.standard4.value;
				code_ary[4] = document.frm.pummok_code5.value + "/" + document.frm.standard5.value;
				code_ary[5] = document.frm.pummok_code6.value + "/" + document.frm.standard6.value;
				code_ary[6] = document.frm.pummok_code7.value + "/" + document.frm.standard7.value;
				code_ary[7] = document.frm.pummok_code8.value + "/" + document.frm.standard8.value;
				code_ary[8] = document.frm.pummok_code9.value + "/" + document.frm.standard9.value;
				code_ary[9] = document.frm.pummok_code10.value + "/" + document.frm.standard10.value;
				slip_id = document.frm.slip_id.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('sales_goods_select.asp?code_ary='+code_ary+'', '상품선택', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');

				var url = "sales_goods_select.asp?code_ary="+code_ary+'&slip_id='+slip_id;				
				pop_Window(url,'상품선택','scrollbars=yes,width=600,height=400');

		
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
//				alert("aaaa");
				code_ary[0] = document.frm.pummok_code1.value + "/" + document.frm.standard1.value + "/" + document.frm.order_qty1.value + "/" + document.frm.buy_cost1.value + "/" + document.frm.order_cost1.value;
				code_ary[1] = document.frm.pummok_code2.value + "/" + document.frm.standard2.value + "/" + document.frm.order_qty2.value + "/" + document.frm.buy_cost2.value + "/" + document.frm.order_cost2.value;
				code_ary[2] = document.frm.pummok_code3.value + "/" + document.frm.standard3.value + "/" + document.frm.order_qty3.value + "/" + document.frm.buy_cost3.value + "/" + document.frm.order_cost3.value;
				code_ary[3] = document.frm.pummok_code4.value + "/" + document.frm.standard4.value + "/" + document.frm.order_qty4.value + "/" + document.frm.buy_cost4.value + "/" + document.frm.order_cost4.value;
				code_ary[4] = document.frm.pummok_code5.value + "/" + document.frm.standard5.value + "/" + document.frm.order_qty5.value + "/" + document.frm.buy_cost5.value + "/" + document.frm.order_cost5.value;
				code_ary[5] = document.frm.pummok_code6.value + "/" + document.frm.standard6.value + "/" + document.frm.order_qty6.value + "/" + document.frm.buy_cost6.value + "/" + document.frm.order_cost6.value;
				code_ary[6] = document.frm.pummok_code7.value + "/" + document.frm.standard7.value + "/" + document.frm.order_qty7.value + "/" + document.frm.buy_cost7.value + "/" + document.frm.order_cost7.value;
				code_ary[7] = document.frm.pummok_code8.value + "/" + document.frm.standard8.value + "/" + document.frm.order_qty8.value + "/" + document.frm.buy_cost8.value + "/" + document.frm.order_cost8.value;
				code_ary[8] = document.frm.pummok_code9.value + "/" + document.frm.standard9.value + "/" + document.frm.order_qty9.value + "/" + document.frm.buy_cost9.value + "/" + document.frm.order_cost9.value;
				code_ary[9] = document.frm.pummok_code10.value + "/" + document.frm.standard10.value + "/" + document.frm.order_qty10.value + "/" + document.frm.buy_cost10.value + "/" + document.frm.order_cost10.value;

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
//				window.open('sales_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'', '선택상품삭제', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				var url = "sales_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary;				
				pop_Window(url,'선택상품삭제','scrollbars=yes,width=600,height=400');
				alert("삭제되었습니다 !!!");
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
				qty_ary[j] = eval("document.frm.order_qty" + j + ".value").replace(/,/g,"");
				buy_ary[j] = eval("document.frm.buy_cost" + j + ".value").replace(/,/g,"");
				sales_ary[j] = eval("document.frm.order_cost" + j + ".value").replace(/,/g,"");
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

			tot_cal = sales_tot_cost * 0.1 ;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.sales_tot_cost_vat.value = tot_cal");

			tot_cal = sales_tot_cost + sales_tot_cost * 0.1 ;
			tot_cal = String(tot_cal);
			num_len = tot_cal.length;
			sil_len = num_len;
			if (tot_cal.substr(0,1) == "-") sil_len = num_len - 1;
			if (sil_len > 3) tot_cal = tot_cal.substr(0,num_len -3) + "," + tot_cal.substr(num_len -3,3);
			if (sil_len > 6) tot_cal = tot_cal.substr(0,num_len -6) + "," + tot_cal.substr(num_len -6,3) + "," + tot_cal.substr(num_len -2,3);
			if (sil_len > 9) tot_cal = tot_cal.substr(0,num_len -9) + "," + tot_cal.substr(num_len -9,3) + "," + tot_cal.substr(num_len -5,3) + "," + tot_cal.substr(num_len -1,3);
			eval("document.frm.sales_tot_price.value = tot_cal");

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
			function pummok_list_view() {
				pummok_cnt = parseInt(document.frm.pummok_cnt.value);
				for (j=1;j<pummok_cnt+1;j++) {
					eval("document.getElementById('pummok_list" + j + "')").style.display = '';
				}
				NumCal();
			}
			function no_sales_view() 
			{
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.sales_yn[" + j + "].checked")) {
						k = j
					}
				}
				if (k == 0) {
					document.getElementById('no_sales1').style.display = ''; 
					document.getElementById('no_sales2').style.display = ''; 
					document.getElementById('sales_date').disabled = ''; }
				if (k == 1) {
					document.getElementById('no_sales1').style.display = 'none'; 
					document.getElementById('no_sales2').style.display = 'none'; 
					document.getElementById('sales_date').disabled = 'false'; }
			}
			function bill_issue_view() 
			{
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.bill_issue_yn[" + j + "].checked")) {
						k = j
					}
				}
				if (k == 0) {
					document.getElementById('bill_due_date').disabled = ''; 
					document.getElementById('bill_issue_date').disabled = ''; }
				if (k == 1) {
					document.getElementById('bill_due_date').disabled = 'false'; 
					document.getElementById('bill_issue_date').disabled = 'false'; }
			}
			function collect_view() 
			{
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.collect_stat[" + j + "].checked")) {
						k = j
					}
				}
				if (k == 0) {
					document.getElementById('collect_due_date').disabled = ''; 
					document.getElementById('collect_date').disabled = 'false'; }
				if (k == 1) {
					document.getElementById('collect_due_date').disabled = 'false'; 
					document.getElementById('collect_date').disabled = ''; }
			}
			function trade_person_view () {
				if(document.frm.trade_code.value =="") {
					alert('거래처를 검색하세요');
					return false;}
				var trade_code = document.frm.trade_code.value;
				var url = "trade_person_search.asp?trade_code="+trade_code;				
				pop_Window(url,'trade_person_search_pop','scrollbars=yes,width=600,height=400');
			}			
        </script>
	</head>
	<body onload="pummok_list_view();">
		<div id="wrap">			
	<% if u_type <> "U" then	%>
		<!--#include virtual = "/include/sales_header.asp" -->
		<!--#include virtual = "/include/sales_menu.asp" -->
	<% end if	%>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="sales_slip_exchange_ok.asp" enctype="multipart/form-data">
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
							  <td class="left"><%=view_slip_id%>&nbsp;<%=view_slip_no%></td>
							  <th>매출조직</th>
							  <td colspan="5" class="left">
							  <%=sales_company%>-<%=sales_org_name%>&nbsp;&nbsp;
                              <strong>매출조직변경</strong>
								<input name="sales_mod_ck" type="checkbox" id="sales_mod_ck" value="1"  onClick="sales_mod_view()">
                				<input name="sales_company" id="sales_company" type="text" value="<%=sales_company%>" readonly="true" style="display:none;width:150px">
                				<input name="sales_org_name" id="sales_org_name" type="text" value="<%=sales_org_name%>" readonly="true" style="display:none;width:150px">
                				<input name="sales_bonbu" type="hidden" id="sales_bonbu" value="<%=sales_bonbu%>">
                				<input name="sales_saupbu" type="hidden" id="sales_saupbu" value="<%=sales_saupbu%>">
                				<input name="sales_team" type="hidden" id="sales_team" value="<%=sales_team%>">
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
							  <td class="left">
                              <input name="trade_person" type="text" value="<%=trade_person%>" id="trade_person" style="width:120px; ime-mode:active" onKeyUp="checklength(this,20);">
						      <a href="#" onClick="trade_person_view();" class="btnType03">조회</a>
                              </td>
							  <th>계산서 메일</th>
							  <td class="left"><input name="trade_email" type="text" value="<%=trade_email%>" id="trade_email" style="width:180px" onKeyUp="checklength(this,30);"></td>
						    </tr>
							<tr>
							  <th>담당자<br>연락처</th>
							  <td class="left"><input name="trade_person_tel_no" type="text" id="trade_person_tel_no" style="width:150px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=trade_person_tel_no%>"></td>
							  <th>제품출고<br>요청일</th>
							  <td class="left"><input name="out_request_date" type="text" value="<%=out_request_date%>" style="width:80px;text-align:center" id="datepicker"></td>
							  <th>매출구분</th>
							  <td class="left">
                              <input type="radio" name="sales_yn" value="Y" <% if sales_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="no_sales_view();">매출
  							  <input type="radio" name="sales_yn" value="N" <% if sales_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="no_sales_view();">비매출
                              </td>
							  <th>매출일자</th>
							  <td class="left"><input name="sales_date" type="text" value="<%=sales_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
						    </tr>
							<tr style="display:" id="no_sales1">
							  <th>계산서<br>
							  발행여부</th>
							  <td class="left"><input type="radio" name="bill_issue_yn" value="Y" <% if bill_issue_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio" onClick="bill_issue_view();">
							    발행
                                  <input type="radio" name="bill_issue_yn" value="N" <% if bill_issue_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio2" onClick="bill_issue_view();">
                              미발행</td>
							  <th>계산서<br>
							    발행예정일</th>
							  <td class="left"><input name="bill_due_date" type="text" id="datepicker2" style="width:80px;text-align:center" value="<%=bill_due_date%>"></td>
							  <th>계산서발행일</th>
							  <td class="left"><input name="bill_issue_date" type="text" value="<%=bill_issue_date%>" style="width:80px;text-align:center" id="datepicker3"></td>
							  <th>수금상태</th>
							  <td class="left">
                              <input type="radio" name="collect_stat" value="청구" <% if collect_stat = "청구" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="collect_view();">청구
  							  <input type="radio" name="collect_stat" value="영수" <% if collect_stat = "영수" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="collect_view();">영수
                              </td>
						    </tr>
							<tr style="display:" id="no_sales2">
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
							  <td class="left"><input name="collect_date" type="text" disabled id="datepicker5" style="width:80px;text-align:center" value="<%=collect_date%>"></td>
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
							<col width="7%" >
							<col width="10%" >
							<col width="*" >
							<col width="4%" >
							<col width="8%" >
							<col width="8%" >
							<col width="4%" >
							<col width="4%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">순번</th>
								<th colspan="6" scope="col" style=" border-bottom:1px solid #e3e3e3;">대기전표 내역</th>
								<th rowspan="2">기존<br>수주<br>수량</th>
								<th colspan="5" style=" border-bottom:1px solid #e3e3e3;" scope="col">수주전표 내역</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">서비스유형</th>
							  <th scope="col">품목</th>
							  <th scope="col">규격</th>
							  <th scope="col">수량</th>
							  <th scope="col">구입단가</th>
							  <th scope="col">판매단가</th>
							  <th scope="col">수주<br>수량</th>
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
								<td class="first"><%=i%></td>
								<td>
								<%=pummok_tab(2,i)%>
                                <input name="pummok_code<%=i%>" type="hidden" id="pummok_code<%=i%>" value="<%=pummok_tab(1,i)%>">
                                <input name="srv_type<%=i%>" type="hidden" id="srv_type<%=i%>" value="<%=pummok_tab(2,i)%>">
                                </td>
								<td><%=pummok_tab(3,i)%><input name="pummok<%=i%>" type="hidden" id="pummok<%=i%>" value="<%=pummok_tab(3,i)%>"></td>
								<td><%=pummok_tab(4,i)%>&nbsp;<input name="standard<%=i%>" type="hidden" id="standard<%=i%>" value="<%=pummok_tab(4,i)%>"></td>
								<td><%=formatnumber(cost_tab(1,i),0)%><input name="qty<%=i%>" type="hidden" id="qty<%=i%>" value="<%=cost_tab(1,i)%>"></td>
								<td class="right"><%=formatnumber(cost_tab(2,i),0)%><input name="buy_cost<%=i%>" type="hidden" id="buy_cost<%=i%>" value="<%=cost_tab(2,i)%>"></td>
								<td><%=formatnumber(cost_tab(3,i),0)%><input name="sales_cost<%=i%>" type="hidden" id="sales_cost<%=i%>" value="<%=cost_tab(3,i)%>"></td>
								<td><%=formatnumber(cost_tab(7,i),0)%></td>
								<td><input name="order_qty<%=i%>" type="text" id="order_qty<%=i%>" value="<%=formatnumber(cost_tab(8,i),0)%>" style="width:38px;text-align:right" onKeyUp="NumCal(this);"></td>
								<td><input name="order_cost<%=i%>" type="text" id="order_cost<%=i%>" value="<%=formatnumber(cost_tab(9,i),0)%>" style="width:80px;text-align:right" onKeyUp="NumCal(this);" ></td>
								<td><input name="sales_tot<%=i%>" type="text" disabled id="sales_tot<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(cost_tab(4,i),0)%>"></td>
								<td><input name="margin_cost<%=i%>" type="text" disabled id="margin_cost<%=i%>" style="width:80px;text-align:right"></td>
								<td><input name="margin_tot<%=i%>" type="text" disabled id="margin_tot<%=i%>" style="width:80px;text-align:right"></td>
							</tr>
						<%
							next
						%>
						</tbody>
					</table>                    
<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="9%" >
							<col width="8%" >
							<col width="9%" >
							<col width="8%" >
							<col width="9%" >
							<col width="8%" >
							<col width="9%" >
							<col width="8%" >
							<col width="9%" >
							<col width="7%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>매입총액</th>
							  <td class="right"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:90px;text-align:right" value="<%=formatnumber(buy_tot_cost,0)%>" readonly="true"></td>
							  <th>매출총액</th>
							  <td class="right"><input name="sales_tot_cost" type="text" id="sales_tot_cost" style="width:90px;text-align:right" value="<%=formatnumber(sales_tot_cost,0)%>" readonly="true"></td>
							  <th>매출부가세</th>
							  <td class="right"><input name="sales_tot_cost_vat" type="text" id="sales_tot_cost_vat" style="width:90px;text-align:right" value="<%=formatnumber(sales_tot_cost_vat,0)%>" readonly="true"></td>
							  <th>총매출액</th>
							  <td class="right"><input name="sales_tot_price" type="text" id="sales_tot_price" style="width:90px;text-align:right" value="<%=formatnumber(sales_tot_price,0)%>" readonly="true"></td>
							  <th>마진총액</th>
							  <td class="right"><input name="margin_tot_cost" type="text" id="margin_tot_cost" style="width:90px;text-align:right" value="<%=formatnumber(margin_tot_cost,0)%>" readonly="true"></td>
							  <th>마진비율</th>
							  <td class="right"><input name="margin_per" type="text" id="margin_per" style="padding-right:5px;width:60px;text-align:right" value="<%=formatnumber(margin_per,2)%>" readonly="true"><strong>%</strong></td>
							</tr>
							<tr>
							  <th>첨부</th>
							  <td colspan="11" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=view_att_file%>"><%=view_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="70">
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
					<input type="hidden" name="trade_code" value="<%=trade_code%>">
					<input type="hidden" name="sign_yn" value="<%=sign_yn%>">
					<input type="hidden" name="slip_id" value="<%=slip_id%>">
					<input type="hidden" name="slip_no" value="<%=slip_no%>">
					<input type="hidden" name="slip_seq" value="<%=slip_seq%>">
					<input type="hidden" name="old_att_file" value="<%=view_att_file%>">
					<input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
					<input type="hidden" name="u_type" value="<%=u_type%>">
				</form>
                </div>
			</div>
		</div>
	</body>
</html>

