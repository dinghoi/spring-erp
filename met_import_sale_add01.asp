<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(7,20)
dim amount_tab(3,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")

emp_company = request.cookies("nkpmg_user")("coo_emp_company")
bonbu = request.cookies("nkpmg_user")("coo_bonbu")
saupbu = request.cookies("nkpmg_user")("coo_saupbu")
team = request.cookies("nkpmg_user")("coo_team")
org_name = request.cookies("nkpmg_user")("coo_org_name")

u_type = request("u_type")

view_condi=Request("view_condi")
goods_type=Request("goods_type")
chulgo_id=Request("chulgo_id")

curr_date = mid(cstr(now()),1,10)
chulgo_date = curr_date
chulgo_goods_type = goods_type

chulgo_seq = ""
service_no = ""
chulgo_trade_name = ""
chulgo_trade_dept = ""
chulgo_type = ""
chulgo_stock_company = ""
chulgo_stock_name = ""
chulgo_emp_no = ""
chulgo_emp_name = ""
chulgo_company = ""
chulgo_bonbu = ""
chulgo_saupbu = ""
chulgo_team = ""
chulgo_org_name = ""
chulgo_memo = ""

mok_cnt = 0
pummok_cnt = 0

path_name = "/met_upload"

for i = 1 to 7
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
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&user_id&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
    	emp_no = rs_emp("emp_no")
		emp_name = rs_emp("emp_name")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
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

chulgo_emp_no = emp_no
chulgo_emp_name = emp_name
chulgo_company = emp_company
chulgo_bonbu = emp_bonbu
chulgo_saupbu = emp_saupbu
chulgo_team = emp_team
chulgo_org_name = emp_org_name

'chulgo_stock = emp_no
'chulgo_stock_company = emp_company
'chulgo_stock_name = emp_name

title_line = " N/W 고객사 출고등록 "

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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=chulgo_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=chulgo_date%>" );
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
				if(document.frm.chulgo_stock_name.value == "") {
					alert('출고창고를 선택하세요');
					frm.chulgo_stock_name.focus();
					return false;}
				if(document.frm.chulgo_goods_type.value == "") {
					alert('용도구분을 선택하세요');
					frm.chulgo_goods_type.focus();
					return false;}
				if(document.frm.service_no.value == "") {
					alert('전표번호를 선택하세요');
					frm.service_no.focus();
					return false;}
				if(document.frm.chulgo_trade_name.value == "") {
					alert('고객사를 선택하세요');
					frm.chulgo_trade_name.focus();
					return false;}
				if(document.frm.chulgo_date.value == "") {
					alert('출고일자를 입력하세요');
					frm.chulgo_date.focus();
					return false;}

										
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
				code_ary[10] = document.frm.goods_code11.value + "/" + document.frm.goods_standard11.value;
				code_ary[11] = document.frm.goods_code12.value + "/" + document.frm.goods_standard12.value;
				code_ary[12] = document.frm.goods_code13.value + "/" + document.frm.goods_standard13.value;
				code_ary[13] = document.frm.goods_code14.value + "/" + document.frm.goods_standard14.value;
				code_ary[14] = document.frm.goods_code15.value + "/" + document.frm.goods_standard15.value;
				code_ary[15] = document.frm.goods_code16.value + "/" + document.frm.goods_standard16.value;
				code_ary[16] = document.frm.goods_code17.value + "/" + document.frm.goods_standard17.value;
				code_ary[17] = document.frm.goods_code18.value + "/" + document.frm.goods_standard18.value;
				code_ary[18] = document.frm.goods_code19.value + "/" + document.frm.goods_standard19.value;
				code_ary[19] = document.frm.goods_code20.value + "/" + document.frm.goods_standard20.value;
				goods_type = document.frm.chulgo_goods_type.value
				stock_code = document.frm.chulgo_stock.value
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_goods_select.asp?code_ary='+code_ary+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
				
				var url = "met_stock_goods_select.asp?code_ary="+code_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;
//				var url = "met_goods_select.asp?code_ary="+code_ary;
				pop_Window(url,'출고품목선택','scrollbars=yes,width=960,height=600');
				
			} 
			function pop_pummok_del() 
			{ 
				var code_ary = new Array();
				var del_ary = new Array();
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
				code_ary[10] = document.frm.goods_code11.value
				code_ary[11] = document.frm.goods_code12.value
				code_ary[12] = document.frm.goods_code13.value
				code_ary[13] = document.frm.goods_code14.value
				code_ary[14] = document.frm.goods_code15.value
				code_ary[15] = document.frm.goods_code16.value
				code_ary[16] = document.frm.goods_code17.value
				code_ary[17] = document.frm.goods_code18.value
				code_ary[18] = document.frm.goods_code19.value
				code_ary[19] = document.frm.goods_code20.value

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
					
				if (document.frm.del_check11.checked == true) {
					del_ary[10] = 'Y'; } 					
				if (document.frm.del_check11.checked == false) {
					del_ary[10] = 'N'; } 
				if (document.frm.del_check12.checked == true) {
					del_ary[11] = 'Y'; } 					
				if (document.frm.del_check12.checked == false) {
					del_ary[11] = 'N'; } 
				if (document.frm.del_check13.checked == true) {
					del_ary[12] = 'Y'; } 					
				if (document.frm.del_check13.checked == false) {
					del_ary[12] = 'N'; } 
				if (document.frm.del_check14.checked == true) {
					del_ary[13] = 'Y'; } 					
				if (document.frm.del_check14.checked == false) {
					del_ary[13] = 'N'; } 
				if (document.frm.del_check15.checked == true) {
					del_ary[14] = 'Y'; } 					
				if (document.frm.del_check15.checked == false) {
					del_ary[14] = 'N'; } 
				if (document.frm.del_check16.checked == true) {
					del_ary[15] = 'Y'; } 					
				if (document.frm.del_check16.checked == false) {
					del_ary[15] = 'N'; } 
				if (document.frm.del_check17.checked == true) {
					del_ary[16] = 'Y'; } 					
				if (document.frm.del_check17.checked == false) {
					del_ary[16] = 'N'; } 
				if (document.frm.del_check18.checked == true) {
					del_ary[17] = 'Y'; } 					
				if (document.frm.del_check18.checked == false) {
					del_ary[17] = 'N'; } 
				if (document.frm.del_check19.checked == true) {
					del_ary[18] = 'Y'; } 					
				if (document.frm.del_check19.checked == false) {
					del_ary[18] = 'N'; } 
				if (document.frm.del_check20.checked == true) {
					del_ary[19] = 'Y'; } 					
				if (document.frm.del_check20.checked == false) {
					del_ary[19] = 'N'; } 
				goods_type = document.frm.chulgo_goods_type.value
				stock_code = document.frm.chulgo_stock.value	
//				var popupW = 600;
//				var popupH = 400;
//				var left = Math.ceil((window.screen.width - popupW)/2);
//				var top = Math.ceil((window.screen.height - popupH)/2);
//				window.open('met_stock_goods_del_ok.asp?code_ary='+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code+'', '팝업공지', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
				
				var url = "met_stock_goods_del_ok.asp?code_ary="+code_ary+'&del_ary='+del_ary+'&goods_type='+goods_type+'&stock_code='+stock_code;				
				pop_Window(url,'선택상품삭제','scrollbars=yes,width=600,height=400');
				alert("삭제되었습니다 !!!");
				NumCal();
			} 

		function NumCal(txtObj){
			var qty_ary = new Array();
			var amt_ary = new Array();
			var b_qty_ary = new Array();

            buy_tot_cost = 0;
			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
				amt_ary[j] = eval("document.frm.c_amt" + j + ".value").replace(/,/g,"");
				b_qty_ary[j] = eval("document.frm.jqty" + j + ".value").replace(/,/g,"");
				buy_tot_cost = buy_tot_cost + parseInt(amt_ary[j]);
				
				acpt_qty = parseInt(qty_ary[j]);
				sign_qty = parseInt(b_qty_ary[j]);

	            if (acpt_qty > sign_qty) {
					alert ("재고수량보다 출고수량이 많습니다!!");
					return false;
				}
		
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
			
			if (txtObj.value.length<5) {
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
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_import_sale_add01_save.asp" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="7%" >
							<col width="12%" >
							<col width="7%" >
							<col width="12%" >
							<col width="7%" >
							<col width="18%" >
							<col width="7%" >
							<col width="30%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>출고회사</th>
							  <td class="left">
							<%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="chulgo_company" id="chulgo_company" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = chulgo_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>사업부</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="chulgo_saupbu" id="chulgo_saupbu" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = chulgo_saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
                              <th>출고담당자</th>
							  <td class="left"><%=chulgo_emp_name%>(<%=chulgo_emp_no%>)&nbsp;-&nbsp;<%=chulgo_org_name%>
                              <input type="hidden" name="chulgo_emp_no" value="<%=chulgo_emp_no%>" ID="chulgo_emp_no">
                              <input type="hidden" name="chulgo_emp_name" value="<%=chulgo_emp_name%>" ID="chulgo_emp_name">
                              <input type="hidden" name="chulgo_bonbu" value="<%=chulgo_bonbu%>" ID="chulgo_bonbu">
                              <input type="hidden" name="chulgo_team" value="<%=chulgo_team%>" ID="chulgo_team">
                              <input type="hidden" name="chulgo_org_name" value="<%=chulgo_org_name%>" ID="chulgo_org_name">
                              </td>
                              <th>출고창고</th>
							  <td  class="left">
                              <input name="chulgo_stock_company" type="text" value="<%=chulgo_stock_company%>" readonly="true" style="width:120px">
                              <input name="chulgo_stock_name" type="text" value="<%=chulgo_stock_name%>" readonly="true" style="width:120px">
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="chulgo"%>&view_condi=<%=view_condi%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>" ID="Hidden1">
                              <input type="hidden" name="stock_bonbu" value="<%=stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_saupbu" value="<%=stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="stock_team" value="<%=stock_team%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                              </td>
						    </tr>
							<tr>
							  <th>출고일자</th>
							  <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                              <th>용도구분</th>
							  <td class="left">
							<%
                                Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
                            %>
                                <select name="chulgo_goods_type" id="chulgo_goods_type" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until Rs_etc.eof 
                            %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If chulgo_goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                            <%
                                    Rs_etc.movenext()  
                                loop 
                                Rs_etc.Close()
                            %>
                                </select>
                              </td>
                              <th>전표번호</th>
							  <td class="left">
							  <input name="service_no" type="text" id="service_no" style="width:180px" value="<%=service_no%>"></td>
							  <th>고객사/지점</th>
							  <td class="left">
                              <input name="chulgo_trade_name" type="text" value="<%=chulgo_trade_name%>" readonly="true" style="width:120px">
                              
                              <input name="chulgo_trade_dept" type="text" value="<%=chulgo_trade_dept%>" style="width:120px; text-align:left; ime-mode:active"">
						      <a href="#" class="btnType03" onClick="pop_Window('insa_trade_select.asp?gubun=<%="chulgo"%>','trade_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              </td>
						    </tr>
                            <tr>
							  <th>비고</th>
							  <td class="left" colspan="8" ><textarea name="chulgo_memo" rows="3" style="text-align:left; ime-mode:active" id="textarea"><%=chulgo_memo%></textarea></td>
						    </tr>
						</tbody>
					</table>
				</div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="10%" >
							<col width="10%" >
							<col width="*" >
                            <col width="12%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="10%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="10" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">출고품목선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="tot_check" id="tot_check"></th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격/Part_No</th>
                                <th scope="col">재고수량</th>
								<th scope="col">출고수량</th>
                                <th scope="col">출고금액</th>
                                <th scope="col">Serial_No</th>
							</tr>
						</thead>
						<tbody>
						<%
							for i = 1 to 20
						%>
			  				<tr id="pummok_list<%=i%>" style="display:none">
								<td class="first"><input type="checkbox" name="del_check<%=i%>" id="del_check<%=i%>" value="Y"></td>
								<td>
                                <input name="srv_type<%=i%>" type="text" id="srv_type<%=i%>" value="<%=pummok_tab(1,i)%>" style="width:90px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:110px" readonly="true">
                                </td>
                                <td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:100px" readonly="true">
                                </td>
								<td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:230px" readonly="true"></td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:130px"></td>
                                <td><input name="jqty<%=i%>" type="text" id="jqty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(jqty,0)%>" readonly="true"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(amount_tab(1,i),0)%>" onKeyUp="NumCal(this);">
                                <input type="hidden" name="goods_grade<%=i%>" value="<%=pummok_tab(6,i)%>" ID="Hidden1">
                                </td>
                                <td><input name="c_amt<%=i%>" type="text" id="c_amt<%=i%>" style="width:100px;text-align:right" value="<%=formatnumber(amount_tab(2,i),0)%>" onKeyUp="NumCal(this);">
                                </td>
                                <td class="left">
                                <input name="excel_att_file<%=i%>" type="file" id="excel_att_file<%=i%>" size="2">
                                </td>
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
							  <th>출고총액</th>
							  <td class="left"><input name="buy_tot_price" type="text" id="buy_tot_price" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_price,0)%>" readonly="true"></td>
							  <th>출고금액</th>
							  <td class="left"><input name="buy_tot_cost" type="text" id="buy_tot_cost" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost,0)%>" readonly="true"></td>
							  <th>출고부가세</th>
							  <td class="left"><input name="buy_tot_cost_vat" type="text" id="buy_tot_cost_vat" style="width:150px;text-align:right" value="<%=formatnumber(buy_tot_cost_vat,0)%>" readonly="true"></td>
						    </tr>
						</tbody>
					</table>
<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="old_chulgo_date" value="<%=chulgo_date%>">
                <input type="hidden" name="old_chulgo_stock" value="<%=chulgo_stock%>">
				<input type="hidden" name="old_chulgo_seq" value="<%=chulgo_seq%>">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
				</form>
		</div>				
	</body>
</html>
