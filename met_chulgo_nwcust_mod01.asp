<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim pummok_tab(10,20)
dim amount_tab(3,20)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

u_type = request("u_type")

chulgo_date=Request("chulgo_date")
chulgo_stock=Request("chulgo_stock")
chulgo_seq=Request("chulgo_seq")

curr_date = mid(cstr(now()),1,10)

mok_cnt = 0
pummok_cnt = 0

for i = 1 to 10
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
Set Rs_rele = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if u_type = "U" then
    sql = "select * from met_chulgo where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
    Set rs = DbConn.Execute(SQL)
    	chulgo_goods_type = rs("chulgo_goods_type")
        chulgo_id = rs("chulgo_id")
	    service_no = rs("service_no")
	    chulgo_trade_name = rs("chulgo_trade_name")
	    chulgo_trade_dept = rs("chulgo_trade_dept")
	    chulgo_type = rs("chulgo_type")
	    service_acpt = rs("service_no")
	
        chulgo_stock_company = rs("chulgo_stock_company")
        chulgo_stock_name = rs("chulgo_stock_name")
        chulgo_emp_no = rs("chulgo_emp_no")
        chulgo_emp_name = rs("chulgo_emp_name")
        chulgo_company = rs("chulgo_company")
        chulgo_bonbu = rs("chulgo_bonbu")
        chulgo_saupbu = rs("chulgo_saupbu")
        chulgo_team = rs("chulgo_team")
        chulgo_org_name = rs("chulgo_org_name")
        chulgo_memo = rs("chulgo_memo")
		chulgo_att_file = rs("chulgo_att_file")
		
		rele_no = rs("rele_no")
		rele_seq = rs("rele_seq")
		rele_date = rs("rele_date")
		rele_stock = rs("rele_stock")
		rele_stock_company = rs("rele_stock_company")
		rele_stock_name = rs("rele_stock_name")
		rele_company = rs("rele_company")
		rele_saupbu = rs("rele_saupbu")
		rele_emp_no = rs("rele_emp_no")
		rele_emp_name = rs("rele_emp_name")

    rs.close()
	
	j = 0
	Sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
	Rs.Open Sql, Dbconn, 1
	do until rs.eof
		j = j + 1
		pummok_tab(1,j) = rs("cg_goods_type")
		pummok_tab(2,j) = rs("cg_goods_gubun")
		pummok_tab(3,j) = rs("cg_goods_code")
		pummok_tab(4,j) = rs("cg_goods_name")
		pummok_tab(5,j) = rs("cg_standard")
		pummok_tab(6,j) = rs("cg_goods_grade")
		pummok_tab(7,j) = rs("rl_service_no")
		pummok_tab(8,j) = rs("rl_trade_name")
		pummok_tab(9,j) = rs("rl_trade_dept")
		pummok_tab(10,j) = rs("rl_bigo")

		amount_tab(1,j) = rs("cg_qty")

		rs.movenext()
	loop
	pummok_cnt = j
	mok_cnt = j
    
	title_line = " N/W출고 변경 "

end if
rs.close()

path_name = "/met_upload"

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
				a=confirm('변경하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function chkfrm1() {
				if(document.frm.chulgo_goods_type.value == "") {
					alert('용도구분을 선택하세요');
					frm.chulgo_goods_type.focus();
					return false;}
				if(document.frm.chulgo_stock_name.value == "") {
					alert('출고창고를 선택하세요');
					frm.chulgo_stock_name.focus();
					return false;}
				if(document.frm.rele_company.value == "") {
					alert('출고요청회사를 선택하세요');
					frm.rele_company.focus();
					return false;}
				if(document.frm.rele_saupbu.value == "") {
					alert('출고요청사업부를 선택하세요');
					frm.rele_saupbu.focus();
					return false;}
				if(document.frm.rele_stock_name.value == "") {
					alert('요청창고(입고)를 선택하세요');
					frm.rele_stock_name.focus();
					return false;}
				if(document.frm.chulgo_date.value == "") {
					alert('출고일를 입력하세요');
					frm.chulgo_date.focus();
					return false;}

										
				{
				a=confirm('변경하시겠습니까?')
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
				code_ary[0] = document.frm.goods_code1.value;
				code_ary[1] = document.frm.goods_code2.value;
				code_ary[2] = document.frm.goods_code3.value;
				code_ary[3] = document.frm.goods_code4.value;
				code_ary[4] = document.frm.goods_code5.value;
				code_ary[5] = document.frm.goods_code6.value;
				code_ary[6] = document.frm.goods_code7.value;
				code_ary[7] = document.frm.goods_code8.value;
				code_ary[8] = document.frm.goods_code9.value;
				code_ary[9] = document.frm.goods_code10.value;
				code_ary[10] = document.frm.goods_code11.value;
				code_ary[11] = document.frm.goods_code12.value;
				code_ary[12] = document.frm.goods_code13.value;
				code_ary[13] = document.frm.goods_code14.value;
				code_ary[14] = document.frm.goods_code15.value;
				code_ary[15] = document.frm.goods_code16.value;
				code_ary[16] = document.frm.goods_code17.value;
				code_ary[17] = document.frm.goods_code18.value;
				code_ary[18] = document.frm.goods_code19.value;
				code_ary[19] = document.frm.goods_code20.value;
				goods_type = document.frm.chulgo_goods_type.value;
				stock_code = document.frm.chulgo_stock.value;
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

			for (j=1;j<21;j++) {
				qty_ary[j] = eval("document.frm.qty" + j + ".value").replace(/,/g,"");
		
			}

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
					document.frm.action = "met_chulgo_nwcust_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
			}
		</script>

	</head>
	<body onload="pummok_list_view();">
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
                <form method="post" name="frm" action="met_chulgo_nwcust_mod01_save.asp" enctype="multipart/form-data">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="15%" >
                            <col width="10%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<tr>
							  <tr>
							  <th>출고일자</th>
							  <td class="left"><input name="chulgo_date" type="text" value="<%=chulgo_date%>" style="width:120px;text-align:center" id="datepicker"></td>
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
                              <th>출고창고</th>
							  <td colspan="3" class="left">
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
                              <th>요청창고(입고)</th>
							  <td colspan="3" class="left">
                              <input name="rele_stock_company" type="text" value="<%=rele_stock_company%>" readonly="true" style="width:120px">
                              
                              <input name="rele_stock_name" type="text" value="<%=rele_stock_name%>" readonly="true" style="width:120px">
                              
						      <a href="#" class="btnType03" onClick="pop_Window('meterials_stock_select.asp?gubun=<%="mvreg"%>&view_condi=<%=view_condi%>','rstock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                              <input type="hidden" name="rele_stock" value="<%=rele_stock%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_bonbu" value="<%=rele_stock_bonbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_saupbu" value="<%=rele_stock_saupbu%>" ID="Hidden1">
                              <input type="hidden" name="rele_stock_team" value="<%=rele_stock_team%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_code" value="<%=rele_manager_code%>" ID="Hidden1">
                              <input type="hidden" name="rele_manager_name" value="<%=rele_manager_name%>" ID="Hidden1">
                              </td>
                              <th>요청그룹사</th>
							  <td class="left">
							<%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_company" id="rele_company" value="<%=rele_company%>" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = rele_company  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
                              </td>
							  <th>요청사업부</th>
							  <td class="left">
							<%
                                Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                rs_org.Open Sql, Dbconn, 1
                            %>
                                <select name="rele_saupbu" id="rele_saupbu" value="<%=rele_saupbu%>" style="width:120px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_org.eof 
                            %>
                                    <option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = rele_saupbu  then %>selected<% end if %>><%=rs_org("org_name")%></option>
                            <%
                                    rs_org.movenext()  
                                loop 
                                rs_org.Close()
                            %>
                                </select>
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
							<col width="12%" >
                            <col width="12%" >
							<col width="12%" >
							<col width="*" >
                            <col width="16%" >
							<col width="10%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first; left" colspan="8" scope="col" style=" border-bottom:1px solid #e3e3e3;"><a href="#" onClick="pop_pummok_del()" class="btnType03">선택삭제</a>&nbsp;<a href="#" onClick="pop_pummok()" class="btnType03">출고 품목선택</a></th>
							</tr>
							<tr>
								<th class="first" scope="col">선택</th>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
                                <th scope="col">재고수량</th>
								<th scope="col">출고수량</th>
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
                                <input name="goods_gubun<%=i%>" type="text" id="goods_gubun<%=i%>" value="<%=pummok_tab(2,i)%>" style="width:120px" readonly="true">
                                </td>
								<td>
                                <input name="goods_code<%=i%>" type="text" id="goods_code<%=i%>" value="<%=pummok_tab(3,i)%>" style="width:90px" readonly="true">
                                </td>
                                <td><input name="goods_name<%=i%>" type="text" id="goods_name<%=i%>" value="<%=pummok_tab(4,i)%>" style="width:250px" readonly="true">
                                </td>
								<td><input name="goods_standard<%=i%>" type="text" id="goods_standard<%=i%>" value="<%=pummok_tab(5,i)%>" style="width:160px"></td>
                                <td><input name="jqty<%=i%>" type="text" id="jqty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(jqty,0)%>" readonly="true"></td>
								<td><input name="qty<%=i%>" type="text" id="qty<%=i%>" style="width:80px;text-align:right" value="<%=formatnumber(amount_tab(1,i),0)%>" onKeyUp="NumCal(this);">
                                <input type="hidden" name="goods_grade<%=i%>" value="<%=pummok_tab(6,i)%>" ID="Hidden1">
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
							  <th>첨부</th>
                              <td colspan="5" class="left">
                              <a href="download.asp?path=<%=path_name%>&att_file=<%=chulgo_att_file%>"><%=chulgo_att_file%></a>
                              <input name="att_file" type="file" id="att_file" size="100">
                              </td>
						    </tr>
						</tbody>
					</table>                    
					<br>
				</div>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"></span>
			<% end if	%>                          
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                
                <input type="hidden" name="mok_cnt" value="<%=mok_cnt%>">
                <input type="hidden" name="pummok_cnt" value="<%=pummok_cnt%>">
                <input type="hidden" name="chulgo_id" value="<%=chulgo_id%>">
                
                <input type="hidden" name="old_chulgo_date" value="<%=chulgo_date%>">
                <input type="hidden" name="old_chulgo_stock" value="<%=chulgo_stock%>">
				<input type="hidden" name="old_chulgo_seq" value="<%=chulgo_seq%>">
				<input type="hidden" name="old_chulgo_goods_type" value="<%=chulgo_goods_type%>">
				<input type="hidden" name="old_chulgo_att_file" value="<%=chulgo_att_file%>">
                <input type="hidden" name="old_rele_stock" value="<%=rele_stock%>">
				</form>
                </div>
			</div>
		</div>		
	</body>
</html>
