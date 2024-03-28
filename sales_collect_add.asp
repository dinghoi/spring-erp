<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
u_type = request("u_type")
approve_no = request("approve_no")
curr_date = mid(now(),1,10)

Sql="select * from saupbu_sales where approve_no = '"&approve_no&"'"
Set rs_etc=DbConn.Execute(Sql)

sql_sales="select * from sales_collect where approve_no = '"&approve_no&"' order by collect_date , collect_seq asc"
rs.Open sql_sales, Dbconn, 1

title_line = "수금 등록"

bill_collect = "현금"

collect_amt = 0
collect_id = "1"
end_date = curr_date
collect_date = curr_date
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>수금 등록</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=collect_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=bill_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=unpaid_due_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=end_date%>" );
			});	  

			function goAction () {
		  		 window.close () ;
			}

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.collect_id[" + j + "].checked")) {
						k = j
					}
				}

				if(k==0){
					if(document.frm.collect_amt.value == 0) {
						alert('수금금액을 입력하세요.');
						frm.collect_amt.focus();
						return false;}
				}

				if(k==1 || k==2){
					if(document.frm.unpaid_due_date.value == "") {
						alert('미수금예정일을 입력하세요.');
						frm.unpaid_due_date.focus();
						return false;}
					if(document.frm.change_memo.value == "") {
						alert('변동사항을 입력하세요.');
						frm.change_memo.focus();
						return false;}
					if(document.frm.unpaid_memo.value == "") {
						alert('미수금사유를 입력하세요.');
						frm.unpaid_memo.focus();
						return false;}
				}

				if(k==3){
					if(document.frm.end_date.value == "") {
						alert('완료일자를 입력하세요.');
						frm.end_date.focus();
						return false;}
					if(document.frm.change_memo1.value == "") {
						alert('완료사유를 입력하세요.');
						frm.change_memo1.focus();
						return false;}
				}
				
				kk = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.bill_collect[" + j + "].checked")) {
						kk = j
					}
				}
				
				if(kk==1) {
					if(document.frm.bill_date.value =="") {
						frm.bill_date.focus();
						alert('만기일을 입력하세요');
						return false;}}

				{
				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function condi_view() {
				if (eval("document.frm.collect_id[0].checked")) {
					document.getElementById('collect_view1').style.display = '';
					document.getElementById('collect_view2').style.display = '';
					document.getElementById('end_view').style.display = 'none';
					document.getElementById('unpaid_view1').style.display = 'none';
					document.getElementById('unpaid_view2').style.display = 'none';
				}	
				if (eval("document.frm.collect_id[1].checked")) {
					document.getElementById('collect_view1').style.display = 'none';
					document.getElementById('collect_view2').style.display = 'none';
					document.getElementById('end_view').style.display = 'none';
					document.getElementById('unpaid_view1').style.display = '';
					document.getElementById('unpaid_view2').style.display = '';
				}	
				if (eval("document.frm.collect_id[2].checked")) {
					document.getElementById('collect_view1').style.display = 'none';
					document.getElementById('collect_view2').style.display = 'none';
					document.getElementById('end_view').style.display = 'none';
					document.getElementById('unpaid_view1').style.display = '';
					document.getElementById('unpaid_view2').style.display = '';
				}	
				if (eval("document.frm.collect_id[3].checked")) {
					document.getElementById('collect_view1').style.display = 'none';
					document.getElementById('collect_view2').style.display = 'none';
					document.getElementById('end_view').style.display = '';
					document.getElementById('unpaid_view1').style.display = 'none';
					document.getElementById('unpaid_view2').style.display = 'none';
				}	
			}
			function bill_view() {
				if (eval("document.frm.bill_collect[0].checked")) {
					document.getElementById('bill_date_view').style.display = 'none';
				}	
				if (eval("document.frm.bill_collect[1].checked")) {
					document.getElementById('bill_date_view').style.display = '';
				}	
				if (eval("document.frm.bill_collect[2].checked")) {
					document.getElementById('bill_date_view').style.display = 'none';
				}	
				if (eval("document.frm.bill_collect[3].checked")) {
					document.getElementById('bill_date_view').style.display = 'none';
				}	
			}
        </script>
	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="sales_collect_add_save.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="9%" >
							<col width="13%" >
							<col width="9%" >
							<col width="*" >
							<col width="9%" >
							<col width="13%" >
							<col width="9%" >
							<col width="13%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>전표번호</th>
							  <td class="left"><%=mid(rs_etc("slip_no"),1,17)%></td>
							  <th>거래처명</th>
							  <td class="left"><%=rs_etc("company")%></td>
							  <th>매출일자</th>
							  <td class="left"><%=rs_etc("sales_date")%></td>
							  <th>매출총액</th>
							  <td class="left"><%=formatnumber(rs_etc("sales_amt"),0)%></td>
			              </tr>
						</tbody>
                    </table>
	        <h3 class="stit">* 입금 내역</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
							<col width="6%" >
							<col width="7%" >
							<col width="5%" >
							<col width="10%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">유형</th>
								<th scope="col">수금자<br>입력자</th>
								<th scope="col">수금일자<br>입력일자</th>
								<th scope="col">수금<br>방법</th>
								<th scope="col">수금금액</th>
								<th scope="col">만기일</th>
								<th scope="col">입금<br>예정일</th>
								<th scope="col">미수금<br>예정일</th>
								<th scope="col">변동사항</th>
								<th scope="col">미수금사유</th>
							</tr>
						</thead>
						<tbody>
						<%
                        i = 0
						tot_collect = 0
                        do until rs.eof 
							i = i + 1
							tot_collect = tot_collect + int(rs("collect_amt"))
 							if rs("collect_id") = "1" then
								collect_id_view = "입금"
							end if
 							if rs("collect_id") = "2" then
								collect_id_view = "미수금변경"
							end if
 							if rs("collect_id") = "3" then
								collect_id_view = "예정일변경"
							end if
 							if rs("collect_id") = "4" then
								collect_id_view = "입금완료"
							end if
                        %>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=collect_id_view%></td>
								<td><%=rs("reg_name")%></td>
								<td><%=rs("collect_date")%></td>
								<td><%=rs("bill_collect")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("collect_amt"),0)%></td>
								<td><%=rs("bill_date")%>&nbsp;</td>
								<td><%=rs("collect_due_date")%>&nbsp;</td>
								<td><%=rs("unpaid_due_date")%>&nbsp;</td>
								<td><%=rs("change_memo")%>&nbsp;</td>
								<td><%=rs("unpaid_memo")%>&nbsp;</td>
							</tr>
<%
                            rs.movenext()  
                        loop
                        rs.Close()
                        %>
							<tr bgcolor="#FFE8E8">
								<td class="first">총계</td>
								<td colspan="10">총 매출액 : <%=formatnumber(rs_etc("sales_amt"),0)%>&nbsp;&nbsp;,&nbsp;총 입금액 : <%=formatnumber(tot_collect,0)%>&nbsp;&nbsp;,&nbsp;미수금 총액 : <%=formatnumber(rs_etc("sales_amt")-tot_collect,0)%></td>
							</tr>
						</tbody>
					</table>                    
					<br>
					<h3 class="stit">* 입금 등록</h3>
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="12%">
							<col width="21%" >
							<col width="12%">
							<col width="22%" >
							<col width="12%">
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>입력유형</th>
							  <td colspan="5" class="left">
                              <input type="radio" name="collect_id" value="1" <% if collect_id = "1" then %>checked<% end if %> style="width:50px" id="Radio5" onClick="condi_view()">입금입력
  							  <input type="radio" name="collect_id" value="2" <% if collect_id = "2" then %>checked<% end if %> style="width:50px" id="Radio5" onClick="condi_view()">미수금정보입력
                              <input type="radio" name="collect_id" value="3" <% if collect_id = "3" then %>checked<% end if %> style="width:50px" id="Radio5" onClick="condi_view()">미수금예정일->입금예정일
  							  <input type="radio" name="collect_id" value="4" <% if collect_id = "4" then %>checked<% end if %> style="width:50px" id="Radio5" onClick="condi_view()">입금완료처리
                              </td>
						    </tr>
							<tr id="collect_view1">
							  <th>입금예정일</th>
							  <td class="left"><%=rs_etc("collect_due_date")%></td>
							  <th>수금일자</th>
							  <td class="left"><input name="collect_date" type="text" readonly="true" id="datepicker" style="width:70px;"></td>
							  <th>수금금액</th>
							  <td class="left"><input name="collect_amt" type="text" id="collect_amt" style="width:100px;text-align:right" onKeyUp="plusComma(this);" value="<%=collect_amt%>" ></td>
	                      	</tr>
                            <tr id="collect_view2">
                              <th>수금방법</th>
							  <td colspan="5" class="left">
                              <input type="radio" name="bill_collect" value="현금" <% if bill_collect = "현금" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="bill_view()">현금
  							  <input type="radio" name="bill_collect" value="어음" <% if bill_collect = "어음" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="bill_view()">어음
                              <input type="radio" name="bill_collect" value="카드" <% if bill_collect = "카드" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="bill_view()">카드
  							  <input type="radio" name="bill_collect" value="외환" <% if bill_collect = "외환" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="bill_view()">외환

                    		  <span class="left" style="display:none" id="bill_date_view">&nbsp;&nbsp;<strong>만기일 : &nbsp;</strong>
                    		  <input name="bill_date" type="text" readonly="true" id="datepicker1" style="width:70px;">
                    		  </span></td>
						    </tr>
							<tr style="display:none" id="unpaid_view1">
							  <th>미수금예정일</th>
							  <td class="left"><input name="unpaid_due_date" type="text" readonly="true" id="datepicker2" style="width:70px;"></td>
							  <th>변동사항</th>
							  <td colspan="3" class="left"><input name="change_memo" type="text" id="change_memo" style="width:450px" onKeyUp="checklength(this,80)"></td>
						    </tr>
							<tr style="display:none" id="end_view">
							  <th>완료일자</th>
							  <td class="left"><input name="end_date" type="text" readonly="true" id="datepicker3" style="width:70px;"></td>
							  <th>완료사유</th>
							  <td colspan="3" class="left"><input name="change_memo1" type="text" id="change_memo1" style="width:450px" onKeyUp="checklength(this,80)"></td>
						  	</tr>
							<tr style="display:none" id="unpaid_view2">
							  <th>미수금 사유</th>
							  <td colspan="5" class="left"><textarea name="unpaid_memo" id="unpaid_memo" rows="3"></textarea></td>
						    </tr>
						</tbody>
					</table>
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();"></span>            
                        <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    </div>
                    <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                    <input type="hidden" name="slip_no" value="<%=rs_etc("slip_no")%>" ID="Hidden1">
                    <input type="hidden" name="approve_no" value="<%=approve_no%>" ID="Hidden1">
                    <input type="hidden" name="collect_due_date" value="<%=rs_etc("collect_due_date")%>" ID="Hidden1">
                    <input type="hidden" name="collect_tot_amt" value="<%=rs_etc("collect_tot_amt")%>" ID="Hidden1">
                    <input type="hidden" name="sales_amt" value="<%=rs_etc("sales_amt")%>" ID="Hidden1">
				</form>
				</div>
			</div>
	</body>
</html>

