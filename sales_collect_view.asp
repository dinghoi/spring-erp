<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
u_type = request("u_type")
approve_no = request("approve_no")

Sql="select * from saupbu_sales where approve_no = '"&approve_no&"'"
Set rs_etc=DbConn.Execute(Sql)

'sql_sales="select * from sales_collect where approve_no = '"&approve_no&"' and collect_amt <> 0 order by collect_date , collect_seq desc"
sql_sales="select * from sales_collect where approve_no = '"&approve_no&"' and (collect_amt > 0) order by collect_date , collect_seq desc"
rs.Open sql_sales, Dbconn, 1

title_line = "수금 내역"

bill_collect = "현금"
collect_amt = 0

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

			function goAction () {
		  		 window.close () ;
			}

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.collect_amt.value > 0) {
					if(document.frm.collect_date.value == "") {
						alert('수금일자를 입력하세요.');
						frm.collect_date.focus();
						return false;}}
				if(document.frm.collect_amt.value == "") {
					alert('수금금액이 NULL 입니다.');
					frm.collect_amt.focus();
					return false;}
//				if(document.frm.collect_amt.value == "" || document.frm.collect_amt.value == 0) {
//					alert('수금금액을 입력하세요.');
//					frm.collect_amt.focus();
//					return false;}
				
				k = 0;
				for (j=0;j<4;j++) {
					if (eval("document.frm.bill_collect[" + j + "].checked")) {
						k = j
					}
				}
				
				if(k==1) {
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
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="16%" >
							<col width="34%" >
							<col width="16%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>전표번호</th>
							  <td class="left"><%=mid(rs_etc("slip_no"),1,17)%></td>
							  <th>거래처명</th>
							  <td class="left"><%=rs_etc("company")%></td>
					      	</tr>
							<tr>
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
							<col width="5%" >
							<col width="19%" >
							<col width="19%" >
							<col width="*" >
							<col width="19%" >
							<col width="19%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">수금자</th>
								<th scope="col">수금일자</th>
								<th scope="col">수금방법</th>
								<th scope="col">수금금액</th>
								<th scope="col">만기일</th>
							</tr>
						</thead>
						<tbody>
						<%
                        i = 0
						tot_collect = 0
                        do until rs.eof 
							i = i + 1
							tot_collect = tot_collect + int(rs("collect_amt"))
                        %>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rs("reg_name")%></td>
								<td><%=rs("collect_date")%></td>
								<td><%=rs("bill_collect")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("collect_amt"),0)%></td>
								<td><%=rs("bill_date")%>&nbsp;</td>
							</tr>
						<%
                            rs.movenext()  
                        loop
                        rs.Close()
                        %>
							<tr bgcolor="#FFE8E8">
								<td class="first">총계</td>
								<td colspan="5">총 매출액 : <%=formatnumber(rs_etc("sales_amt"),0)%>&nbsp;&nbsp;,&nbsp;총 입금액 : <%=formatnumber(tot_collect,0)%>&nbsp;&nbsp;,&nbsp;미수금 총액 : <%=formatnumber(rs_etc("sales_amt")-tot_collect,0)%></td>
							</tr>
						</tbody>
					</table>                    
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    </div>
				</form>
				</div>
			</div>
	</body>
</html>

