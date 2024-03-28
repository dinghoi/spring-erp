<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

chulgo_date = request("chulgo_date")
chulgo_stock = request("chulgo_stock")
chulgo_seq = request("chulgo_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_mv_go where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
Set Rs_reg = DbConn.Execute(SQL)
if not Rs_reg.eof then
    	chulgo_date = Rs_reg("chulgo_date")
        chulgo_stock = Rs_reg("chulgo_stock")
		chulgo_seq = Rs_reg("chulgo_seq")
        chulgo_id = Rs_reg("chulgo_id")
        chulgo_type = Rs_reg("chulgo_type")
        chulgo_stock_company = Rs_reg("chulgo_stock_company")
        chulgo_stock_name = Rs_reg("chulgo_stock_name")
		chulgo_emp_no = Rs_reg("chulgo_emp_no")
        chulgo_emp_name = Rs_reg("chulgo_emp_name")
		rele_date = Rs_reg("rele_date")
        rele_stock = Rs_reg("rele_stock")
		rele_seq = Rs_reg("rele_seq")
        rele_stock_company = Rs_reg("rele_stock_company")
        rele_stock_name = Rs_reg("rele_stock_name")
        rele_emp_no = Rs_reg("rele_emp_no")
        rele_emp_name = Rs_reg("rele_emp_name")
		rele_goods_type = Rs_reg("rele_goods_type")
		rele_memo = Rs_reg("rele_memo")
		in_stock_date = Rs_reg("in_stock_date")

	    if in_stock_date = "1900-01-01" then
	          in_stock_date = ""
	    end if
		if in_stock_date = "" or isnull(in_stock_date) then
	            in_stock = "이동중"
		   else 
		        in_stock = "입고완료"
	    end if
end if
Rs_reg.close()

sql = "select * from met_mv_go_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"') ORDER BY chulgo_goods,chulgo_goods_seq ASC"

Rs.Open Sql, Dbconn, 1

title_line = "창고이동 출고 품목 내역"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>자재관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
				if(document.frm.rele_date.value =="") {
					alert('의뢰일자를 입력하세요');
					frm.rele_date.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_move_reg_detail.asp?rele_date=<%=rele_date%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="*" >
							<col width="12%" >
							<col width="20%" >
							<col width="12%" >
							<col width="20%" >
						</colgroup>
						<tbody> 
							<tr>
							  <th style="background:#f8f8f8;">의뢰회사</th>
							  <td class="left"><%=rele_stock_company%>&nbsp;</td>
							  <th style="background:#f8f8f8;">의뢰창고</th>
							  <td class="left"><%=rele_stock_name%>(<%=rele_stock%>)&nbsp;</td>
							  <th style="background:#f8f8f8;">의뢰담당</th>
							  <td class="left"><%=rele_emp_name%>(<%=rele_emp_no%>)</td>
						    </tr>
							<tr>
							  <th style="background:#f8f8f8;">의뢰일자</th>
							  <td colspan="3" class="left"><%=rele_date%>&nbsp;</td>
                              <th style="background:#f8f8f8;">출고형태</th>
							  <td class="left"><%=chulgo_type%>&nbsp;</td>
						    </tr>
							<tr>
							  <th style="background:#f8f8f8;">출고처창고</th>
							  <td class="left"><%=chulgo_stock_name%>(<%=chulgo_stock_company%>)&nbsp;</td>
                              <th style="background:#f8f8f8;">실출고일</th>
							  <td class="left"><%=chulgo_date%>&nbsp;</td>
                              <th style="background:#f8f8f8;">의뢰창고<br>입고일</th>
							  <td class="left"><%=in_stock_date%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th style="background:#f8f8f8;">비고</th>
							  <td colspan="5" class="left"><%=rele_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                </div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 창고이동 출고 세부 내용 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="12%" >
                            <col width="*" >
                            <col width="16%" >
							<col width="16%" >
							<col width="16%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">요청수량</th>
                                <th scope="col">출고수량</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>     
						<%
							do until rs.eof or rs.bof
						
						%>
							<tr>
                                <td><%=rs("chulgo_goods_type")%>&nbsp;</td>
								<td><%=rs("chulgo_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("chulgo_goods")%>&nbsp;</td>
                                <td><%=rs("chulgo_goods_name")%>&nbsp;</td>
                                <td><%=rs("chulgo_standard")%>&nbsp;</td>
                                <td style="width:70px;text-align:right"><%=formatnumber(rs("rele_qty"),0)%>&nbsp;</td>
                                <td style="width:70px;text-align:right"><%=formatnumber(rs("chulgo_qty"),0)%>&nbsp;</td>
                                <td><%=in_stock%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
			</div>				
	   </div>     
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onClick="pop_Window('met_move_go_receipt_print.asp?chulgo_date=<%=chulgo_date%>&chulgo_stock=<%=chulgo_stock%>&chulgo_seq=<%=chulgo_seq%>','met_move_go_receipt_print_pop','scrollbars=yes,width=750,height=600')">인수증 출력</a>
                        <a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>       				
	</form>
	</body>
</html>

