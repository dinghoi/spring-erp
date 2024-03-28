<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

rele_date = request("rele_date")
rele_no = request("rele_no")
rele_seq = request("rele_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_chulgo_reg where (rele_date = '"&rele_date&"') and (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"')"
Set Rs_reg = DbConn.Execute(SQL)
if not Rs_reg.eof then
    	rele_date = Rs_reg("rele_date")
		rele_seq = Rs_reg("rele_seq")
        rele_goods_type = Rs_reg("rele_goods_type")
        rele_emp_no = Rs_reg("rele_emp_no")
        rele_emp_name = Rs_reg("rele_emp_name")
        rele_company = Rs_reg("rele_company")
        rele_bonbu = Rs_reg("rele_bonbu")
        rele_saupbu = Rs_reg("rele_saupbu")
        rele_team = Rs_reg("rele_team")
        rele_org_name = Rs_reg("rele_org_name")
        rele_trade_name = Rs_reg("rele_trade_name")
		rele_trade_dept = Rs_reg("rele_trade_dept")
        service_no = Rs_reg("service_no")
        chulgo_ing = Rs_reg("chulgo_ing")
        chulgo_date = Rs_reg("chulgo_date")
        chulgo_stock = Rs_reg("chulgo_stock")
        chulgo_stock_name = Rs_reg("chulgo_stock_name")
    	chulgo_stock_company = Rs_reg("chulgo_stock_company")

	    if chulgo_date = "1900-01-01" then
	          chulgo_date = ""
	    end if
end if
Rs_reg.close()

sql = "select * from met_chulgo_reg_goods where (rele_date = '"&rele_date&"') and (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') ORDER BY rele_goods_seq,rele_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "출고의뢰 품목 내역"

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
				if(document.frm.in_buy_no.value =="") {
					alert('구매의뢰번호를 입력하세요');
					frm.in_buy_no.focus();
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
				<form action="met_chulgo_cust_detail.asp?buy_no=<%=buy_no%>" method="post" name="frm">
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
							  <th>회사</th>
							  <td class="left"><%=rele_company%>&nbsp;</td>
							  <th>소속</th>
							  <td class="left"><%=rele_org_name%>&nbsp;</td>
							  <th>신청자</th>
							  <td class="left"><%=rele_emp_name%>(<%=rele_emp_no%>)</td>
						    </tr>
							<tr>
							  <th>출고의뢰일</th>
							  <td class="left"><%=rele_date%>&nbsp;</td>
                              <th>서비스번호</th>
							  <td class="left"><%=service_no%>&nbsp;</td>
							  <th>고객사</th>
							  <td class="left"><%=rele_trade_name%>&nbsp<%=rele_trade_name%>;</td>
						    </tr>
							<tr>
							  <th>출고창고</th>
							  <td class="left"><%=chulgo_stock_name%>(<%=chulgo_stock_company%>)&nbsp;</td>
                              <th>출고(요청)일</th>
							  <td class="left"><%=chulgo_date%>&nbsp;</td>
                              <th>납품장소</th>
							  <td class="left"><%=rele_delivery%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                </div>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고의뢰 세부 내용 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="12%" >
                            <col width="*" >
                            <col width="16%" >
							<col width="16%" >
							<col width="16%" >
							<col width="10%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">용도구분</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">의뢰수량</th>
                                <th scope="col">출고수량</th>
							</tr>
						</thead>
						<tbody>     
						<%
							do until rs.eof or rs.bof
						
						%>
							<tr>
                                <td><%=rs("rele_goods_type")%>&nbsp;</td>
								<td><%=rs("rele_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("rele_goods_code")%>&nbsp;</td>
                                <td><%=rs("rele_goods_name")%>&nbsp;</td>
                                <td><%=rs("rele_goods_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rele_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("chulgo_qty"),0)%>&nbsp;</td>
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
                        <a href="#" class="btnType04" onClick="pop_Window('met_chulgo_referral_print.asp?rele_date=<%=rele_date%>&rele_stock=<%=rele_stock%>&rele_seq=<%=rele_seq%>&rele_goods_type=<%=rele_goods_type%>','met_chulgo_referral_print_pop','scrollbars=yes,width=750,height=600')">의뢰서 출력</a>
                        <a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>       				
	</form>
	</body>
</html>

