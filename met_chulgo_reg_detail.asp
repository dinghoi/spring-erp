<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

rele_no = request("rele_no")
rele_seq = request("rele_seq")
rele_date = request("rele_date")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_chulgo_reg where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')"
Set Rs_chul = DbConn.Execute(SQL)
if not Rs_chul.eof then
    	rele_no = Rs_chul("rele_no")
        rele_seq = Rs_chul("rele_seq")
	    rele_date = Rs_chul("rele_date")
        rele_id = Rs_chul("rele_id")
        rele_goods_type = Rs_chul("rele_goods_type")
		
		rele_stock = Rs_chul("rele_stock")
        rele_stock_company = Rs_chul("rele_stock_company")
        rele_stock_name = Rs_chul("rele_stock_name")
		
        rele_emp_no = Rs_chul("rele_emp_no")
        rele_emp_name = Rs_chul("rele_emp_name")
        rele_company = Rs_chul("rele_company")
        rele_bonbu = Rs_chul("rele_bonbu")
        rele_saupbu = Rs_chul("rele_saupbu")
        rele_team = Rs_chul("rele_team")
        rele_org_name = Rs_chul("rele_org_name")
        rele_trade_name = Rs_chul("rele_trade_name")
	    rele_trade_dept = Rs_chul("rele_trade_dept")
	    rele_delivery = Rs_chul("rele_delivery")
        service_no = Rs_chul("service_no")
        chulgo_ing = Rs_chul("chulgo_ing")
        chulgo_date = Rs_chul("chulgo_date")
        chulgo_stock = Rs_chul("chulgo_stock")
        chulgo_stock_name = Rs_chul("chulgo_stock_name")
	    chulgo_stock_company = Rs_chul("chulgo_stock_company")
	    rele_att_file = Rs_chul("rele_att_file")
	    rele_memo = Rs_chul("rele_memo")
        rele_sign_yn = Rs_chul("rele_sign_yn")
	    rele_sign_no = Rs_chul("rele_sign_no")
	    rele_sign_date = Rs_chul("rele_sign_date")
	    if chulgo_date = "0000-00-00" then
	          chulgo_date = ""
	    end if
   else
		rele_no = ""
        rele_seq = ""
	    rele_date = ""
        rele_id = ""
        rele_goods_type = ""
		
		rele_stock = ""
        rele_stock_company = ""
        rele_stock_name = ""
		
        rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""
        rele_trade_name = ""
	    rele_trade_dept = ""
	    rele_delivery = ""
        service_no = ""
        chulgo_ing = ""
        chulgo_date = ""
        chulgo_stock = ""
        chulgo_stock_name = ""
	    chulgo_stock_company = ""
	    rele_att_file = ""
	    rele_memo = ""
        rele_sign_yn = ""
	    rele_sign_no = ""
	    rele_sign_date = ""
end if
Rs_chul.close()

sql = "select * from met_chulgo_reg_goods where (rele_no = '"&rele_no&"') and (rele_seq = '"&rele_seq&"') and (rele_date = '"&rele_date&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "출고의뢰 조회"

view_att_file = rele_att_file
path = "/met_upload"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>자재관리 시스템</title>
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
			function frmcheck () {
				if (chkfrm()) {
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
						
				{
				a=confirm('출고의뢰를 취소하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
//			function approve_request(slip_id,slip_no,slip_seq) 
			function approve_request() 
				{
				a=confirm('결재 요청하시겠습니까?')
				if (a==true) {
//					document.frm.action = "met_buy_approve_ok.asp?slip_id="+slip_id+'&slip_no='+slip_no+'&slip_seq='+slip_seq;
					document.frm.action = "met_chulgo_reg_approve_ok.asp";
					document.frm.submit();
				}
				return false;
				}
		</script>

	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
				<h3 class="insa"><%=title_line%></h3>
				<form method="post" name="frm" action="met_chulgo_reg_cancel.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody> 
							<tr>
                                <th>의뢰번호(일자)</th>
							    <td class="left"><%=rele_no%>&nbsp;<%=rele_seq%>&nbsp;&nbsp;&nbsp;(<%=rele_date%>)</td>
							    <th>용도구분</th>
							    <td class="left"><%=rele_goods_type%>&nbsp;</td>
							    <th>의뢰창고</th>
							    <td class="left"><%=rele_stock_name%>(<%=rele_stock_company%>)&nbsp;</td>
 							</tr>
                            <tr>
							    <th>의뢰회사</th>
							    <td class="left"><%=rele_company%></td>
							    <th>사업부</th>
							    <td class="left"><%=rele_saupbu%></td>
							    <th>신청자</th>
							    <td class="left"><%=rele_emp_name%>&nbsp;<%=rele_org_name%></td>
						    </tr>
							<tr>
                                <th>출고요청일</th>
							    <td class="left"><%=chulgo_date%></td>
							    <th>출고창고</th>
							    <td colspan="3" class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=rele_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고의뢰 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="10%" >
                            <col width="9%" >
							<col width="10%" >
							<col width="14%" >
							<col width="6%" >
                            
                            <col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">용도구분</th>
                                <th scope="col">상태</th>
                                <th scope="col">품목구분</th>
                                <th scope="col">품목코드</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
								<th scope="col">의뢰수량</th>
                                
                                <th scope="col">서비스No</th>
                                <th scope="col">고객사</th>
                                <th scope="col">점소명</th>
                                <th scope="col">비고(사유)</th>
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
					
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("rl_goods_type")%>&nbsp;</td>
                                <td><%=rs("rl_goods_grade")%>&nbsp;</td>
								<td><%=rs("rl_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("rl_goods_code")%>&nbsp;</td>
                                <td><%=rs("rl_goods_name")%>&nbsp;</td>
                                <td><%=rs("rl_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
                                
                                <td><%=rs("rl_service_no")%>&nbsp;</td>
                                <td><%=rs("rl_trade_name")%>&nbsp;</td>
                                <td><%=rs("rl_trade_dept")%>&nbsp;</td>
                                <td class="left"><%=rs("rl_bigo")%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
                    <br>
                    <table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody>
                            <tr>
							  <th>첨부</th>
							  <td colspan="5" class="left">
                        <% 
                           If rele_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rele_att_file%>"><%=rele_att_file%></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						   End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
          	     <br>
     				<div class="noprint">
                        <div align=center>
                        <% if rele_sign_yn = "N" then	%>
                            <span class="btnType01"><input type="button" value="결재요청" onclick="javascript:approve_request('<%=rele_no%>','<%=rele_seq%>','<%=rele_date%>');"></span>
                        <% end if	%>
                        <% if cancel_yn = "Y" then	%>
                            <span class="btnType01"><input type="button" value="취소" onclick="javascript:frmcheck();"></span>
                        <% end if	%>
                            <span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                            <span class="btnType01"><input type="button" value="의뢰서 출력" onClick="pop_Window('met_chulgo_referral_print.asp?rele_no=<%=rele_no%>&rele_date=<%=rele_date%>&rele_seq=<%=rele_seq%>','met_chulgo_referral_print_pop','scrollbars=yes,width=940,height=600')"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="rele_no" value="<%=rele_no%>">
					<input type="hidden" name="rele_seq" value="<%=rele_seq%>">
					<input type="hidden" name="rele_date" value="<%=rele_date%>">
					<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>">      				
	     </form>
    	</div>				
	  </div>     
	</body>
</html>
