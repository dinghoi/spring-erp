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
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_chulgo where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
Set rs = DbConn.Execute(SQL)
if not rs.eof then
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

   else
		chulgo_goods_type = ""
        chulgo_id = ""
	    service_no = ""
	    chulgo_trade_name = ""
	    chulgo_trade_dept = ""
	    chulgo_type = ""
	    service_acpt = ""
	
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
		
		chulgo_att_file = ""
		
		rele_no = ""
		rele_seq = ""
		rele_date = ""
		rele_stock = ""
		rele_stock_company = ""
		rele_stock_name = ""
		rele_company = ""
		rele_saupbu = ""
		rele_emp_no = ""
		rele_emp_name = ""
end if
rs.close()

sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"') ORDER BY cg_goods_seq,cg_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "N/W실출고 상세조회"

view_att_file = chulgo_att_file
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
				a=confirm('출고를 취소하겠습니까?')
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
					document.frm.action = "met_buy_approve_ok.asp";
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
				<form method="post" name="frm" action="met_buy_cancel.asp">
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
							    <th>출고일자</th>
							    <td class="left"><%=chulgo_date%></td>
							    <th>출고번호</th>
							    <td class="left"><%=rele_no%>&nbsp;<%=rele_seq%></td>
                                <th>용도구분</th>
							    <td class="left"><%=chulgo_goods_type%></td>
						    </tr>
                            <tr>
							    <th>출고유형</th>
							    <td class="left"><%=chulgo_id%></td>
                                <th>출고담당</th>
							    <td class="left"><%=chulgo_org_name%>&nbsp;<%=chulgo_emp_name%></td>
                                <th>출고창고</th>
							    <td class="left">(<%=chulgo_stock_company%>)&nbsp;<%=chulgo_stock_name%></td>
						    </tr>
                            <tr>
							    <th>요청그룹사</th>
							    <td class="left"><%=rele_company%></td>
                                <th>요청사업부</th>
							    <td class="left"><%=rele_saupbu%>&nbsp;</td>
                                <th>요청(입고)창고</th>
							    <td class="left"><%=rele_stock_name%>&nbsp;(<%=rele_stock_company%>)</td>
						    </tr>
                            <tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=chulgo_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">◈ 출고 세부 내역 ◈</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="12%" >
                            <col width="12%" >
							<col width="22%" >
                            <col width="*" >
                            <col width="10%" >
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
                                <th scope="col">출고수량</th>
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
                                <td><%=rs("cg_goods_type")%>&nbsp;</td>
                                <td><%=rs("cg_goods_grade")%>&nbsp;</td>
								<td><%=rs("cg_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("cg_goods_code")%>&nbsp;</td>
                                <td><%=rs("cg_goods_name")%>&nbsp;</td>
                                <td><%=rs("cg_standard")%>&nbsp;</td>
                                
                                <td class="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
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
                           If chulgo_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=chulgo_att_file%>"><%=chulgo_att_file%></a>
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
                            <span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                            <span class="btnType01"><input type="button" value="출고증 출력" onClick="pop_Window('met_chulgo_chulgo_referral_print01.asp?chulgo_date=<%=chulgo_date%>&chulgo_stock=<%=chulgo_stock%>&chulgo_seq=<%=chulgo_seq%>','met_chulgo_chulgo_print_pop','scrollbars=yes,width=750,height=600')"></span>
                            <span class="btnType01"><input type="button" value="인수증 출력" onClick="pop_Window('met_chulgo_chulgo_receip_print01.asp?chulgo_date=<%=chulgo_date%>&chulgo_stock=<%=chulgo_stock%>&chulgo_seq=<%=chulgo_seq%>','met_chulgo_chulgo_receip_pop','scrollbars=yes,width=750,height=600')"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="chulgo_date" value="<%=chulgo_date%>">
					<input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>">
					<input type="hidden" name="chulgo_seq" value="<%=chulgo_seq%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

