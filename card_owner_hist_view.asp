<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

card_no = request("card_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs_hist = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from card_owner INNER JOIN memb ON card_owner.emp_no = memb.emp_no where card_no ='" + card_no + "'"
'response.write(sql)
set rs = dbconn.execute(sql)

sql = "select * from card_owner_history where card_no = '" + card_no + "' order by history_seq desc"
'response.write(sql)
rs_hist.Open Sql, Dbconn, 1

title_line = "카드 사용자 변경 History 조회"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
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
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>카드종류</th>
							  <td class="left"><%=rs("card_type")%></td>
							  <th>카드번호</th>
							  <td class="left"><%=rs("card_no")%></td>
					      	</tr>
							<tr>
							  <th>발행구분</th>
							  <td class="left"><%=rs("card_issue")%>&nbsp;</td>
							  <th>사용한도</th>
							  <td class="left"><%=rs("card_limit")%>&nbsp;</td>
					      	</tr>
							<tr>
							  <th>유효기간</th>
							  <td class="left"><%=rs("valid_thru")%>&nbsp;</td>
							  <th>발급일</th>
							  <td class="left"><%=rs("create_date")%>&nbsp;</td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="20%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">사용자</th>
								<th scope="col">부서</th>
								<th scope="col">시작일</th>
								<th scope="col">종료일</th>
								<th scope="col" class="left">이전사유</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first">00</td>
								<td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("start_date")%></td>
								<td>&nbsp;</td>
								<td class="left">현재 사용중</td>
							</tr>
						<%
                        do until rs_hist.eof 
                        %>
							<tr>
								<td class="first"><%=rs_hist("history_seq")%></td>
								<td><%=rs_hist("emp_name")%>&nbsp;<%=rs_hist("emp_job")%></td>
								<td><%=rs_hist("org_name")%></td>
								<td><%=rs_hist("start_date")%></td>
								<td><%=rs_hist("end_date")%></td>
								<td class="left"><%=rs_hist("mod_memo")%></td>
							</tr>
						<%
                            rs_hist.movenext()  
                        loop
                        rs_hist.Close()
                        %>
						</tbody>
					</table>                    
					<br>
				</form>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				</div>
			</div>
	</body>
</html>

