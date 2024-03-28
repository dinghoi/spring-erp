<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	card_gubun = request.form("card_gubun")
	slip_month = request.form("slip_month")
	objFile = request.form("objFile")

'	objFile = SERVER.MapPath(".") & "\srv_upload\주소록.xls"
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	dbconn.BeginTrans

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"

	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
	read_cnt = 0
	write_cnt = 0
    if rowcount > -1 then
		for i=0 to rowcount
			if xgr(1,i) = "" or isnull(xgr(1,i)) then
				exit for
			end if
	' BC카드
			if card_gubun = "BC카드" then
				if trim(xgr(0,i)) = "신규" then
					cancel_yn = "N"
				  else
				  	cancel_yn = "Y"
				end if
				slip_date = xgr(8,i)
				card_no = xgr(1,i)
				customer = xgr(22,i)
				customer_no = xgr(21,i)
				upjong = replace(xgr(23,i)," ","")
				price = int(xgr(11,i))
				cost_vat = int(xgr(15,i))
				approve_no = xgr(7,i)
			end If
	'//2017-06-08 add. kb국민카드
		if card_gubun = "kb국민카드" Then
			  if trim(xgr(15,i)&"") = "정상" then
					cancel_yn = "N"
				  else
					cancel_yn = "Y"
				end if
				slip_date = xgr(0,i)
				If Trim(slip_date&"")<>"" Then slip_date = Replace(slip_date,".","-")

				card_num = xgr(4,i)
				card_no = xgr(4,i)
				card_no = right(card_no,7)
				customer = xgr(6,i)
				customer_no = xgr(18,i)
				upjong = xgr(7,i)
				price = xgr(10,i)
				cost_vat = xgr(11,i)
				approve_no = xgr(14,i)
				if price = "" or isnull(price) then
				price = "'"&xgr(11,i)
				end if
			end If
	' 씨티카드
			if card_gubun = "씨티카드" then
				if trim(xgr(14,i)) = "정상" then
					cancel_yn = "N"
				else
					cancel_yn = "Y"
				end if
				slip_date = xgr(4,i)
				imsi_no = xgr(1,i)
				card_no = mid(imsi_no,1,4) & "-" & mid(imsi_no,5,4) & "-" & mid(imsi_no,9,4) & "-" &right(imsi_no,4)
				customer = xgr(8,i)
				customer_no = xgr(9,i)
				upjong = replace(trim(xgr(17,i))," ","")
				price = int(xgr(10,i))	'20180402 eulro
				cost_vat = int(xgr(21,i))
				approve_no = xgr(20,i)
			end if

			' 신한카드 	L(9410-6440-9)
			If card_gubun = "신한카드" then

				'신규 작성[허정호_20201215]	=======================
				slip_date = replace(xgr(6,i),".","-")

				'imsi_no = xgr(0,i)
				'card_no = mid(imsi_no,2,3) & "-" &right(imsi_no,4)
				card_no = xgr(0, i)	'이용 카드(카드 번호)

				customer = xgr(19,i)
				imsi_no = xgr(18,i)
				customer_no = mid(imsi_no,1,3) & "-" & mid(imsi_no,4,2) & "-" &right(imsi_no,5)
				upjong = replace(xgr(20,i)," ","")
				price = int(xgr(9,i))
				cost_vat = int(xgr(12,i))
				approve_no = xgr(5,i)

				'approve_no = xgr(2,i)	'승인번호
				'slip_date = Replace(xgr(1,i), ".", "-")	'접수일[승인일자]
				'card_no = xgr(3, i)	'이용 카드(카드 번호)
				'customer = xgr(5, i)	'가맹점명
				'customer_no = ""	'가맹점 번호
				'upjong = ""	'업종
				'price = xgr(6, i)	'이용금액

				'slip_date = Replace(Left(xgr(0, i), 10), ".", "-")	'접수일[승인일자]
				'approve_no = xgr(2, i)	'승인번호
				'card_no = xgr(3, i)	'이용 카드(카드 번호)
				'customer = xgr(5, i)	'가맹점 명
				'customer_no = ""	'가맹점 번호
				'upjong = xgr(6, i)	'업종
				'price = xgr(7, i)	'이용금액

				If xgr(10,i) = "국내" Then
					cost_vat = price - Int(price/1.1)
				Else
				  	cost_vat = 0
				End If

				If price < 0 Then
					cancel_yn = "Y"
				Else
				  	cancel_yn = "N"
				End If
			End If
	' 롯데카드 	LOCAL -> 첫 4자리 9409, AMEX -> 첫 4자리 3762 , VISA -> 첫 4자리 4670
			if card_gubun = "롯데카드" then
				slip_date = replace(xgr(5,i),".","-")

				imsi_no = xgr(2,i)
'				if xgr(1,i) = "LOCAL" then
'					card_no = "9409" + mid(imsi_no,5)
'				  elseif xgr(1,i) = "VISA" then
'					card_no = "4670" + mid(imsi_no,5)
'				  else
'					card_no = "3762" + mid(imsi_no,5)
'				end if
				imsi_card_no = right(imsi_no,3)
				sql = "select * from card_owner where card_type like '%롯데%' and right(card_no,3) = '"&imsi_card_no&"'"
				set rs_card=dbconn.execute(sql)
				if rs_card.eof or rs_card.bof then
					card_no = imsi_no
				  else
					card_no = rs_card("card_no")
				end if

				customer = xgr(7,i)
				customer_no = xgr(25,i)
				upjong = replace(xgr(26,i)," ","")
				price = int(xgr(8,i))
				if xgr(1,i) = "LOCAL" then
					cost_vat = price - int(price/1.1)
				  else
				  	cost_vat = 0
				end if
				approve_no = xgr(21,i)

				if trim(xgr(12,i)) = "매입여부" then
					cancel_yn = "N"
				  else
				  	cancel_yn = "Y"
				end if
			end if

		if approve_no = "" or isnull(approve_no) or approve_no=  " " then
			approve_no = cstr(mid(slip_date,1,4)) + cstr(mid(slip_date,6,2)) + cstr(mid(slip_date,9,2))
		end if

		cost = price - cost_vat

		sql = "select * from card_upjong where card_upjong = '" & upjong &"'"
		set rs_upjong=dbconn.execute(sql)
		if rs_upjong.eof or rs_upjong.bof then
			tot_upjong = tot_upjong + 1
			account = "미등록"
			account_item = "미등록"
		  else
			account = rs_upjong("account")
			account_item = rs_upjong("account_item")
			if account = "" or account_item = "" or isnull(account) or isnull(account_item) then
				tot_account = tot_account + 1
			end if
			if rs_upjong("tax_yn") = "N" then
				if cost_vat <> 0 then
					cost_vat = 0
					cost = price
				end if
			end if
			if rs_upjong("tax_yn") = "Y" then
				if cost_vat = 0 then
					cost_vat = clng((price/1.1)/10)
					cost = price - cost_vat
				end if
			end if
		end if

		if card_gubun = "신한카드" Then
			'카드 번호 조회_앞 8자리 비교 조회에서 앞 4자리, 뒤 4자리로 변경[허정호_20210104]
			'sql = "select * from card_owner where right(card_no,8) = '" & card_no &"'"
			sql = "SELECT card_no, car_vat_sw, card_type, emp_no, owner_company FROM card_owner "
			sql = sql & "WHERE RIGHT(card_no, 4) = '" & Right(card_no, 4) &"' "
			sql = sql & "AND LEFT(card_no, 4) = '" & Left(card_no, 4) &"' "
		Elseif card_gubun = "kb국민카드" then
			'sql = "select * from card_owner where right(card_no,4) = '" & card_no &"'"
			'sql = "select * from card_owner where left(card_no,7) = '" & left(card_num,7) & "' and  right(card_no,4) = '" & right(card_num, 4) &"'"
			'쿼리 조건 변경[허정호_20220203]
			sql = "select * from card_owner where card_no = '" & card_num & "' "
		Else
			sql = "select * from card_owner where card_no = '" & card_no &"'"
		End If
		'Response.write sql

		set rs_card=dbconn.execute(sql)

		if rs_card.eof or rs_card.bof then
			card_type = "미등록"
			emp_name = "미지정"
			emp_no = ""
			emp_grade = ""
			owner_company = ""
			emp_company = ""
			bonbu = ""
			saupbu = ""
			team = ""
			org_name = ""
			reside_place = ""
			reside_company = ""
			car_vat_sw = "C"
		  else
			card_no = rs_card("card_no")
			car_vat_sw = rs_card("car_vat_sw")
			card_type = rs_card("card_type")
			emp_no = rs_card("emp_no")
			owner_company = rs_card("owner_company")
			sql = "select * from memb where user_id = '"&emp_no&"'"
			set rs_emp=dbconn.execute(sql)
			if rs_emp.eof or rs_emp.bof then
				emp_name = "미지정"
				emp_company = ""
				emp_grade = ""
				bonbu = ""
				saupbu = ""
				team = ""
				org_name = ""
				reside_place = ""
				reside_company = ""
			  else
				emp_name = rs_emp("user_name")
				emp_grade = rs_emp("user_grade")
				emp_company = rs_emp("emp_company")
				bonbu = rs_emp("bonbu")
				saupbu = rs_emp("saupbu")
				team = rs_emp("team")
				org_name = rs_emp("org_name")
				reside_place = rs_emp("reside_place")
				reside_company = rs_emp("reside_company")
			end if
		end if
' 주유카드가 바뀌면 수정 해야함
		if card_type = "롯데주유카드" then
			account = "차량유지비"
			account_item = "유류대"
' 2014년 12월부터 변경
'			car_vat_sw = "Y"
		end if
' 주유카드 변경 끝
		if account = "차량유지비" then
			if car_vat_sw = "N" then
				cost_vat = 0
				cost = price
			end if
			if car_vat_sw = "Y" then
				if cost_vat = 0 then
					cost_vat = clng((price/1.1)/10)
					cost = price - cost_vat
				end if
			end if
		end if


		customer = Replace(customer,"'","&quot;")
		read_cnt = read_cnt + 1
		'sql = "select * from card_slip where approve_no = '"&approve_no&"' and cancel_yn = '"&cancel_yn&"'"
		sql = "select * from card_slip where approve_no = '"&approve_no&"' and cancel_yn = '"&cancel_yn&"' and card_type = '"&card_type&"' "
        set rs_card=dbconn.execute(sql)
        'Response.write sql&"<br>"

		if rs_card.eof or rs_card.bof then
            write_cnt = write_cnt + 1

        'Response.write "ㅁㄴㅇㄻㄴㅇㄻㄴㅇㄹ<br>"

            ' 해당 카드의 손익포함 여부값을 그대로 card_slip에 적용
            pl_yn = "Y"
            sql = "select pl_yn from card_owner where card_no = '"&card_no&"' "
            set rs_owner = dbconn.execute(sql)
            if not rs_owner.eof then
                pl_yn = rs_owner("pl_yn")
            end if

            sql = "insert into card_slip ( approve_no             "&chr(13)&_
                  "                      , cancel_yn              "&chr(13)&_
                  "                      , slip_date              "&chr(13)&_
                  "                      , card_type              "&chr(13)&_
                  "                      , card_no                "&chr(13)&_
                  "                      , emp_no                 "&chr(13)&_
                  "                      , emp_name               "&chr(13)&_
                  "                      , emp_grade              "&chr(13)&_
                  "                      , emp_company            "&chr(13)&_
                  "                      , bonbu                  "&chr(13)&_
                  "                      , saupbu                 "&chr(13)&_
                  "                      , team                   "&chr(13)&_
                  "                      , org_name               "&chr(13)&_
                  "                      , reside_place           "&chr(13)&_
                  "                      , reside_company         "&chr(13)&_
                  "                      , customer               "&chr(13)&_
                  "                      , customer_no            "&chr(13)&_
                  "                      , upjong                 "&chr(13)&_
                  "                      , account                "&chr(13)&_
                  "                      , account_item           "&chr(13)&_
                  "                      , price                  "&chr(13)&_
                  "                      , cost                   "&chr(13)&_
                  "                      , cost_vat               "&chr(13)&_
                  "                      , card_gubun             "&chr(13)&_
                  "                      , account_end            "&chr(13)&_
                  "                      , person_end             "&chr(13)&_
                  "                      , end_sw                 "&chr(13)&_
                  "                      , reg_id                 "&chr(13)&_
                  "                      , reg_name               "&chr(13)&_
                  "                      , reg_date               "&chr(13)&_
                  "                      , owner_company          "&chr(13)&_
                  "                      , pl_yn                  "&chr(13)&_
                  "                      )                        "&chr(13)&_
                  "               values ( '"&approve_no&"'       "&chr(13)&_
                  "                      , '"&cancel_yn&"'        "&chr(13)&_
                  "                      , '"&slip_date&"'        "&chr(13)&_
                  "                      , '"&card_type&"'        "&chr(13)&_
                  "                      , '"&card_no&"'          "&chr(13)&_
                  "                      , '"&emp_no&"'           "&chr(13)&_
                  "                      , '"&emp_name&"'         "&chr(13)&_
                  "                      , '"&emp_grade&"'        "&chr(13)&_
                  "                      , '"&emp_company&"'      "&chr(13)&_
                  "                      , '"&bonbu&"'            "&chr(13)&_
                  "                      , '"&saupbu&"'           "&chr(13)&_
                  "                      , '"&team&"'             "&chr(13)&_
                  "                      , '"&org_name&"'         "&chr(13)&_
                  "                      , '"&reside_place&"'     "&chr(13)&_
                  "                      , '"&reside_company&"'   "&chr(13)&_
                  "                      , '"&customer&"'         "&chr(13)&_
                  "                      , '"&customer_no&"'      "&chr(13)&_
                  "                      , '"&upjong&"'           "&chr(13)&_
                  "                      , '"&account&"'          "&chr(13)&_
                  "                      , '"&account_item&"'     "&chr(13)&_
                  "                      , "&price&"              "&chr(13)&_
                  "                      , "&cost&"               "&chr(13)&_
                  "                      , "&cost_vat&"           "&chr(13)&_
                  "                      , '"&card_gubun&"'       "&chr(13)&_
                  "                      , 'N'                    "&chr(13)&_
                  "                      , 'N'                    "&chr(13)&_
                  "                      , 'N'                    "&chr(13)&_
                  "                      , '"&user_id&"'          "&chr(13)&_
                  "                      , '"&user_name&"'        "&chr(13)&_
                  "                      , now()                  "&chr(13)&_
                  "                      , '"&owner_company&"'    "&chr(13)&_
                  "                      , '"&pl_yn&"'            "&chr(13)&_
                  "                      )                        "&chr(13)
			dbconn.execute(sql)
			'Response.write "<pre>"&sql&"</pre><br>"
		else
			'Response.write card_no&"<br>"
			error_cards = error_cards & ", " & card_no
		end if

		'Response.write card_no&"<br>"
		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "변경중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
	end if

	err_msg = "총 " & cstr(read_cnt) & "건 읽고 " & cstr(write_cnt) & " 건 처리되었습니다... ("& error_cards &") "
	'Response.write err_msg&"<br>"
	response.write"<script language=javascript>"
	response.write"alert('"&err_msg&"');"
    response.write"location.replace('card_slip_up.asp');"
	response.write"</script>"
	Response.End

	rs.close
	cn.close

	set rs = nothing
	set cn = nothing

%>