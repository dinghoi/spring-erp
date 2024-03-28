<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim objFile, slip_month, cn, rs
Dim rowcount, xgr, fldcount, tot_cnt, read_cnt, write_cnt
Dim as_company, as_set, set_time, as_error, as_collect, as_testing
Dim as_total, total_time
Dim end_msg, i
Dim rsCnt, rsAsCnt, where_sql
Dim as_give_cowork, as_get_cowork, as_date, start_date, end_date
Dim total_cnt, time_total, cowork_give_company, cowork_get_company, cowork_cnt
Dim rsAs, arrAs, arr_company, arr_set, arr_time, arr_error, arr_testing, arr_collect
Dim arr_total, arr_total_time, arr_give_cowork, arr_get_cowork, j
Dim rsAsTemp

objFile = f_Request("objFile")
slip_month = f_Request("slip_month")

as_date = Left(slip_month, 4) & "-" & Right(slip_month, 2)
start_date = as_date & "-01"
end_date = as_date & "-31"

DBConn.BeginTrans

'장애 건수/협업 건수 임시 정보 조회
objBuilder.Append "SELECT as_seq FROM as_temp WHERE as_month = '"&slip_month&"';"

Set rsAsTemp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAsTemp.EOF Then
	'AS 장애 건수 임시 정보 삭제
	objBuilder.Append "DELETE FROM as_temp WHERE as_month = '"&slip_month&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'협업 건수 임시 정보 삭제
	objBuilder.Append "DELETE FROM as_cowork WHERE co_month = '"&slip_month&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
rsAsTemp.Close() : Set rsAsTemp = Nothing

where_sql = "WHERE CAST(acpt_date as date) >= '"&start_date&"' AND CAST(acpt_date as date) <= '"&end_date&"';"

'신규 테이블에 AS 현황 데이터 Insert
objBuilder.Append "SELECT COUNT(*) FROM as_acpt_end " & where_sql

Set rsAsCnt = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If CInt(rsAsCnt(0)) > 0 Then
	objBuilder.Append "DELETE FROM as_acpt_end " & where_sql

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

rsAsCnt.Close() : Set rsAsCnt = Nothing

objBuilder.Append "INSERT INTO as_acpt_end "
objBuilder.Append "SELECT * FROM as_acpt " & where_sql

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
'AS 데이터 Insert End =============

'AS 현황 데이터 Insert
objBuilder.Append "SELECT COUNT(*) FROM as_acpt_status WHERE as_month = '"&slip_month&"';"

Set rsCnt = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'해당 월 데이타가 있을 경우 삭제
If CInt(rsCnt(0)) > 0 Then
	'AS 현황 정보 삭제
	objBuilder.Append "DELETE FROM as_acpt_status WHERE as_month = '"&slip_month&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
rsCnt.Close() : Set rsCnt = Nothing

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "select * from [1:10000]",cn,"0"

rowcount = -1
xgr = rs.getrows
rowcount = UBound(xgr, 2)
fldcount = rs.fields.count

tot_cnt = rowcount + 1
read_cnt = 0
write_cnt = 0

If rowcount > -1 Then
	For i=0 To rowcount
		as_company = xgr(0, i)	'거래처
		as_set = f_toString(xgr(1, i), 0)	'설치/공사
		set_time = f_toString(xgr(2, i), 0)	'설치/공사(시간)
		as_error = f_toString(xgr(3, i), 0)	'장애
		as_testing = f_toString(xgr(4, i), 0)	'점검
		as_collect = f_toString(xgr(5, i), 0)	'회수
		as_total = f_toString(xgr(6, i), 0)'합계(거래처 별 건수 합계)

		cowork_give_company = f_toString(xgr(7, i), "")'협업소속(협업한 고객사)
		cowork_get_company = f_toString(xgr(8, i), "")'고객사(협업 받은 고객사)
		cowork_cnt = f_toString(xgr(9, i), 0)'건수(협업)

		total_cnt = f_toString(xgr(10, i), 0)'총합계
		time_total = f_toString(xgr(11, i), 0)'총시간

		'거래처 별 총 시간
		total_time = Round(time_total * as_total / total_cnt, 0)

		'장애 건수 임시 테이블 insert
		If as_company <> "NaN" Then'엑셀 취합 시 장애 리스트와 협업 리스트 개수가 맞지 않을 경우 'NaN'으로 거래처 및 협업 소속에 표기
			objBuilder.Append "INSERT INTO as_temp(as_month, as_company, as_set, set_time, as_error, "
			objBuilder.Append "as_testing, as_collect, as_total, total_time, reg_date)"
			objBuilder.Append "VALUES('"&slip_month&"', '"&as_company&"', '"&as_set&"', '"&set_time&"', '"&as_error&"', "
			objBuilder.Append "'"&as_testing&"', '"&as_collect&"', '"&as_total&"', '"&total_time&"', NOW());"

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If

		'협업 건수 임시 테이블 insert
		If cowork_give_company <> "NaN" Then
			objBuilder.Append "INSERT INTO as_cowork(co_month, co_company, as_company, co_cnt, reg_date)"
			objBuilder.Append "VALUES('"&slip_month&"', '"&cowork_give_company&"', '"&cowork_get_company&"', '"&cowork_cnt&"', NOW());"

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If

		read_cnt = read_cnt + 1'읽은 개수
		write_cnt = write_cnt + 1'처리 개수
	Next

	'월별 AS 현황 입력
	objBuilder.Append "SELECT as_company, as_set, set_time, as_error, as_testing, as_collect, "
	objBuilder.Append "	as_total, total_time, "
	objBuilder.Append "	IFNULL((SELECT SUM(co_cnt) FROM as_cowork "
	objBuilder.Append "		WHERE co_month = '"&slip_month&"' AND co_company = astt.as_company), 0) AS 'as_give_cowork', "
	objBuilder.Append "	IFNULL((SELECT SUM(co_cnt) FROM as_cowork "
	objBuilder.Append "		WHERE co_month = '"&slip_month&"' AND as_company = astt.as_company), 0) AS 'as_get_cowork' "
	objBuilder.Append "FROM as_temp AS astt "
	objBuilder.Append "WHERE as_month = '"&slip_month&"';"

	Set rsAs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not rsAs.EOF Then
		arrAs = rsAs.getRows()
	End If
	rsAs.Close() : Set rsAs = Nothing

	If IsArray(arrAs) Then
		For j = LBound(arrAs) To UBound(arrAs, 2)
			arr_company = arrAs(0, j)
			arr_set = arrAs(1, j)
			arr_time = arrAs(2, j)
			arr_error = arrAs(3, j)
			arr_testing = arrAs(4, j)
			arr_collect = arrAs(5, j)
			arr_total = arrAs(6, j)
			arr_total_time = arrAs(7, j)
			arr_give_cowork = arrAs(8, j)
			arr_get_cowork = arrAs(9, j)

			objBuilder.Append "INSERT INTO as_acpt_status(as_month, as_company, as_set, set_time, as_error, as_testing, as_collect, "
			objBuilder.Append "as_total, total_time, reg_date, reg_id, as_give_cowork, as_get_cowork) "
			objBuilder.Append "VALUES('"&slip_month&"', '"&arr_company&"', '"&arr_set&"', '"&arr_time&"', '"&arr_error&"', '"&arr_testing&"', '"&arr_collect&"', "
			objBuilder.Append "'"&arr_total&"', '"&arr_total_time&"', NOW(), '"&user_id&"', '"&arr_give_cowork&"', '"&arr_get_cowork&"');"

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		Next
	End If
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "총 " & cstr(read_cnt) & "건 읽고 " & cstr(write_cnt) & " 건이 정상적으로 처리되었습니다."
End If

rs.close : Set rs = Nothing
cn.close : Set cn = Nothing

DBConn.Close() : Set DBConn = Nothing

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	location.replace('/service/as_acpt_statics_list.asp?slip_month="&slip_month&"');"
Response.write "</script>"
Response.End
%>