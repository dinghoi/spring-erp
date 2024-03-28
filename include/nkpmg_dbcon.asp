<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<%
'Dim DBConnect
'실 서버
'DBConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=nkp;PWD=nkp2014;"

'개발 서버
'DBConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=211.43.210.66;DATABASE=nkp_dev;UID=nkp_dev;PWD=zpdldnjs!@3;"

If Request.Cookies("nkpmg_user")("coo_user_id") = "" Then
	Response.Write "<script type='text/javascript'>"
'	Response.Write "	alert('로그인이 필요합니다.');"
	Response.Write "	location.replace('warning.asp');"
	Response.Write "</script>"
End If

'//2016-09-07 replace xss string
Function replaceXSS(txt)
	If isNull(txt) = false Then
		txt = trim(replace(txt,"<","&lt;"))
		txt = trim(replace(txt,">","&gt;"))
		txt = trim(replace(txt,"""","&#34;"))
		txt = trim(replace(txt,",","&#39;"))
		txt = trim(replace(txt,"-","&#45"))
	Else
		txt = ""
	End If

	replaceXSS = txt
End Function

'//
Function toString(str1, str2)
	If IsNull(str1) Or str1="" Then
		str1 = ""
		If Not IsNull(str2) And str2<>"" Then
			str1 = str2
		End If
	End If

	toString = str1
End Function

'//2016-09-08 recordset to dictionary
Public Function getRsToDic(rs)
	Dim rsRow				'결과 레코드셋 2차원 배열
	Dim rsCount				'레코드셋 rows
	Dim resultList			'리턴할 리스트
	Dim dic					'레코드 row dictionary
	Dim i, j
	Dim name, value

	Set resultList = Server.CreateObject("scripting.dictionary")

	if Not rs.EOF Then
		rsRow = rs.GetRows()
		rsCount = ubound(rsRow,2)

		for i = 0 to rsCount
			Set dic = Server.CreateObject("scripting.dictionary")

			for j = 0 to rs.fields.count - 1
				name = rs.fields(j).name
				if isnull(rsRow(j, i)) Then
					value = ""
				Else
					value = CStr(rsRow(j, i))
				End If

				dic.add name, value
			next

			resultList.add i, dic
		next
	end if

	set getRsToDic = resultList
End Function
%>
