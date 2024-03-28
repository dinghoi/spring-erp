<%
'/*****************************************************
'   작성자     : 조형렬 (lyoul@k-net.or.kr)
'   최초작성일 : 2001.12.31
'   최종수정일 : 2002.03.05
'   파  일     : cls_boardpage2.asp (Ver 2.0)
'   설  명     : 게시판 페이징 클래스
'******************************************************/


	Class LBoardPage2
		Private curr_page
		Private page_size
		Private page_count
		Private record_count
		Private link_url
		Private page_var_name
		Private left_img
		Private right_img
		Private pleft_img
		Private pright_img
		Private left_img_sep
		Private right_img_sep
		Private img_is_text
		Private page_link_count
		Private link_sep
		Private link_prefix
		Private link_suffix
		Private nav_pprefix
		Private nav_prefix
		Private nav_suffix
		Private nav_psuffix
		Private curr_page_prefix
		Private curr_page_suffix

		Private main_sql
		Private count_sql
		Private page_sql

		Private obj_rs

		Private rs_cursor_type
		Private rs_lock_type
		Private rs_options

		Private obj_conn

		Private reset_data
		Private parse_query

		Private main_sql_retrieve
		Private main_sql_order
		Private main_sql_condition

		Private Sub Class_Initialize ()
			curr_page = 1
			page_size = 10
			page_count = 1
			record_count = 0
			link_url = ""
			page_var_name = ""
			left_img = ""
			right_img = ""
			pleft_img = ""
			pright_img = ""
			left_img_sep = ""
			right_img_sep = ""
			img_is_text = False
			page_link_count = 9
			link_sep = "|"
			link_prefix = ""
			link_suffix = ""
			nav_pprefix = ""
			nav_prefix = ""
			nav_suffix = ""
			nav_psuffix = ""
			curr_page_prefix = ""
			curr_page_suffix = ""

			main_sql = ""
			count_sql = ""
			page_sql = ""

			Set obj_rs = Server.CreateObject ("ADODB.RecordSet")

			rs_cursor_type = adOpenStatic
			rs_lock_type = adLockReadOnly
			rs_options = adCmdText

			Set obj_conn = Nothing

			reset_data = True
			parse_query = True

			main_sql_retrieve = ""
			main_sql_order = ""
			main_sql_condition = ""
		End Sub


		Public Property Let CurrPage (ByVal new_cpage)
			if new_cpage = "" then
				new_cpage = 1
			End if

			new_cpage = CInt (new_cpage)

			if new_cpage = curr_page then
				Exit Property
			End if

			if new_cpage < 1 then
				new_cpage = 1
			End if

			curr_page = new_cpage

			reset_data = True
		End Property

		Public Property Let PageSize (ByVal new_psize)
			if new_psize = "" then
				new_psize = 1
			End if

			new_psize = CInt (new_psize)

			if new_psize = page_size then
				Exit Property
			End if

			if new_psize < 1 then
				new_psize = 1
			End if

			page_size = new_psize

			reset_data = True
		End Property

		Public Property Let LinkURL (ByVal url)
			link_url = url
		End Property

		Public Property Let PageVarName (ByVal name)
			page_var_name = name
		End Property

		Public Property Let LeftImg (ByVal img)
			left_img = img
		End Property

		Public Property Let RightImg (ByVal img)
			right_img = img
		End Property

		Public Property Let PLeftImg (ByVal img)
			pleft_img = img
		End Property

		Public Property Let PRightImg (ByVal img)
			pright_img = img
		End Property

		Public Property Let LeftImgSep (ByVal sep)
			left_img_sep = sep
		End Property

		Public Property Let RightImgSep (ByVal sep)
			right_img_sep = sep
		End Property

		Public Property Let ImgIsText (ByVal bool)
			img_is_text = bool
		End Property

		Public Property Let PageLinkCount (ByVal count)
			if count < 3 then
				count = 3
			End if

			page_link_count = count

			if (count Mod 2) = 0 then
'--				page_link_count = page_link_count + 1
			End if
		End Property

		Public Property Let Separator (ByVal sep)
			link_sep = sep
		End Property

		Public Property Let LinkPrefix (ByVal prefix)
			link_prefix = prefix
		End Property

		Public Property Let LinkSuffix (ByVal suffix)
			link_suffix = suffix
		End Property

		Public Property Let NavPPrefix (ByVal pprefix)
			nav_pprefix = pprefix
		End Property

		Public Property Let NavPrefix (ByVal prefix)
			nav_prefix = prefix
		End Property

		Public Property Let NavSuffix (ByVal suffix)
			nav_suffix = suffix
		End Property

		Public Property Let NavPSuffix (ByVal psuffix)
			nav_psuffix = psuffix
		End Property

		Public Property Let CPagePrefix (ByVal prefix)
			curr_page_prefix = prefix
		End Property

		Public Property Let CPageSuffix (ByVal suffix)
			curr_page_suffix = suffix
		End Property


		Public Property Let Sql (ByVal user_sql)
			if main_sql = user_sql then
				Exit Property
			End if

			main_sql = Trim(user_sql)
			parse_query = True
			reset_data = True
		End Property

		Public Property Let RsCursorType (ByVal new_cursor_type)
			rs_cursor_type = new_cursor_type
		End Property

		Public Property Let RsLockType (ByVal new_lock_type)
			rs_lock_type = new_lock_type
		End Property

		Public Property Let RsOptions (ByVal new_options)
			rs_options = new_options
		End Property

		Public Property Let Conn (ByRef opened_conn)
			Set obj_conn = opened_conn
		End Property


		Private Sub ParseQuery ()
			Dim		select_end, from_start, order_start
			Dim		lcase_sql

			if Not parse_query then
				Exit Sub
			End if

			lcase_sql = LCase (main_sql)

			select_end = InStr (lcase_sql, "select") + 6
			from_start = InStrRev (lcase_sql, " from ")
			order_start = InStrRev (lcase_sql, " order ")
			If Cint(order_start) <= 0 Then
				order_start = len(main_sql) + 1
			End If

			main_sql_condition = Mid (main_sql, from_start, order_start - from_start)

			parse_query = False
		End Sub

		Private Function AdjustData ()
			if Not reset_data then
				AdjustData = True
				Exit Function
			End if

			if main_sql = "" or Not IsObject (obj_conn) then
				AdjustData = False
				Exit Function
			End if

			Call ParseQuery ()


			' 총 레코드 수 얻기
			Set obj_rs = Nothing
			Set obj_rs = Server.CreateObject ("ADODB.RecordSet")

			count_sql = "select count(0) from (select 0 " & main_sql_condition & ")"

			obj_rs.Open count_sql, obj_conn, adOpenStatic, adLockReadOnly, adCmdText
			record_count = cLng (obj_rs (0))
			obj_rs.Close ()

			' 페이지 수 구하기
			if record_count = 0 then
				page_count = 1
			else
				page_count = record_count \ page_size
				if record_count mod page_size <> 0 then
					page_count = page_count + 1
				End if
			End if

			' 현재 페이지 조정
			if curr_page > page_count then
				curr_page = page_count
			End if

			page_sql = "select * from ( " & main_sql & ") " & _
				"where ROWNUM <= " & (page_size * curr_page)

			obj_rs.Open page_sql, obj_conn, rs_cursor_type, rs_lock_type, rs_options
			if Not obj_rs.EOF then
				Call obj_rs.Move ((curr_page - 1) * page_size)
			End if

			AdjustData = True
			reset_data = False
		End Function


		Public Property Get CurrPage ()
			if Not AdjustData () then
				CurrPage = -1
				Exit Property
			End if

			CurrPage = curr_page
		End Property

		Public Property Get PageSize ()
			PageSize = page_size
		End Property

		Public Property Get PageCount ()
			if Not AdjustData () then
				PageCount = -1
				Exit Property
			End if

			PageCount = page_count
		End Property

		Public Property Get RecordCount ()
			if Not AdjustData () then
				RecordCount = -1
				Exit Property
			End if

			RecordCount = record_count
		End Property

		Public Default Property Get Rs (ByVal key)
			Call AdjustData ()
			Rs = obj_rs (key)
		End Property

		Public Property Get BOF ()
			Call AdjustData ()
			BOF = obj_rs.BOF
		End Property

		Public Property Get EOF ()
			Call AdjustData ()
			EOF = obj_rs.EOF
		End Property

		Public Sub Move (ByVal offset)
			Call AdjustData ()
			Call obj_rs.Move (offset)
		End Sub

		Public Sub MoveNext ()
			Call AdjustData ()
			Call obj_rs.MoveNext ()
		End Sub

		Public Sub MovePrevious ()
			Call AdjustData ()
			Call obj_rs.MovePrevious ()
		End Sub

		Public Sub MoveFirst ()
			Call AdjustData ()
			Call obj_rs.MoveFirst ()
		End Sub

		Public Sub MoveLast ()
			Call AdjustData ()
			Call obj_rs.MoveLast ()
		End Sub


		Public Property Get LinkURL ()
			LinkURL = link_url
		End Property

		Public Property Get PageVarName ()
			PageVarName = page_var_name
		End Property

		Public Property Get LeftImg ()
			LeftImg = left_img
		End Property

		Public Property Get RightImg ()
			RightImg = right_img
		End Property

		Public Property Get PLeftImg ()
			PLeftImg = pleft_img
		End Property

		Public Property Get PRightImg ()
			PRightImg = pright_img
		End Property

		Public Property Get LeftImgSep ()
			LeftImgSep = left_img_sep
		End Property

		Public Property Get RightImgSep ()
			RightImgSep = right_img_sep
		End Property

		Public Property Get BeginNum ()
			BeginNum = record_count - (page_size * (curr_page - 1)) + 1
		End Property

		Public Property Get BeginNumRev ()
			BeginNumRev = page_size * (curr_page - 1)
		End Property

		Public Property Get ImgIsText ()
			ImgIsText = img_is_text
		End Property

		Public Property Get PageLinkCount ()
			PageLinkCount = page_link_count
		End Property

		Public Property Get Separator ()
			Separator = link_sep
		End Property

		Public Property Get LinkPrefix ()
			LinkPrefix = link_prefix
		End Property

		Public Property Get LinkSuffix ()
			LinkSuffix = link_suffix
		End Property

		Public Property Get NavPPrefix ()
			NavPPrefix = nav_pprefix
		End Property

		Public Property Get NavPrefix ()
			NavPrefix = nav_prefix
		End Property

		Public Property Get NavSuffix ()
			NavSuffix = nav_suffix
		End Property

		Public Property Get NavPSuffix ()
			NavPSuffix = nav_psuffix
		End Property

		Public Property Get CPagePrefix ()
			CPagePrefix = curr_page_prefix
		End Property

		Public Property Get CPageSuffix ()
			CPageSuffix = curr_page_suffix
		End Property

		Public Property Get Sql ()
			Sql = main_sql
		End Property

		Public Property Get CountSql ()
			Call AdjustData ()
			CountSql = count_sql
		End Property

		Public Property Get PageSql ()
			Call AdjustData ()
			PageSql = page_sql
		End Property

		Public Property Get RsCursorType ()
			RsCursorType = rs_cursor_type
		End Property

		Public Property Get RsLockType ()
			RsLockType = rs_lock_type
		End Property

		Public Property Get RsOptions ()
			RsOptions = rs_options
		End Property

		Public Property Get Conn ()
			Set conn = obj_conn
		End Property



		Private Sub DrawPageImg (ByVal arrow_img)
			if arrow_img = "" then
				Exit Sub
			End if

			if img_is_text then
				Response.Write arrow_img
			else
				Response.Write "<img src=""" & arrow_img & """ align=absmiddle border=0>"
			End if
		End Sub


		Public Sub Draw ()
			Dim			actual_link
			Dim			first_page	'-- page list 에서 << 버튼 눌렀을때 나오는 페이지
			Dim			start_page	'-- page list 처음으로 나오는 페이지 번호
			Dim			end_page		'-- page list 마지막으로 나
			Dim 		final_page	'-- page list 에서 >> 버튼 눌렀을때 나오는 페이지오는 페이지 번호
			Dim			i

			if InStr (link_url, "?") = 0 then
				actual_link = link_url & "?" & page_var_name & "="
			else
				actual_link = link_url & "&" & page_var_name & "="
			End if

			select case (curr_page mod page_link_count)	'-- 처음 페이지 번호
			case 0
				start_page = ((int(curr_page/page_link_count) - 1) * page_link_count) + 1
			case else
				start_page = (int(curr_page/page_link_count) * page_link_count) + 1
			end select

			end_page = start_page + page_link_count - 1		'-- 마지막 페이지 번호
			if end_page > page_count then
				end_page = page_count
			end if

			first_page = start_page - page_link_count
			If first_page <= 0 Then
				first_page = 1
			End If
			final_page = end_page + 1
			If final_page > page_count Then
				final_page = page_count
			End If


			if curr_page <> 1 then
				if pleft_img <> "" then
					Response.Write ("<a href=""" & actual_link & first_page & """>")
					Call DrawPageImg (pleft_img)
					Response.Write ("</a>")
					Response.Write (nav_pprefix)
				End if
				Response.Write (left_img_sep)
				if left_img <> "" then
					Response.Write ("<a href=""" & actual_link & (curr_page - 1) & """>")
					Call DrawPageImg (left_img)
					Response.Write ("</a>")
				End if
			else
				Call DrawPageImg (pleft_img)
				Response.Write (nav_pprefix)
				Response.Write (left_img_sep)
				Call DrawPageImg (left_img)
			End if

			Response.Write (nav_prefix)

			For i = start_page to end_page step 1
				if i <> start_page then
					Response.Write (link_sep)
				End if

				if curr_page = i then
					Response.Write " "
					Response.Write (curr_page_prefix)
					Response.Write (link_prefix)
					Response.Write (i)
					Response.Write (link_suffix)
					Response.Write (curr_page_suffix)
					Response.Write " "
				else
					Response.Write " "
					Response.Write ("<a href=""" & actual_link & i & """>")
					Response.Write (link_prefix)
					Response.Write (i)
					Response.Write (link_suffix)
					Response.Write ("</a>")
					Response.Write " "
				End if

				if i = page_count then
					Exit For
				End if
			Next

			Response.Write (nav_suffix)

			if curr_page <> page_count then
				if right_img <> "" then
					Response.Write ("<a href=""" & actual_link & (curr_page + 1) & """>")
					Call DrawPageImg (right_img)
					Response.Write ("</a>")
				End if
				Response.Write (right_img_sep)
				Response.Write (nav_psuffix)
				if pright_img <> "" then
					Response.Write ("<a href=""" & actual_link & final_page & """>")
					Call DrawPageImg (pright_img)
					Response.Write ("</a>")
				End if
			else
				Call DrawPageImg (right_img)
				Response.Write (right_img_sep)
				Response.Write (nav_psuffix)
				Call DrawPageImg (pright_img)
			End if
		End Sub

	End Class
%>
