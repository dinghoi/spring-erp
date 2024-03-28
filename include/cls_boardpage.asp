<%
'/*****************************************************
'   작성자     : 조형렬 (lyoul@k-net.or.kr)
'   최초작성일 : 2001.12.31
'   최종수정일 : 2002.02.18
'   파  일     : cls_boardpage.asp
'   설  명     : 게시판 페이징 클래스
'******************************************************/


	Class LBoardPage
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
		Private nav_prefix
		Private nav_suffix
		Private curr_page_prefix
		Private curr_page_suffix

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
			left_img_sep = ""&"&nbsp;&nbsp;"
			right_img_sep = ""&"&nbsp;&nbsp;"
			img_is_text = False
			page_link_count = 9
			link_sep = "|"
			link_prefix = ""
			link_suffix = ""
			nav_prefix = ""
			nav_suffix = ""
			curr_page_prefix = ""
			curr_page_suffix = ""
		End Sub

		Private Function GetCurrPage ()
			if curr_page > page_count then
				curr_page = page_count
				GetCurrPage = page_count
			else
				GetCurrPage = curr_page
			End if
		End Function


		Public Property Let CurrPage (ByVal new_cpage)
			if new_cpage = "" then
				curr_page = 1
				Exit Property
			End if

			new_cpage = CInt (new_cpage)

			if new_cpage < 1 then
				new_cpage = 1
			End if

			curr_page = new_cpage
		End Property

		Public Property Let PageSize (ByVal new_psize)
			if new_psize = "" then
				page_size = 1
				Exit Property
			End if

			new_psize = CInt (new_psize)

			if new_psize < 1 then
				new_psize = 1
			End if

			page_size = new_psize
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
			if page_link_count < 3 then
				page_link_count = 3
				Exit Property
			End if

			page_link_count = count

			if (count Mod 2) = 0 then
				page_link_count = page_link_count + 1
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

		Public Property Let NavPrefix (ByVal prefix)
			nav_prefix = prefix
		End Property

		Public Property Let NavSuffix (ByVal suffix)
			nav_suffix = suffix
		End Property

		Public Property Let CPagePrefix (ByVal prefix)
			curr_page_prefix = prefix
		End Property

		Public Property Let CPageSuffix (ByVal suffix)
			curr_page_suffix = suffix
		End Property

		Public Property Get CurrPage ()
			CurrPage = GetCurrPage ()
		End Property

		Public Property Get PageSize()
			PageSize = page_size
		End Property

		Public Property Get PageCount ()
			PageCount = page_count
		End Property

		Public Property Get RecordCount ()
			RecordCount = record_count
		End Property

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

		Public Property Get NavPrefix ()
			NavPrefix = nav_prefix
		End Property

		Public Property Get NavSuffix ()
			NavSuffix = nav_suffix
		End Property

		Public Property Get CPagePrefix ()
			CPagePrefix = curr_page_prefix
		End Property

		Public Property Get CPageSuffix ()
			CPageSuffix = curr_page_suffix
		End Property


		Public Function SetRs (ByRef rs)
			if rs.RecordCount = 0 then
				record_count = 0
				page_count = 1
				curr_page = 1
				Exit Function
			End if

			rs.PageSize = page_size
			record_count = rs.RecordCount
			page_count = rs.PageCount

			curr_page = GetCurrPage ()
			rs.AbsolutePage = curr_page

			SetRs = curr_page
		End Function


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
		
		Public Default Sub Draw ()
			Dim			actual_link
			Dim			first_page
			Dim			page_num
			Dim			i

			if InStr (link_url, "?") = 0 then
				actual_link = link_url & "?" & page_var_name & "="
			else
				actual_link = link_url & "&" & page_var_name & "="
			End if

			if page_count < curr_page + (page_link_count \ 2) then
				first_page = page_count - page_link_count + 1
			else
				first_page = curr_page - page_link_count \ 2
			End if

			if first_page <= 0 or curr_page <= page_link_count \ 2 then
				first_page = 1
			End if

			if curr_page <> 1 then
				if pleft_img <> "" then
					Response.Write ("<a href=""" & actual_link & (curr_page - page_link_count) & """>")
					Call DrawPageImg (pleft_img)
					Response.Write ("</a>")
				End if
				Response.Write (left_img_sep)
				if left_img <> "" then
					Response.Write ("<a href=""" & actual_link & (curr_page - 1) & """>")
					Call DrawPageImg (left_img)
					Response.Write ("</a>")
				End if
			else
				Call DrawPageImg (pleft_img)
				Response.Write (left_img_sep)
				Call DrawPageImg (left_img)
			End if

			Response.Write (nav_prefix)

			For i = 0 to page_link_count - 1 step 1
				if i <> 0 then
					Response.Write (link_sep)
				End if

				page_num = first_page + i

				if curr_page = page_num then
					Response.Write " "
					Response.Write (curr_page_prefix)
					Response.Write (link_prefix)
					Response.Write (page_num)
					Response.Write (link_suffix)
					Response.Write (curr_page_suffix)
					Response.Write " "
				else
					Response.Write " "
					Response.Write ("<a href=""" & actual_link & page_num & """>")
					Response.Write (link_prefix)
					Response.Write (page_num)
					Response.Write (link_suffix)
					Response.Write ("</a>")
					Response.Write " "
				End if

				if page_num = page_count then
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
				if pright_img <> "" then
					Response.Write ("<a href=""" & actual_link & (curr_page + page_link_count) & """>")
					Call DrawPageImg (pright_img)
					Response.Write ("</a>")
				End if
			else
				Call DrawPageImg (right_img)
				Response.Write (right_img_sep)
				Call DrawPageImg (pright_img)
			End if
		End Sub
	End Class
%>
