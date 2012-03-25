<%
'#################################################################################
'## Snitz Forums 2000 v3.4.06
'#################################################################################
'## Copyright (C) 2000-06 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<%

if (Request.QueryString("TOPIC_ID") = "" or IsNumeric(Request.QueryString("TOPIC_ID")) = False) and Request.Form("Method_Type") <> "login" and Request.Form("Method_Type") <> "logout" then
	Response.Redirect "default.asp"
	Response.End
else
	Topic_ID = cLng(Request.QueryString("TOPIC_ID"))
end if 
Dim ArchiveView, ArchiveLink, CColor
if request("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
	ArchiveView = "true"
	ArchiveLink = "ARCHIVE=true&"
elseif request("ARCHIVE") <> "" then
	Response.Redirect "default.asp"
	Response.End
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
	ArchiveLink = ""
end if

%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->

<% 
Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"    function ChangePage(fnum){" & vbNewLine & _
		"    	if (fnum == 1) {" & vbNewLine & _
		"    		document.PageNum1.submit();" & vbNewLine & _
		"    	}" & vbNewLine & _
		"    	else {" & vbNewLine & _
		"    		document.PageNum2.submit();" & vbNewLine & _
		"    	}" & vbNewLine & _
		"    }" & vbNewLine & _
		"    </script>" & vbNewLine

mypage = request("whichpage")
if ((Trim(mypage) = "") or (IsNumeric(mypage) = False)) then mypage = 1
mypage = cLng(mypage)

if Request("SearchTerms") <> "" then
	SearchLink = "&SearchTerms=" & Request("SearchTerms")
else
	SearchLink = ""
end if

if strSignatures = "1" and strDSignatures = "1" then
	if ViewSig(MemberID) <> "0" then
		CanShowSignature = 1
	end if
end if

'## Forum_SQL - Get original topic and check for the Category, Forum or Topic Status and existence
strSql = "SELECT M.M_NAME, M.M_RECEIVE_EMAIL, M.M_AIM, M.M_ICQ, M.M_MSN, M.M_YAHOO" & _
	", M.M_TITLE, M.M_HOMEPAGE, M.MEMBER_ID, M.M_LEVEL, M.M_POSTS, M.M_COUNTRY" & _
	", T.T_DATE, T.T_SUBJECT, T.T_AUTHOR, T.TOPIC_ID, T.T_STATUS, T.T_LAST_EDIT" & _
	", T.T_LAST_EDITBY, T.T_LAST_POST, T.T_SIG, T.T_REPLIES" & _
	", C.CAT_STATUS, C.CAT_ID, C.CAT_NAME, C.CAT_SUBSCRIPTION, C.CAT_MODERATION" & _
	", F.F_STATUS, F.FORUM_ID, F.F_SUBSCRIPTION, F.F_SUBJECT, F.F_MODERATION, T.T_MESSAGE"
if CanShowSignature = 1 then
	strSql = strSql & ", M.M_SIG"
end if
strSql = strSql & " FROM " & strActivePrefix & "TOPICS T, " & strTablePrefix & "FORUM F, " & _ 
	strTablePrefix & "CATEGORY C, " & strMemberTablePrefix & "MEMBERS M " & _
	" WHERE T.TOPIC_ID = " & Topic_ID & _
	" AND F.FORUM_ID = T.FORUM_ID " & _
	" AND C.CAT_ID = T.CAT_ID " & _
	" AND M.MEMBER_ID = T.T_AUTHOR "

set rsTopic = Server.CreateObject("ADODB.Recordset")
rsTopic.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rsTopic.EOF then
	recTopicCount = ""
else
	recTopicCount = 1
	Member_Name = rsTopic("M_NAME")
	Member_ReceiveMail = rsTopic("M_RECEIVE_EMAIL")
	Member_AIM = rsTopic("M_AIM")
	Member_ICQ = rsTopic("M_ICQ")
	Member_MSN = rsTopic("M_MSN")
	Member_YAHOO = rsTopic("M_YAHOO")
	Member_Title = rsTopic("M_TITLE")
	Member_Homepage = rsTopic("M_HOMEPAGE")
	TMember_ID = rsTopic("MEMBER_ID")
	Member_Level = rsTopic("M_LEVEL")
	Member_Posts = rsTopic("M_POSTS")
	Member_Country = rsTopic("M_COUNTRY")
	Topic_Date = rsTopic("T_DATE")
	Topic_Subject = rsTopic("T_SUBJECT")
	Topic_Author = rsTopic("T_AUTHOR")
	TopicID = rsTopic("TOPIC_ID")
	Topic_Status = rsTopic("T_STATUS")
	Topic_LastEdit = rsTopic("T_LAST_EDIT")
	Topic_LastEditby = rsTopic("T_LAST_EDITBY")
	Topic_LastPost = rsTopic("T_LAST_POST")
	Topic_Sig = rsTopic("T_SIG")
	Topic_Replies = rsTopic("T_REPLIES")
	Cat_Status = rsTopic("CAT_STATUS")
	Cat_ID = rsTopic("CAT_ID")
	Cat_Name = rsTopic("CAT_NAME")
	Cat_Subscription = rsTopic("CAT_SUBSCRIPTION")
	Cat_Moderation = rsTopic("CAT_MODERATION")
	Forum_Status = rsTopic("F_STATUS")
	Forum_ID = rsTopic("FORUM_ID")
	Forum_Subject = rsTopic("F_SUBJECT")
	Forum_Subscription = rsTopic("F_SUBSCRIPTION")
	Forum_Moderation = rsTopic("F_MODERATION")
	Topic_Message = rsTopic("T_MESSAGE")
	if CanShowSignature = 1 then
		Topic_MemberSig = trim(rsTopic("M_SIG"))
	end if
end if

rsTopic.close
set rsTopic = nothing

if recTopicCount = "" then
	if ArchiveView <> "true" then
		Response.Redirect("topic.asp?ARCHIVE=true&" & ChkString(Request.QueryString,"sqlstring"))
	else
		Response.Redirect("default.asp")
	end if
end if

if mLev = 4 then
	AdminAllowed = 1
	ForumChkSkipAllowed = 1
elseif mLev = 3 then
	if chkForumModerator(Forum_ID, chkString(strDBNTUserName,"decode")) = "1" then
		AdminAllowed = 1
		ForumChkSkipAllowed = 1
	else
		if lcase(strNoCookies) = "1" then
			AdminAllowed = 1
			ForumChkSkipAllowed = 0
		else
			AdminAllowed = 0
			ForumChkSkipAllowed = 0
		end if
	end if
elseif lcase(strNoCookies) = "1" then
 	AdminAllowed = 1
	ForumChkSkipAllowed = 0
else   
 	AdminAllowed = 0
	ForumChkSkipAllowed = 0
end if

if strPrivateForums = "1" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") and ForumChkSkipAllowed = 0 then
	result = ChkForumAccess(Forum_ID, MemberID, true)
end if

if strModeration > 0 and Cat_Moderation > 0 and Forum_Moderation > 0 and AdminAllowed = 0 then
        Moderation = "Y"
else
        Moderation = "N"
end if

if mypage = -1 then
	strSql = "SELECT REPLY_ID FROM " & strActivePrefix & "REPLY WHERE TOPIC_ID = " & Topic_ID & " "
	if AdminAllowed = 0 then
		strSql = strSql & " AND (R_STATUS < "
		if Moderation = "Y" then
			strSql = strSql & "2 "
		else
			strSql = strSql & "3 "
		end if
		strSql = strSql & "OR R_AUTHOR = " & MemberID & ") "
	end if
	strSql = strSql & "ORDER BY R_DATE ASC "

	set rsReplies = Server.CreateObject("ADODB.Recordset")
	if strDBType = "mysql" then
		rsReplies.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	else
		rsReplies.open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	end if
	
	if not rsReplies.EOF then
		arrReplyData = rsReplies.GetRows(adGetRowsRest)
		iReplyCount = UBound(arrReplyData, 2)
		
		if Request.Querystring("REPLY_ID") <> "" and IsNumeric(Request.Querystring("REPLY_ID")) then
			LastPostReplyID = cLng(Request.Querystring("REPLY_ID"))
			for iReply = 0 to iReplyCount
				intReplyID = arrReplyData(0, iReply)
				if LastPostReplyID = intReplyID then
					intPageNumber = ((iReply+1)/strPageSize)
					exit for
				end if
			next
		else
			LastPostReplyID = cLng(arrReplyData(0, iReplyCount))
			intPageNumber = ((iReplyCount+1)/strPageSize)
		end if
		if intPageNumber > cLng(intPageNumber) then
			intPageNumber = cLng(intPageNumber) + 1
		end if
		strwhichpage = "whichpage=" & intPageNumber & "&"
	else
		strwhichpage = ""
	end if
	
	rsReplies.close
	set rsReplies = nothing
	my_Conn.close
	set my_Conn = nothing
	
	Response.Redirect "topic.asp?" & ArchiveLink & strwhichpage & "TOPIC_ID=" & Topic_ID & SearchLink & "&#" & LastPostReplyID
	Response.End
end if

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
if MySubCount > 0 then
	strSubString = PullSubscriptions(0, 0, 0)
	strSubArray  = Split(strSubString,";")
	if uBound(strSubArray) < 0 then
		strBoardSubs = ""
		strCatSubs = ""
		strForumSubs = ""
		strTopicSubs = ""
	else
		strBoardSubs = strSubArray(0)
		strCatSubs = strSubArray(1)
		strForumSubs = strSubArray(2)
		strTopicSubs = strSubArray(3)
	end if
end If

if (Moderation = "Y" and Topic_Status > 1 and Topic_Author <> MemberID) then
	Response.write  "<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><br />Viewing of this Topic is not permitted until it has been moderated.<br />Please try again later</font></p>" & vbNewLine & _
			"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back</a></font></p><br />" & vbNewLine
	WriteFooter
	Response.end
else
	Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
			"    <!--" & vbNewLine & _
			"    function jumpTo(s) {if (s.selectedIndex != 0) location.href = s.options[s.selectedIndex].value;return 1;}" & vbNewLine & _
			"    // -->" & vbNewLine & _
			"    </script>" & vbNewLine

	'## Forum_SQL
	strSql = "SELECT M.M_NAME, M.M_RECEIVE_EMAIL, M.M_AIM, M.M_ICQ, M.M_MSN, M.M_YAHOO"
	strSql = strSql & ", M.M_TITLE, M.MEMBER_ID, M.M_HOMEPAGE, M.M_LEVEL, M.M_POSTS, M.M_COUNTRY"
	strSql = strSql & ", R.REPLY_ID, R.FORUM_ID, R.R_AUTHOR, R.TOPIC_ID, R.R_MESSAGE, R.R_LAST_EDIT"
	strSql = strSql & ", R.R_LAST_EDITBY, R.R_SIG, R.R_STATUS, R.R_DATE"
	if CanShowSignature = 1 then
		strSql = strSql & ", M.M_SIG"
	end if
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "REPLY R "
	strSql3 = " WHERE M.MEMBER_ID = R.R_AUTHOR "
	strSql3 = strSql3 & " AND R.TOPIC_ID = " & Topic_ID & " "
		' DEM --> if not a Moderator, all unapproved posts should not be viewed.
		if AdminAllowed = 0 then
			strSql3 = strSql3 & " AND (R.R_STATUS < "
			if Moderation = "Y" then
				' Ignore unapproved/rejected posts
				strSql3 = strSql3 & "2"
			else
				' Ignore any previously rejected topic
				strSql3 = strSql3 & "3"
			end if
		strSql3 = strSql3 & " OR R.R_AUTHOR = " & MemberID & ")"
	end if
	strSql4 = " ORDER BY R.R_DATE ASC"

	if strDBType = "mysql" then 'MySql specific code
		if mypage > 1 then
 			intOffset = cLng((mypage-1) * strPageSize)
			strSql5 = " LIMIT " & intOffset & ", " & strPageSize & " "
		end if

		'## Forum_SQL - Get the total pagecount 
		strSql1 = "SELECT COUNT(R.TOPIC_ID) AS REPLYCOUNT "
		
		set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
		iPageTotal = rsCount(0).value
		rsCount.close
		set rsCount = nothing

		if iPageTotal > 0 then
			maxpages = (iPageTotal  \ strPageSize )
			if iPageTotal mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
			if iPageTotal < (strPageSize + 1) then
				intGetRows = iPageTotal
			elseif (mypage * strPageSize) > iPageTotal then
				intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
			else
				intGetRows = strPageSize
			end if
		else
			iPageTotal = 0
			maxpages = 0
		end if
	
		if iPageTotal > 0 then
			set rsReplies = Server.CreateObject("ADODB.Recordset")
			rsReplies.Open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
				arrReplyData = rsReplies.GetRows(intGetRows)
				iReplyCount = UBound(arrReplyData, 2)
			rsReplies.Close
			set rsReplies = nothing
		else
			iReplyCount = ""
		end if
		
	else 'end MySql specific code
	
		set rsReplies = Server.CreateObject("ADODB.Recordset")
		rsReplies.cachesize = strPageSize
		rsReplies.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

			if not (rsReplies.EOF or rsReplies.BOF) then
				rsReplies.pagesize = strPageSize
				rsReplies.absolutepage = mypage '**
				maxpages = cLng(rsReplies.pagecount)
				if maxpages >= mypage then
					arrReplyData = rsReplies.GetRows(strPageSize)
					iReplyCount = UBound(arrReplyData, 2)
				else
					iReplyCount = ""
				end if
			else  '## No replies found in DB
				iReplyCount = ""
			end if

		rsReplies.Close
		set rsReplies = nothing
	end if

	Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td width=""50%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
			"          " & getCurrentIcon(strIconBar,"","align=""absmiddle""")
	if Cat_Status <> 0 then 
		Response.Write	getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""")
	else 
		Response.Write	getCurrentIcon(strIconFolderClosed,"","align=""absmiddle""")
	end if
	Response.Write	"&nbsp;<a href=""default.asp?CAT_ID=" & Cat_ID & """>" & ChkString(Cat_Name,"display") & "</a><br />" & vbNewLine & _
			"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""")
	if ArchiveView = "true" then
		Response.Write	getCurrentIcon(strIconFolderArchived,"","align=""absmiddle""")
	else
		if Forum_Status <> 0 and Cat_Status <> 0 then
			Response.Write	getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""")
		else
			Response.Write	getCurrentIcon(strIconFolderClosed,"","align=""absmiddle""")
		end if
	end if
	Response.Write	"&nbsp;<a href=""forum.asp?" & ArchiveLink & "FORUM_ID=" & Forum_ID & """>" & ChkString(Forum_Subject,"display") & "</a><br />" & vbNewLine
	if ArchiveView = "true" then
		Response.Write	"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""") & getCurrentIcon(strIconFolderArchived,"","align=""absmiddle""") & "&nbsp;"
	elseif Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
		Response.Write	"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""") & getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""") & "&nbsp;"
	else
		Response.Write	"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""") & getCurrentIcon(strIconFolderClosedTopic,"","align=""absmiddle""") & "&nbsp;"
	end if
	if Request.QueryString("SearchTerms") <> "" then
		Response.Write	SearchHiLite(ChkString(Topic_Subject,"title"))
	else
		Response.Write	ChkString(Topic_Subject,"title")
	end if
	Response.Write	"</font></td>" & vbNewLine & _
			"          <td align=""center"" width=""50%"">" & vbNewLine
	call PostingOptions()
	Response.Write	"</td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    </td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
	if maxpages > 1 then
		Response.Write	"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""1"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
		if mypage > 1 then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&whichpage=" & mypage-1 & SearchLink & """ title=""Goto the Previous page in this Topic""" & dWStatus("Goto the Previous page in this Topic") & ">Previous Page</a>")
		'if mypage > 1 then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage-1 & ";document.PageNum1.submit();"" title=""Goto the Previous page in this Topic""" & dWStatus("Goto the Previous page in this Topic") & ">Previous Page</a>")
		if mypage > 1 and mypage < maxpages then Response.Write(" | ")
		if mypage < maxpages then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&whichpage=" & mypage+1 & SearchLink & """ title=""Goto the Next page in this Topic""" & dWStatus("Goto the Next page in this Topic") & ">Next Page</a>")
		'if mypage < maxpages then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage+1 & ";document.PageNum1.submit();"" title=""Goto the Next page in this Topic""" & dWStatus("Goto the Next page in this Topic") & ">Next Page</a>")
		Response.Write	"</td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
	end if
	Response.Write	"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td>" & vbNewLine & _
  			"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ width=""" & strTopicWidthLeft & """"
	if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap")
	Response.Write	"><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Author</font></b></td>" & vbNewLine & _
			"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ width=""" & strTopicWidthRight & """"
	if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap")
	Response.Write	"><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>" & vbNewLine
	if strShowTopicNav = "1" then
		Call Topic_nav()
	else
		Response.Write("Topic")
	end if
	Response.Write	"</font></b></td>" & vbNewLine
	if (AdminAllowed = 1) then
		if maxpages > 1 then
			Call DropDownPaging(1)
			Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap>" & vbNewLine
			call AdminOptions()
			Response.Write	"                </td>" & vbNewLine
		else
			Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap>" & vbNewLine
			call AdminOptions()
			Response.Write	"                </td>" & vbNewLine
		end if
	else
		if maxpages > 1 then
			Call DropDownPaging(1)
		else
	        	Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></td>" & vbNewLine
	 	end if
	end if
	Response.Write	"              </tr>" & vbNewLine
 
	if mypage = 1 then 
		Call GetFirst() 
	end if
	
	'## Forum_SQL
	strSql = "UPDATE " & strActivePrefix & "TOPICS "
	strSql = strSql & " SET T_VIEW_COUNT = (T_VIEW_COUNT + 1) "
	strSql = strSql & " WHERE (TOPIC_ID = " & Topic_ID & ")"

	my_conn.Execute (strSql),,adCmdText + adExecuteNoRecords
 
	if iReplyCount = "" then  '## No replies found in DB
		' Nothing
	else
		intI = 0 
	
		rM_NAME = 0
		rM_RECEIVE_EMAIL = 1
		rM_AIM = 2
		rM_ICQ = 3
		rM_MSN = 4
		rM_YAHOO = 5
		rM_TITLE = 6
		rMEMBER_ID = 7
		rM_HOMEPAGE = 8
		rM_LEVEL = 9
		rM_POSTS = 10
		rM_COUNTRY = 11
		rREPLY_ID = 12
		rFORUM_ID = 13
		rR_AUTHOR = 14
		rTOPIC_ID = 15
		rR_MESSAGE = 16
		rR_LAST_EDIT = 17
		rR_LAST_EDITBY = 18
		rR_SIG = 19
		rR_STATUS = 20
		rR_DATE = 21
		if CanShowSignature = 1 then
			rM_SIG = 22
		end if
		
		for iForum = 0 to iReplyCount

			Reply_MemberName = arrReplyData(rM_NAME, iForum)
			Reply_MemberReceiveEmail = arrReplyData(rM_RECEIVE_EMAIL, iForum)
			Reply_MemberAIM = arrReplyData(rM_AIM, iForum)
			Reply_MemberICQ = arrReplyData(rM_ICQ, iForum)
			Reply_MemberMSN = arrReplyData(rM_MSN, iForum)
			Reply_MemberYAHOO = arrReplyData(rM_YAHOO, iForum)
			Reply_MemberTitle = arrReplyData(rM_TITLE, iForum)
			Reply_MemberID = arrReplyData(rMEMBER_ID, iForum)
			Reply_MemberHomepage = arrReplyData(rM_HOMEPAGE, iForum)
			Reply_MemberLevel = arrReplyData(rM_LEVEL, iForum)
			Reply_MemberPosts = arrReplyData(rM_POSTS, iForum)
			Reply_MemberCountry = arrReplyData(rM_COUNTRY, iForum)
			Reply_ReplyID = arrReplyData(rREPLY_ID, iForum)
			Reply_ForumID = arrReplyData(rFORUM_ID, iForum)
			Reply_Author = arrReplyData(rR_AUTHOR, iForum)
			Reply_TopicID = arrReplyData(rTOPIC_ID, iForum)
			Reply_Content = arrReplyData(rR_MESSAGE, iForum)
			Reply_LastEdit = arrReplyData(rR_LAST_EDIT, iForum)
			Reply_LastEditBy = arrReplyData(rR_LAST_EDITBY, iForum)
			Reply_Sig = arrReplyData(rR_SIG, iForum)
			Reply_Status = arrReplyData(rR_STATUS, iForum)
			Reply_Date = arrReplyData(rR_DATE, iForum)
			if CanShowSignature = 1 then
				Reply_MemberSig = trim(arrReplyData(rM_SIG, iForum))
			end if

			if intI = 0 then 
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if

			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgcolor=""" & CColor & """ valign=""top"" width=""" & strTopicWidthLeft & """"
			if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap")
			Response.Write	">" & vbNewLine & _
					"                <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><span class=""spnMessageText"">" & profileLink(ChkString(Reply_MemberName,"display"),Reply_Author) & "</span></b></font><br />" & vbNewLine
			if strShowRank = 1 or strShowRank = 3 then
				Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & ChkString(getMember_Level(Reply_MemberTitle, Reply_MemberLevel, Reply_MemberPosts),"display") & "</small></font><br />" & vbNewLine
			end if
			if strShowRank = 2 or strShowRank = 3 then
				Response.Write	"                " & getStar_Level(Reply_MemberLevel, Reply_MemberPosts) & "<br />" & vbNewLine
			end if
		 	Response.Write	"                </p>" & vbNewLine & _
					"                <p>" & vbNewLine
			if strCountry = "1" and trim(Reply_MemberCountry) <> "" then
				Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & Reply_MemberCountry & "</small></font><br />" & vbNewLine
			end if
			Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & Reply_MemberPosts & " Posts</small></font></p></td>" & vbNewLine & _
					"                <td bgcolor=""" & CColor & """ height=""100%"" width=""" & strTopicWidthRight & """"
			if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap")
			if (AdminAllowed = 1) and (maxpages > 1) then
				Response.Write	(" colspan=""3"" ")
			else
				Response.Write	(" colspan=""2"" ")
			end if
			Response.Write	"valign=""top""><a name=""" & Reply_ReplyID & """></a>" & vbNewLine & _
					"                  <table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
					"                    <tr>" & vbNewLine & _
					"                      <td valign=""top"">" & vbNewLine
			' DEM --> Start of Code altered for moderation
			if Reply_Status < 2 then
				Response.Write  "                      " & getCurrentIcon(strIconPosticon,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>Posted&nbsp;-&nbsp;" & ChkDate(Reply_Date, "&nbsp;:&nbsp;" ,true) & "</font>" & vbNewline
			elseif Reply_Status = 2 then
				Response.Write  "                      <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>NOT MODERATED!!!</font>" & vbNewline
			elseif Reply_Status = 3 then
				Response.Write  "                      " & getCurrentIcon(strIconPosticonHold,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>ON HOLD</font>" & vbNewline
			end if
			' DEM --> End of Code added for moderation.
			Response.Write	"                      &nbsp;" & profileLink(getCurrentIcon(strIconProfile,"Show Profile","align=""absmiddle"" hspace=""6"""),Reply_MemberID) & vbNewLine
			if mLev > 2 or Reply_MemberReceiveEmail = "1" then
				if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_mail.asp?id=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconEmail,"Email Poster","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if strHomepage = "1" then
				if Reply_MemberHomepage <> " " then
					Response.Write	"                      &nbsp;<a href=""" & Reply_MemberHomepage & """ target=""_blank"">" & getCurrentIcon(strIconHomepage,"Visit " & ChkString(Reply_MemberName,"display") & "'s Homepage","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if (AdminAllowed = 1 or Reply_MemberID = MemberID) then
				if (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) or (AdminAllowed = 1) then
					Response.Write	"                      &nbsp;<a href=""post.asp?" & ArchiveLink & "method=Edit&REPLY_ID=" & Reply_ReplyID & "&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconEditTopic,"Edit Reply","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if (strAIM = "1") then
				if Trim(Reply_MemberAIM) <> "" then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Reply_MemberName,"display") & " an AOL message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if strICQ = "1" then
				if Trim(Reply_MemberICQ) <> "" then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=ICQ&ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Reply_MemberName,"display") & " an ICQ Message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if (strMSN = "1") then
				if Trim(Reply_MemberMSN) <> "" then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Reply_MemberName,"display") & "'s MSN Messenger address","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if strYAHOO = "1" then
				if Trim(Reply_MemberYAHOO) <> "" then
					Response.Write	"                      &nbsp;<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Reply_MemberYAHOO, "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Reply_MemberName,"display") & " a Yahoo! Message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			phSQL = "SELECT COUNT(*) AS EditCount FROM " & strTablePrefix & "POST_HISTORY WHERE R_ID = " & Reply_ReplyID
			Set phRS = Server.CreateObject("ADODB.RecordSet")
			phRS.Open phSQL, my_Conn
			EditCount = phRS("EditCount")
			phRS.Close
			Set phRS = Nothing
			if EditCount <> 0 then
				if mlev = 4 or mlev = 3 then
					Response.Write	"                      &nbsp;<a href=""post_history.asp?" & ArchiveLink & "R_ID=" & Reply_ReplyID & """>" & getCurrentIcon(strIconPostHistory,"View Reply History","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if ((Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status = 1) or (AdminAllowed = 1 and Topic_Status <= 1)) and ArchiveView = "" then
				Response.Write	"                      &nbsp;<a href=""post.asp?" & ArchiveLink & "method=ReplyQuote&REPLY_ID=" & Reply_ReplyID & "&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply with Quote","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
			end if
			if (strIPLogging = "1") then
				if (AdminAllowed = 1) then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_viewip.asp?" & ArchiveLink & "mode=getIP&REPLY_ID=" & Reply_ReplyID & "&FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconIP,"View user's IP address","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
			end if
			if (AdminAllowed = 1 or Reply_MemberID = MemberID) then
				if (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) or (AdminAllowed = 1) then
					Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Reply&REPLY_ID=" & Reply_ReplyID & "&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconDeleteReply,"Delete Reply","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
				end if
		                ' DEM --> Start of Code added for Full Moderation
				if (AdminAllowed = 1 and Reply_Status > 1) then
					ReplyString = "REPLY_ID=" & Reply_ReplyID & "&CAT_ID=" & Cat_ID & "&FORUM_ID=" & Forum_ID & "&TOPIC_ID=" & Topic_ID
					Response.Write "                      &nbsp;<a href=""JavaScript:openWindow('pop_moderate.asp?" & ReplyString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Reply","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewline
				end if
				' DEM --> End of Code added for Full Moderation
			end if
			Response.Write	"                      <hr noshade size=""" & strFooterFontSize & """></td>" & vbNewLine & _
					"                    </tr>" & vbNewLine & _
					"                    <tr>" & vbNewLine & _
					"                      <td valign=""top"" height=""100%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"" id=""msg"">"
			if Request.QueryString("SearchTerms") <> "" then
				Response.Write	SearchHiLite(formatStr(Reply_Content))
			else
				Response.Write	formatStr(Reply_Content)
			end if
			Response.Write	"</span id=""msg""></font></td>" & vbNewLine & _
					"                    </tr>" & vbNewLine
			if CanShowSignature = 1 and Reply_Sig = 1 and Reply_MemberSig <> "" then
				Response.Write	"                    <tr>" & vbNewLine & _
						"                      <td valign=""bottom""><hr noshade size=""" & strFooterFontSize & """><font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><span class=""spnMessageText"">" & formatStr(Reply_MemberSig) & "</span></font></td>" & vbNewLine & _
						"                    </tr>" & vbNewLine
			end if
			if strEditedByDate = "1" and Reply_LastEditBy <> "" then
				if Reply_LastEditBy <> Reply_Author then 
					Reply_LastEditByName = getMemberName(Reply_LastEditBy)
				else
					Reply_LastEditByName = chkString(Reply_MemberName,"display")
				end if
				Response.Write	"                    <tr>" & vbNewLine & _
						"                      <td valign=""bottom""><hr noshade size=""" & strFooterFontSize & """ color=""" & CColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & _
						"Edited by - " & Reply_LastEditByName & " on " & chkDate(Reply_LastEdit, " " ,true) & "</font></td>" & vbNewLine & _
						"                    </tr>" & vbNewLine
			end if
			Response.Write	"                    <tr>" & vbNewLine & _
					"                      <td valign=""bottom"" align=""right"" height=""20""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go to Top of Page","align=""right""") & "</a></td>" & vbNewLine & _
					"                    </tr>" & vbNewLine & _
					"                  </table>" & vbNewLine & _
					"                </td>" & vbNewLine & _
					"              </tr>" & vbNewLine
			intI  = intI + 1
			if intI = 2 then 
				intI = 0
			end if
		next
	end if
	Response.Write	"              <tr>" & vbNewLine
	if maxpages > 1 then
		Call DropDownPaging(2)
	else
		Response.Write	"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ width=""" & strTopicWidthLeft & """"
		if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap")
		Response.Write	"><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&nbsp;</font></b></td>" & vbNewLine
	end if
	Response.Write	"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ width=""" & strTopicWidthRight & """"
	if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap")
	'if maxpages > 1 and (AdminAllowed = 1) then Response.Write(" colspan=""2""")
	Response.Write	"><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>" & vbNewLine
	if strShowTopicNav = "1" then
		Call Topic_nav()
	else
		Response.Write("Topic")
	end if
	Response.Write	"</font></b></td>" & vbNewLine
	if (AdminAllowed = 1) then
		if maxpages > 1 then
	        	Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></td>" & vbNewLine
		end if
		Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap>" & vbNewLine
		call AdminOptions()
		Response.Write	"</td>" & vbNewLine
	else
        	Response.Write	"                <td align=""right"" bgcolor=""" & strHeadCellColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></td>" & vbNewLine
	end if
	Response.Write	"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    </td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
	if maxpages > 1 then
		Response.Write	"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""1"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td align=""left"" valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
		if mypage > 1 then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&whichpage=" & mypage-1 & SearchLink & """ title=""Goto the Previous page in this Topic""" & dWStatus("Goto the Previous page in this Topic") & ">Previous Page</a>")
		'if mypage > 1 then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage-1 & ";document.PageNum1.submit();"" title=""Goto the Previous page in this Topic""" & dWStatus("Goto the Previous page in this Topic") & ">Previous Page</a>")
		if mypage > 1 and mypage < maxpages then Response.Write(" | ")
		if mypage < maxpages then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&whichpage=" & mypage+1 & SearchLink & """ title=""Goto the Next page in this Topic""" & dWStatus("Goto the Next page in this Topic") & ">Next Page</a>")
		'if mypage < maxpages then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage+1 & ";document.PageNum1.submit();"" title=""Goto the Next page in this Topic""" & dWStatus("Goto the Next page in this Topic") & ">Next Page</a>")
		Response.Write	"</td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
	end if
	Response.Write	"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td>" & vbNewLine & _
			"      <table width=""100%"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""center"" valign=""top"" width=""50%"">" & vbNewLine
	Call PostingOptions()
	Response.Write	"</td>" & vbNewLine & _
			"          <td align=""right"" valign=""top"" width=""50%"" nowrap>" & vbNewLine
%>
	<!--#INCLUDE FILE="inc_jump_to.asp" -->
<%
	Response.Write	"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine

	if strShowQuickReply = "1" and strDBNTUserName <> "" and ((Cat_Status = 1) and (Forum_Status = 1) and (Topic_Status = 1)) and ArchiveView = "" then
		call QuickReply()
	end if

        WriteFooter
end if

sub GetFirst()
	CColor = strForumFirstCellColor
	Response.Write	"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strForumFirstCellColor & """ valign=""top"" width=""" & strTopicWidthLeft & """"
	if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap")
	Response.Write	">" & vbNewLine & _
			"                <p><font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><span class=""spnMessageText"">" & profileLink(ChkString(Member_Name,"display"),TMember_ID) & "</span></b></font><br />" & vbNewLine
	if strShowRank = 1 or strShowRank = 3 then
 		Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & ChkString(getMember_Level(Member_Title, Member_Level, Member_Posts),"display") & "</small></font><br />" & vbNewLine
	end if
	if strShowRank = 2 or strShowRank = 3 then
		Response.Write	"                " & getStar_Level(Member_Level, Member_Posts) & "<br />" & vbNewLine
	end if
 	Response.Write	"                </p>" & vbNewLine & _
			"                <p>" & vbNewLine
	if strCountry = "1" and trim(Member_Country) <> "" then
		Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & Member_Country & "</small></font><br />" & vbNewLine
	end if
	Response.Write	"                <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><small>" & Member_Posts & " Posts</small></font></p></td>" & vbNewLine & _
			"                <td bgcolor=""" & strForumFirstCellColor & """ width=""" & strTopicWidthRight & """"
	if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap")
	if (AdminAllowed = 1) and (maxpages > 1) then
		Response.Write	(" colspan=""3"" ")
	else
		Response.Write	(" colspan=""2"" ")
	end if
	Response.Write	"valign=""top"">" & vbNewLine & _
			"                  <table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
			"                    <tr>" & vbNewLine & _
			"                      <td valign=""top"">" & vbNewLine
	if Topic_Status < 2 then
		Response.Write  "                      " & getCurrentIcon(strIconPosticon,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>Posted&nbsp;-&nbsp;" & ChkDate(Topic_Date, "&nbsp;:&nbsp;" ,true) & "</font>" & vbNewline
	elseif Topic_Status = 2 then
		Response.Write  "                      <font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>NOT MODERATED!!!</font>" & vbNewline
	elseif Topic_Status = 3 then
		Response.Write  "                      " & getCurrentIcon(strIconPosticonHold,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>ON HOLD</font>" & vbNewline
	end if
	Response.Write	"                      &nbsp;" & profileLink(getCurrentIcon(strIconProfile,"Show Profile","align=""absmiddle"" hspace=""6"""),TMember_ID) & vbNewLine
	if mLev > 2 or Member_ReceiveMail = "1" then 
		if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then 
			Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_mail.asp?id=" & TMember_ID & "')"">" & getCurrentIcon(strIconEmail,"Email Poster","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (strHomepage = "1") then
		if Member_Homepage <> " " then
			Response.Write	"                      &nbsp;<a href=""" & Member_Homepage & """ target=""_blank"">" & getCurrentIcon(strIconHomepage,"Visit " & ChkString(Member_Name,"display") & "'s Homepage","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (AdminAllowed = 1 or TMember_ID = MemberID) then
		if ((Cat_Status <> 0) and (Forum_Status <> 0) and (Topic_Status <> 0)) or (AdminAllowed = 1) then
			Response.Write	"                      &nbsp;<a href=""post.asp?" & ArchiveLink & "method=EditTopic&REPLY_ID=" & Topic_ID & "&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconEditTopic,"Edit Topic","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (strAIM = "1") then
		if Trim(Member_AIM) <> "" then
			Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Member_Name,"display") & " an AOL message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (strICQ = "1") then
		if Trim(Member_ICQ) <> "" then
			Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=ICQ&ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Member_Name,"display") & " an ICQ Message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (strMSN = "1") then
		if Trim(Member_MSN) <> "" then
			Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Member_Name,"display") & "'s MSN Messenger address","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (strYAHOO = "1") then
		if Trim(Member_YAHOO) <> "" then
			Response.Write	"                      &nbsp;<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Member_YAHOO, "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Member_Name,"display") & " a Yahoo! Message","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	phSQL = "SELECT COUNT(*) AS EditCount FROM " & strTablePrefix & "POST_HISTORY WHERE T_ID = " & Topic_ID
	Set phRS = Server.CreateObject("ADODB.RecordSet")
	phRS.Open phSQL, my_Conn
	EditCount = phRS("EditCount")
	phRS.Close
	Set phRS = Nothing
	if EditCount <> 0 then
		if mlev = 4 or mlev = 3 then
			Response.Write	"                      &nbsp;<a href=""post_history.asp?" & ArchiveLink & "T_ID=" & Topic_ID & """>" & getCurrentIcon(strIconPostHistory,"View Topic History","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if ((Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status = 1) or (AdminAllowed = 1 and Topic_Status <= 1) and ArchiveView = "" ) then
		Response.Write	"                      &nbsp;<a href=""post.asp?" & ArchiveLink & "method=TopicQuote&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply with Quote","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
	end if
	if (strIPLogging = "1") then
		if (AdminAllowed = 1) then
			Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_viewip.asp?" & ArchiveLink & "mode=getIP&TOPIC_ID=" & TopicID & "&FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconIP,"View user's IP address","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
		end if
	end if
	if (AdminAllowed = 1) or (TMember_ID = MemberID and Topic_Replies < 1) then
		Response.Write	"                      &nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconDeleteReply,"Delete Topic","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewLine
	end if
	' DEM --> Start of Code added for Full Moderation
	if (AdminAllowed = 1 and Topic_Status > 1) then
		TopicString = "TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID
		Response.Write  "                      &nbsp;<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","align=""absmiddle"" hspace=""6""") & "</a>" & vbNewline
	End if
	' End of Code added for Full Moderation
	Response.Write	"                      <hr noshade size=""" & strFooterFontSize & """></td>" & vbNewLine & _
			"                    </tr>" & vbNewLine & _
			"                    <tr>" & vbNewLine & _
			"                      <td valign=""top"" height=""100%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"" id=""msg"">"
	if Request.QueryString("SearchTerms") <> "" then
		Response.Write	SearchHiLite(formatStr(Topic_Message))
	else
		Response.Write	formatStr(Topic_Message)
	end if
	Response.Write	"</span id=""msg""></font></td>" & vbNewLine & _
			"                    </tr>" & vbNewLine
	if CanShowSignature = 1 and Topic_Sig = 1 and Topic_MemberSig <> "" then
		Response.Write	"                    <tr>" & vbNewLine & _
				"                      <td valign=""bottom""><hr noshade size=""" & strFooterFontSize & """><font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><span class=""spnMessageText"">" & formatStr(Topic_MemberSig) & "</span></font></td>" & vbNewLine & _
				"                    </tr>" & vbNewLine
	end if
	if strEditedByDate = "1" and Topic_LastEditBy <> "" then
		if Topic_LastEditBy <> Topic_Author then 
			Topic_LastEditByName = getMemberName(Topic_LastEditBy)
		else
			Topic_LastEditByName = chkString(Member_Name,"display")
		end if
		Response.Write	"                    <tr>" & vbNewLine & _
				"                      <td valign=""bottom""><hr noshade size=""" & strFooterFontSize & """ color=""" & strForumFirstCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" &_
				"Edited by - " & Topic_LastEditByName & " on " & chkDate(Topic_LastEdit, " ", true) & "</font></td>" & vbNewLine & _
				"                    </tr>" & vbNewLine
	end if
	Response.Write	"                  </table>" & vbNewLine & _
			"                </td>" & vbNewLine & _
			"              </tr>" & vbNewLine
End Sub


sub PostingOptions() 
	Response.Write	"          <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
	if (mlev = 4 or mlev = 3 or mlev = 2 or mlev = 1) or (lcase(strNoCookies) = "1") or (strDBNTUserName = "") then
		if ((Cat_Status = 1) and (Forum_Status = 1)) then
			Response.Write	"          <a href=""post.asp?" & ArchiveLink & "method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"","align=""absmiddle""") & "</a>&nbsp;<a href=""post.asp?" & ArchiveLink & "method=Topic&FORUM_ID=" & Forum_ID & """>New Topic</a>" & vbNewLine
		else
			if (AdminAllowed = 1) then
				Response.Write	"          <a href=""post.asp?" & ArchiveLink & "method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderLocked,"","align=""absmiddle""") & "</a>&nbsp;<a href=""post.asp?" & ArchiveLink & "method=Topic&FORUM_ID=" & Forum_ID & """>New Topic</a>" & vbNewLine
			else
				Response.Write	"          " & getCurrentIcon(strIconFolderLocked,"","align=""absmiddle""") & "&nbsp;Forum Locked" & vbNewLine
			end if 
		end if 
		if ((Cat_Status = 1) and (Forum_Status = 1) and (Topic_Status = 1)) and ArchiveView = "" then
			Response.Write	"          <a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"","align=""absmiddle""") & "</a>&nbsp;<a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>Reply to Topic</a>" & vbNewLine
		else 
			if ((AdminAllowed = 1 and Topic_Status <= 1) and ArchiveView = "")  then
				Response.Write	"          <a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>"
				' DEM --> Added if statement to show normal icon for unmoderated posts.
				if Topic_Status = 1 and Cat_Status <> 0 and Forum_Status <> 0 then
					Response.Write	getCurrentIcon(strIconReplyTopic,"","align=""absmiddle""") & "</a>&nbsp;"
				else
					Response.Write	getCurrentIcon(strIconClosedTopic,"","align=""absmiddle""") & "</a>&nbsp;"
				end if 
				Response.Write	"<a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>Reply to Topic</a>" & vbNewLine
			else 
				if Topic_Status = 0 then
					Response.Write	getCurrentIcon(strIconClosedTopic,"","align=""absmiddle""") & "&nbsp;Topic Locked" & vbNewline
				end if
			end if
		end if 
		if lcase(strEmail) = "1" and Topic_Status < 2 then 
			if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 and mLev > 0 then
				if strSubscription > 0 and Cat_Subscription > 0 and Forum_Subscription > 0 then
					if InArray(strTopicSubs, Topic_ID) then
						Response.Write "          <br />" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "Y") & vbNewLine
					elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
						Response.Write "          <br />" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "Y") & vbNewLine
					end if
				end if
			end if
			if ((mlev <> 0) or (mlev = 0 and strLogonForMail <> "1")) and lcase(strShowSendToFriend) = "1" then
				Response.Write	"          <br /><a href=""JavaScript:openWindow('pop_send_to_friend.asp?url=" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "')"">" & getCurrentIcon(strIconSendTopic,"","align=""absmiddle""") & "</a>&nbsp;<a href=""JavaScript:openWindow('pop_send_to_friend.asp?url=" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "')"">Send Topic to a Friend</a>" & vbNewLine
			end if 
		end if 
		if lcase(strShowPrinterFriendly) = "1" and Topic_Status < 2 then
			Response.Write	"          <br /><a href=""JavaScript:openWindow5('pop_printer_friendly.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "')"">" & getCurrentIcon(strIconPrint,"","align=""absmiddle""") & "</a>&nbsp;<a href=""JavaScript:openWindow5('pop_printer_friendly.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "')"">Printer Friendly</a>" & vbNewLine
		end if
	end if 
	Response.Write	"          </font>"
end sub 

sub AdminOptions() 
	Response.Write	"                <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then
		if (Cat_Status = 0) then 
			if (mlev = 4) then
				Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & vbNewLine
			else
				Response.Write	"                " & getCurrentIcon(strIconFolderUnlocked,"Category Locked","") & vbNewLine
			end if
		else 
			if (Forum_Status = 0) then
				Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & vbNewLine
			else
				if (Topic_Status <> 0) then
					Response.Write	"                <a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Topic","") & "</a>" & vbNewLine
				else
					Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Topic","") & "</a>" & vbNewLine
				end if
			end if
		end if
		if ((Cat_Status <> 0) and (Forum_Status <> 0) and (Topic_Status <> 0)) or (AdminAllowed = 1) then
			Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=EditTopic&REPLY_ID=" & Topic_ID & "&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderPencil,"Edit Topic","hspace=""0""") & "</a>" & vbNewLine
		end if
		Response.Write	"                <a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Topic","") & "</a>" & vbNewLine & _
				"                <a href=""post.asp?" & ArchiveLink & "method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & vbNewLine
		if Topic_Status <= 1 and ArchiveView = "" then
			Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","") & "</a>" & vbNewLine
		end if
	end if 
	' DEM --> Start of Code added for Full Moderation
	if (AdminAllowed = 1 and CheckForUnModeratedPosts("TOPIC", Cat_ID, Forum_ID, Topic_ID) > 0) then
		TopicString = "TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "&REPLY_ID=X"
		Response.Write "                <a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject all posts for this Topic","") & "</a>" & vbNewline
	end if
	' DEM --> End of Code added for Full Moderation
	Response.Write "                </font>"
end sub 

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("                <form name=""PageNum" & fnum & """ action=""topic.asp"">" & vbNewLine)
		Response.Write("                <td bgcolor=""" & strHeadCellColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>" & vbNewLine)
		if Archiveview = "true" then Response.Write("                <input type=""hidden"" name=""ARCHIVE"" value=""" & ArchiveView & """>" & vbNewLine)
		Response.Write("                <input type=""hidden"" name=""TOPIC_ID"" value=""" & Request("TOPIC_ID") & """>" & vbNewLine)
		Response.Write("                <b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		for counter = 1 to maxpages
			if counter <> cLng(pge) then   
				Response.Write "                	<option value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			else
				Response.Write "                	<option selected value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			end if
		next
		Response.Write("                </select><b> of " & maxpages & "</b>" & vbNewLine)
		if Request.QueryString("SearchTerms") <> "" then Response.Write("                <input type=""hidden"" name=""SearchTerms"" value=""" & Request.QueryString("SearchTerms") & """>" & vbNewLine)
		Response.Write("                </font></td>" & vbNewLine)
		Response.Write("                </form>" & vbNewLine)
	end if
	top = "0"
end sub 

Sub Topic_nav()    

	if prevTopic = "" then
		strSQL = "SELECT T_SUBJECT, TOPIC_ID "
		strSql = strSql & "FROM " & strActivePrefix & "TOPICS "
		strSql = strSql & "WHERE T_LAST_POST > '" & Topic_LastPost
		strSql = strSql & "' AND FORUM_ID = " & Forum_ID
		strSql = strSql & " AND T_STATUS < 2"  ' Ignore unapproved/held posts
		strSql = strSql & " ORDER BY T_LAST_POST;"

		set rsPrevTopic = my_conn.Execute(TopSQL(strSql,1))

		if rsPrevTopic.EOF then
			prevTopic = getCurrentIcon(strIconBlank,"","align=""top"" hspace=""6""")
		else
			prevTopic = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & rsPrevTopic("TOPIC_ID") & """>" & getCurrentIcon(strIconGoLeft,"Previous Topic","align=""top"" hspace=""6""") & "</a>"
		end if

		rsPrevTopic.close
		set rsPrevTopic = nothing
	else
		prevTopic = prevTopic
	end if

	if NextTopic = "" then
		strSQL = "SELECT T_SUBJECT, TOPIC_ID "
		strSql = strSql & "FROM " & strActivePrefix & "TOPICS "
		strSql = strSql & "WHERE T_LAST_POST < '" & Topic_LastPost
		strSql = strSql & "' AND FORUM_ID = " & Forum_ID
		strSql = strSql & " AND T_STATUS < 2"  ' Ignore unapproved/held posts
		strSql = strSql & " ORDER BY T_LAST_POST DESC;"

		set rsNextTopic = my_conn.Execute(TopSQL(strSql,1))

		if rsNextTopic.EOF then
			nextTopic = getCurrentIcon(strIconBlank,"","align=""top"" hspace=""6""")
		else
			nextTopic = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & rsNextTopic("TOPIC_ID") & """>" & getCurrentIcon(strIconGoRight,"Next Topic","align=""top"" hspace=""6""") & "</a>"
		end if

		rsNextTopic.close
		set rsNextTopic = nothing
	else
		nextTopic = nextTopic
	end if

	Response.Write ("                " & prevTopic & "<b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&nbsp;Topic&nbsp;</font></b>" & nextTopic)

end sub

function SearchHiLite(fStrMessage)
	'function derived from HiLiTeR by 2eNetWorX
	fArr = split(replace(Request.QueryString("SearchTerms"),";",""), ",")
	strBuffer = ""
	for iPos = 1 to len(fStrMessage)
		bChange = False
		'Looks for html tags
		if mid(fStrMessage, iPos, 1) = "<" then
			bInHTML = True
		end if
		'Looks for End of html tags
		if bInHTML = True then
			if mid(fStrMessage, iPos, 1) = ">" then
				bInHTML = False
			end if
		end if
		if bInHTML <> True then  
			for i = 0 to UBound(fArr)
				if fArr(i) <> "" then
					if lcase(mid(fStrMessage, iPos, len(fArr(i)))) = lcase(fArr(i)) then
						bChange = True
						strBuffer = strBuffer & "<span class=""spnSearchHighlight"" id=""hilite"">" & _ 
						mid(fStrMessage, iPos, len(fArr(i))) & "</span id=""hilite"">"
						iPos = iPos + len(fArr(i)) - 1
					end if
				end if
			next
		end if
		if Not bChange then
			strBuffer = strBuffer & mid(fStrMessage, iPos, 1)
		end if
	next
	SearchHiLite = strBuffer
end function

Sub QuickReply()
	intSigDefault = getSigDefault(MemberID)
	Response.Write	"    <script language=""JavaScript"" type=""text/javascript"" src=""inc_code.js""></script>" & vbNewLine & _
			"      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <form name=""PostTopic"" method=""post"" action=""post_info.asp"" onSubmit=""return validate();"">" & vbNewLine & _
			"              <input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & vbNewLine & _
			"              <input name=""Method_Type"" type=""hidden"" value=""Reply"">" & vbNewLine & _
			"              <input name=""TOPIC_ID"" type=""hidden"" value=""" & Topic_ID & """>" & vbNewLine & _
			"              <input name=""FORUM_ID"" type=""hidden"" value=""" & Forum_ID & """> " & vbNewLine & _
			"              <input name=""CAT_ID"" type=""hidden"" value=""" & Cat_ID & """>" & vbNewLine & _
			"              <input name=""Refer"" type=""hidden"" value=""" & request.servervariables("SCRIPT_NAME") & "?" & chkString(Request.QueryString,"refer") & """>" & vbNewLine & _
			"              <input name=""UserName"" type=""hidden"" value=""" & strDBNTUserName & """>" & vbNewLine & _
			"              <input name=""Password"" type=""hidden"" value=""" & Request.Cookies(strUniqueID & "User")("Pword") & """>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strHeadCellColor & """ noWrap vAlign=""top"" colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Quick Reply</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strForumCellColor & """ noWrap vAlign=""top"" align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText""><b>Message:&nbsp;</b><br />" & vbNewLine & _
			"                <br />" & vbNewLine & _
			"                  <table border=""0"">" & vbNewLine & _
			"                    <tr>" & vbNewLine & _
			"                      <td align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine
	if strAllowHTML = "1" then
		Response.Write	"                      * HTML is ON<br />" & vbNewLine
	else
		Response.Write	"                      * HTML is OFF<br />" & vbNewLine
	end if
	if strAllowForumCode = "1" then
		Response.Write	"                      * <a href=""JavaScript:openWindow6('pop_forum_code.asp')"">Forum Code</a> is ON<br />" & vbNewLine
	else
		Response.Write	"                      * Forum Code is OFF<br />" & vbNewLine
	end if
	if strSignatures = "1" then
		Response.Write	"                      <br /><input name=""Sig"" id=""Sig"" type=""checkbox"" value=""yes""" & chkCheckbox(intSigDefault,1,true) & "><label for=""Sig"">Include Signature</label><br />" & vbNewLine
	end if
	Response.Write	"                      </font></td>" & vbNewLine & _
			"                    </tr>" & vbNewLine & _
			"                  </table>" & vbNewLine & _
			"                </span></font></td>" & vbNewLine & _
			"                <td width=""" & strTopicWidthRight & """ bgColor=""" & strForumCellColor & """><textarea name=""Message"" cols=""50"" rows=""6"" wrap=""VIRTUAL"" style=""width:100%""></textarea><br /></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strForumCellColor & """ noWrap align=""center"" colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><input name=""Submit"" type=""submit"" value=""Submit Reply"">&nbsp;<input name=""Preview"" type=""button"" value=""Preview Reply"" onclick=""OpenPreview()"">"
	'Response.Write	"&nbsp;<input name=""Reset"" type=""reset"" value=""Reset Form""></font></td>" & vbNewLine & _
	Response.Write	"</font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              </form>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      <br />" & vbNewLine
end sub
%>
