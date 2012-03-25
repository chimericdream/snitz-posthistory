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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_func_count.asp" -->
<% 
if request("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
	ArchiveView = "true"
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
end if

'Topic Move Check
Dim blnTopicMoved
Dim fSubscription

fsubscription = 1
blnTopicMoved = false

if strAuthType = "db" and strDBNTUserName = "" and len(Request.Form("Password")) <> 64 then
	strPassword = sha256("" & Request.Form("Password"))
else
	strPassword = ChkString(Request.Form("Password"),"SQLString")
end if

if strAuthType = "db" and strDBNTUserName = "" then
	strDBNTUserName = Request.Form("UserName")
	if mLev = 0 then mLev = cLng(chkUser(strDBNTUserName, strPassword,-1))
end if

MethodType = chkString(Request.Form("Method_Type"),"SQLString")

if Request.Form("CAT_ID") <> "" then
	if IsNumeric(Request.Form("CAT_ID")) = True then
		Cat_ID = cLng(Request.Form("CAT_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.Form("FORUM_ID") <> "" then
	if IsNumeric(Request.Form("FORUM_ID")) = True then
		Forum_ID = cLng(Request.Form("FORUM_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.Form("TOPIC_ID") <> "" then
	if IsNumeric(Request.Form("TOPIC_ID")) = True then
		Topic_ID = cLng(Request.Form("TOPIC_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.Form("REPLY_ID") <> "" then
	if IsNumeric(Request.Form("REPLY_ID")) = True then
		Reply_ID = cLng(Request.Form("REPLY_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.Form("Subscription") <> "" then
	fSubscription = cLng(Request.Form("Subscription"))
else
 	fSubscription = 0
end if

if Request.Form("cookies") = "yes" then
	strSelectSize = Request.Form("SelectSize")
end if

if strSelectSize = "" or IsNull(strSelectSize) then 
	strSelectSize = Request.Cookies(strUniqueID & "strSelectSize")
end if
if not(IsNull(strSelectSize)) and strSelectSize <> "" then 
	if strSetCookieToForum = 1 then
    		Response.Cookies(strUniqueID & "strSelectSize").Path = strCookieURL
	else
		Response.Cookies(strUniqueID & "strSelectSize").Path = "/"
	end if
	Response.Cookies(strUniqueID & "strSelectSize") = strSelectSize
	Response.Cookies(strUniqueID & "strSelectSize").expires = dateAdd("yyyy", 1, strForumTimeAdjust)
end if

if MethodType = "Edit" or _
MethodType = "EditTopic" or _
MethodType = "Reply" or _
MethodType = "ReplyQuote" or _
MethodType = "TopicQuote" then
	'## check if topic exists in TOPICS table
	set rsTCheck = my_Conn.Execute ("SELECT TOPIC_ID FROM " & strActivePrefix & "TOPICS WHERE TOPIC_ID = " & Topic_ID)
	if rsTCheck.EOF or rsTCheck.BOF then
		set rsTCheck = nothing
		Go_Result "Sorry, that Topic no longer exists in the Database", 0
	end if
	set rsTCheck = nothing
end if

'set rs = Server.CreateObject("ADODB.RecordSet")

err_Msg = ""
ok = "" 

if ArchiveView <> "" then
	if MethodType = "Reply" or _
	MethodType = "ReplyQuote" or _
	MethodType = "TopicQuote" then
		Go_Result "This is not allowed in the Archives.", 0
	end if
end if

if MethodType = "Edit" then
	'## Forum_SQL - Get the author of the reply
	strSql = "SELECT R_AUTHOR " 
	strSql = strSql & " FROM " & strActivePrefix & "REPLY "
	strSql = strSql & " WHERE REPLY_ID = " & REPLY_ID
 
	set rsStatus = my_Conn.Execute(strSql)
	if rsStatus.EOF or rsStatus.BOF then
		rsStatus.close
		set rsStatus = nothing
		Go_Result "Please don't attempt to edit the URL<br />to gain access to locked Forums/Categories.", 0
	else
		strReplyAuthor = rsStatus("R_AUTHOR")
		rsStatus.close
		set rsStatus = nothing
	end if
end if

if MethodType = "Edit" or _
MethodType = "EditTopic" or _
MethodType = "Reply" or _
MethodType = "ReplyQuote" or _
MethodType = "Topic" or _
MethodType = "TopicQuote" then
	if MethodType <> "Topic" then
		'## Forum_SQL - Find out if the Category, Forum or Topic is Locked or Un-Locked and if it Exists
		strSql = "SELECT C.CAT_STATUS, C.CAT_NAME, " &_
		"F.FORUM_ID, F.F_STATUS, F.F_TYPE, F.F_SUBJECT, " &_
		"T.T_STATUS, T.T_AUTHOR, T.T_SUBJECT " &_
		" FROM " & strTablePrefix & "CATEGORY C, " &_
		strTablePrefix & "FORUM F, " &_
		strActivePrefix & "TOPICS T" &_
		" WHERE C.CAT_ID = T.CAT_ID " &_
		" AND F.FORUM_ID = T.FORUM_ID " &_
		" AND T.TOPIC_ID = " & Topic_ID & ""
	else
		'## Forum_SQL - Find out if the Category or Forum is Locked or Un-Locked and if it Exists
		strSql = "SELECT C.CAT_STATUS, C.CAT_NAME, " &_
		"F.FORUM_ID, F.F_STATUS, F.F_TYPE, F.F_SUBJECT " &_
		" FROM " & strTablePrefix & "CATEGORY C, " &_
		strTablePrefix & "FORUM F" &_
		" WHERE C.CAT_ID = F.CAT_ID " &_
		" AND F.FORUM_ID = " & Forum_ID & ""
        end if
 
	set rsStatus = my_Conn.Execute(strSql)
	if rsStatus.EOF or rsStatus.BOF then
		rsStatus.close
		set rsStatus = nothing
		Go_Result "Please don't attempt to edit the URL<br />to gain access to locked Forums/Categories.", 0
	else
		blnCStatus = rsStatus("CAT_STATUS")
		strCatTitle = rsStatus("CAT_NAME")
		blnFStatus = rsStatus("F_STATUS")
		Forum_ID = rsStatus("FORUM_ID")
		Forum_Type = rsStatus("F_TYPE")
		strForum_Title = rsStatus("F_SUBJECT")
		if MethodType <> "Topic" then
			blnTStatus = rsStatus("T_STATUS")
			strTopicAuthor = rsStatus("T_AUTHOR")
			strTopicTitle = rsStatus("T_SUBJECT")
		else
			blnTStatus = 1
		end if
		rsStatus.close
		set rsStatus = nothing
	end if
 
	if mLev = 4 then
		AdminAllowed = 1
		ForumChkSkipAllowed = 1
	elseif mLev = 3 then
		if chkForumModerator(Forum_ID, ChkString(strDBNTUserName, "decode")) = "1" then
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
 
	select case MethodType
		case "Topic"
			if (Forum_Type = 1) then
				Go_Result "You have attempted to post a New Topic to a Forum designated as a Web Link",0
			end if
			if (blnCStatus = 0) and (AdminAllowed = 0) then
				Go_Result "You have attempted to post a New Topic to a Locked Category", 0
			end if
			if (blnFStatus = 0) and (AdminAllowed = 0) then
				Go_Result "You have attempted to post a New Topic to a Locked Forum", 0
			end if
		case "EditTopic"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to edit a Locked Topic", 0
			end if
		case "Reply", "ReplyQuote", "TopicQuote"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to Reply to a Locked Topic", 0
			end if
		case "Edit"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to Edit a Reply to a Locked Topic", 0
			end if
	end select
	if strPrivateForums = "1" and ForumChkSkipAllowed = 0 then
		if not(chkForumAccess(Forum_ID,MemberID,false)) then
    			Go_Result "You do not have access to post to this forum", 0
  		end if
	end if
end if

' If Creating a new topic or reply, the subscription and moderation capabilities will need to be checked.
Moderation = "No"
if MethodType = "Topic" or _
   MethodType = "Edit" or _
   MethodType = "Reply" or _
   MethodType = "ReplyQuote" or _
   MethodType = "TopicQuote" or _
   MethodType = "Forum" or _
   MethodType = "EditForum"  then

	if strModeration > 0 or strSubscription > 0 then
		'## Forum_SQL - Get the Cat_Subscription, Cat_Moderation, Forum_Subscription, Forum_Moderation
		strSql = "SELECT C.CAT_MODERATION, C.CAT_SUBSCRIPTION, C.CAT_NAME "
		if MethodType <> "Forum" then
			strSql = strSql & ", F.F_MODERATION, F.F_SUBSCRIPTION "
		end if
		strsql = strsql & " FROM " & strTablePrefix & "CATEGORY C"
		if MethodType <> "Forum" then
			strSql = strSql & ", " & strTablePrefix & "FORUM F"
		end if
		strSql = strSql & " WHERE C.CAT_ID = " & Cat_ID
		if MethodType <> "Forum" then
			strSql = strSql & "   AND F.FORUM_ID = " & Forum_ID
		end if
		set rsCheck = my_Conn.Execute (strSql)

		CatName           = rsCheck("CAT_NAME")
		CatSubscription   = rsCheck("CAT_SUBSCRIPTION")
		CatModeration     = rsCheck("CAT_MODERATION")
		if MethodType <> "Forum" then
			ForumSubscription = rsCheck("F_SUBSCRIPTION")
			ForumModeration   = rsCheck("F_MODERATION")
		end if
		rsCheck.Close
		set rsCheck = nothing
		if MethodType <> "Forum" then
			'## Moderators and Admins are not subject to Moderation
			if strModeration = 0 or mlev = 4 or chkForumModerator(Forum_ID, strDBNTUserName) = "1" then
				Moderation = "No"
			'## Is Moderation allowed on the category?
			elseif CatModeration = 1 then
				'## if this is a topic, is forum moderation set to all posts or topic?
				if (ForumModeration = 1 or ForumModeration = 2) and (MethodType = "Topic") then
					Moderation = "Yes"
				'## if this is a reply, is forum moderation set to all posts or reply?
				elseif (ForumModeration = 1 or ForumModeration = 3) and (MethodType <> "Topic") then
					Moderation = "Yes"
				end if
		    	end if
	      	end if
	end if
end if

if MethodType = "Edit" then
	member = cLng(ChkUser(strDBNTUserName, strPassword, strReplyAuthor))
	Select Case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin, a Moderator or the Author can change this post", 0
			Response.End
		case 3 '## Moderator so OK - check the Moderator of this forum
			if chkForumModerator(Forum_ID, strDBNTUserName) = "0" then
				Go_Result "Only an Admin, a Moderator or the Author can change this post", 0
			end if
			if strReplyAuthor = intAdminMemberID and MemberID <> intAdminMemberID then
				Go_Result "Only the Forum Admin can change this post", 0
			end if
		case 4 '## Admin so OK
			if strReplyAuthor = intAdminMemberID and MemberID <> intAdminMemberID then
				Go_Result "Only the Forum Admin can change this post", 0
			end if
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	Err_Msg = ""

	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for your Reply</li>"
	end if
	if Err_Msg = "" then
		strSql = "SELECT REPLY_ID, R_MESSAGE, R_DATE, R_LAST_EDIT, R_LAST_EDITBY, R_AUTHOR FROM "
		strSql = strSql & strActivePrefix & "REPLY WHERE REPLY_ID = " & Reply_ID
		Set phRS = Server.CreateObject("ADODB.RecordSet")
		phRS.Open strSql, my_Conn

		phReplyID = phRS("REPLY_ID")
		phRMessage = phRS("R_MESSAGE")
		If mlev = 4 then
			phRDate = DateToStr(Now())
		else
			phRDate = phRS("R_LAST_EDIT")
		end if
		If phRDate = "" Or IsNull(phRDate) Then
			phRDate = phRS("R_DATE")
		End If
		phRLastEditby = phRS("R_LAST_EDITBY")
		If phRLastEditby = "" Or IsNull(phRLastEditby) Then
			phRLastEditby = 0
		End If
		phRAuthor = phRS("R_AUTHOR")

		phRS.Close
		Set phRS = Nothing

		strSql = "INSERT INTO " & strTablePrefix & "POST_HISTORY "
		strSql = strSql & "(R_ID, P_MESSAGE, P_AUTHOR, P_LAST_EDITBY, P_DATE) VALUES "
		strSql = strSql & "(" & phReplyID & ", '" & ChkString(phRMessage,"sqlstring") & "', " & phRAuthor & ", " & phRLastEditby & ", '" & phRDate & "')"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		
		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strActivePrefix & "REPLY "
		strSql = strSql & " SET R_MESSAGE = '" & txtMessage & "'"
		if Request.Form("sig") = "yes" and strDSignatures = "1" then
		 	strSql = strSql & ", R_SIG = 1"
		else
			strSql = strSql & ", R_SIG = 0"
		end if
		if mLev < 4 and strEditedByDate = "1" then
			strSql = strSql & ", R_LAST_EDIT = '"  & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ", R_LAST_EDITBY = "  & MemberID
		end if
		strSql = strSql & " WHERE REPLY_ID=" & Reply_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		if mLev <> 4 and Moderation = "No" then
			'## Forum_SQL - Update Last Post
			strSql = " UPDATE " & strTablePrefix & "FORUM"
			strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    F_LAST_POST_AUTHOR = " & MemberID
			strSql = strSql & ",    F_LAST_POST_TOPIC_ID = " & Topic_ID
			strSql = strSql & ",    F_LAST_POST_REPLY_ID = " & Reply_ID
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			'## Forum_SQL - Update Last Post
			strSql = " UPDATE " & strActivePrefix & "TOPICS"
			strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    T_LAST_POST_AUTHOR = " & MemberID
			strSql = strSql & ",    T_LAST_POST_REPLY_ID = " & Reply_ID
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		end if

		if Moderation = "No" then
			'## Subscribe checkbox start ##
			if request.form("TNotify") <> "" then
				if request.form("TNotify") = "1" then
					AddSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, Topic_ID
				else
					DeleteSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, Topic_ID
				end if
			end if
			'## Subscribe checkbox end ##
		end if

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			Go_Result "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "EditTopic" then
	member = cLng(ChkUser(strDBNTUserName, strPassword, strTopicAuthor))
	select case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin, a Moderator or the Author can change this post", 0
			Response.End
		case 3 '## Moderator so 
			if chkForumModerator(Forum_ID, strDBNTUserName) = "0" then
				Go_Result "Only an Admin, a Moderator or the Author can change this post", 0
			end if
			if strTopicAuthor = intAdminMemberID and MemberID <> intAdminMemberID then
				Go_Result "Only the Forum Admin can change this post", 0
			end if
		case 4 '## Admin so OK
			if strTopicAuthor = intAdminMemberID and MemberID <> intAdminMemberID then
				Go_Result "Only the Forum Admin can change this post", 0
			end if
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = chkString(Request.Form("Message"),"message")
	txtSubject = chkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Topic</li>"
	end if
	if Len(Request.Form("Subject")) > 50 then 
		Err_Msg = Err_Msg & "<li>The Subject can not be greater than 50 characters</li>"
	end if
	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the Topic</li>"
	end if
	if Err_Msg = "" then
		'##Get Status of this Topic
		strSql = "SELECT T_STATUS, T_UREPLIES"
		strSql = Strsql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

		set rsTopicStatusCheck = my_Conn.Execute (strSql)

		Topic_Status = rsTopicStatusCheck("T_STATUS")
		Topic_UReplies = rsTopicStatusCheck("T_UREPLIES")

		rsTopicStatusCheck.Close
		set rsTopicStatusCheck = nothing

		'## Set array to pull out CAT_ID and FORUM_ID from dropdown values in post.asp
		aryForum = split(Request.Form("Forum"), "|")

		'## if the forum we are moving to doesn't have MODERATION, and this topic did have that
		'## we are going to have to auto-approve the topic !

		AutoApprove = "No"
		Moderation = "No"

		if Forum_ID <> cLng(aryForum(1)) then
			blnTopicMoved = true
			strSql = "SELECT " & strTablePrefix & "FORUM.F_MODERATION "
			strSql = strsql & " FROM " & strTablePrefix & "FORUM "
			strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Forum_ID

			set rsForumCheck = my_Conn.Execute (strSql)

			ForumModeration = rsForumCheck("F_MODERATION")

			rsForumCheck.Close
			set rsForumCheck = nothing
		
			'## Is Moderation allowed on the topic in the old forum ?
			if (ForumModeration = 1 or ForumModeration = 2) then
				Moderation = "Yes"
			end if

			if Moderation = "Yes" and Topic_Status > 0 then
	
				strSql = "SELECT " & strTablePrefix & "FORUM.F_MODERATION "
				strSql = Strsql & " FROM " & strTablePrefix & "FORUM "
				strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & cLng(aryForum(1))

				set rsNewForumCheck = my_Conn.Execute (strSql)

				NewForumModeration   = rsNewForumCheck("F_MODERATION")

				rsNewForumCheck.Close
				set rsNewForumCheck = nothing

				'## Is Moderation allowed on the topic in the new forum ?
				if not(NewForumModeration = 1 or NewForumModeration = 2) then
					AutoApprove = "Yes"
				end if
			end if
		end if

		if Moderation = "Yes" and AutoApprove = "Yes" and Topic_UReplies > 0 then
			Go_Result "There was an error = The Topic you are attempting to move to an UnModerated Forum has UnModerated Replies<br />Please either approve or delete them and then try again.", 0
			Response.End
		end if

		strSql = "SELECT TOPIC_ID, T_SUBJECT, T_MESSAGE, T_DATE, T_LAST_EDIT, T_LAST_EDITBY, T_AUTHOR FROM "
		strSql = strSql & strActivePrefix & "TOPICS WHERE TOPIC_ID = " & Topic_ID
		Set phRS = Server.CreateObject("ADODB.RecordSet")
		phRS.Open strSql, my_Conn

		phTopicID = phRS("TOPIC_ID")
		phTMessage = phRS("T_MESSAGE")
		phTSubject = phRS("T_SUBJECT")
		phTDate = phRS("T_LAST_EDIT")
		If mlev = 4 then
			phTDate = DateToStr(Now())
		else
			phTDate = phRS("T_LAST_EDIT")
		end if
		If phTDate = "" Or IsNull(phTDate) Then
			phTDate = phRS("T_DATE")
		End If
		phTLastEditby = phRS("T_LAST_EDITBY")
		If phTLastEditby = "" Or IsNull(phTLastEditby) Then
			phTLastEditby = 0
		End If
		phTAuthor = phRS("T_AUTHOR")

		phRS.Close
		Set phRS = Nothing

		strSql = "INSERT INTO " & strTablePrefix & "POST_HISTORY "
		strSql = strSql & "(T_ID, T_SUBJECT, P_MESSAGE, P_AUTHOR, P_LAST_EDITBY, P_DATE) VALUES "
		strSql = strSql & "(" & phTopicID & ", '" & ChkString(phTSubject,"sqlstring") & "', '" & ChkString(phTMessage,"sqlstring") & "', " & phTAuthor & ", " & phTLastEditby & ", '" & phTDate & "')"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		'## Forum_SQL
		strSql = "UPDATE " & strActivePrefix & "TOPICS "
		strSql = strSql & " SET T_MESSAGE = '" & txtMessage & "'"
		strSql = strSql & ", T_SUBJECT = '" & txtSubject & "'"
		if blnTopicMoved then
			strSql = strSql & ", CAT_ID = " & cLng(aryForum(0))
			strSql = strSql & ", FORUM_ID = " & cLng(aryForum(1))
			if AutoApprove = "Yes" then
				strSql = strSql & ", T_STATUS = 1 "
			end if
		end if
		if Request.Form("sig") = "yes" and strDSignatures = "1" then
		 	strSql = strSql & ", T_SIG = 1"
		else
			strSql = strSql & ", T_SIG = 0"
		end if
		if mLev < 4 and strEditedByDate = "1" then
			strSql = strSql & ", T_LAST_EDIT = '"  & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ", T_LAST_EDITBY = "  & MemberID
		end if                                                   
		if ForumChkSkipAllowed = 1 then
			if Request.Form("sticky") = 1 then
				strSql = strSql & ", T_STICKY = " & 1
				strSql = strSql & ", T_ARCHIVE_FLAG = " & 0
			else
				strSql = strSql & ", T_STICKY = " & 0
				strSql = strSql & ", T_ARCHIVE_FLAG = " & 1
			end if
		end if
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

		my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords

		'# Subscribe checkbox start ##
		if request.form("TNotify") <> "" then
			if request.form("TNotify") = "1" then
				AddSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, Topic_ID
			elseif request.form("TNotify") = "0" then
				DeleteSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, Topic_ID
			end if
		end if
		'## Subscribe checkbox end ##

		if blnTopicMoved then
			if strEmail = "1" and strMoveNotify = "1" then DoAutoMoveEmail(Topic_ID)	
			strSQL = "SELECT F_SUBSCRIPTION FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & cLng(aryForum(1))
			set rs = my_conn.execute (strSQL)
			if rs("F_SUBSCRIPTION") < 3 then
				strSQL = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS WHERE TOPIC_ID=" & Topic_ID
				my_conn.execute(strSQL),,adCmdText + adExecuteNoRecords
			end if
			rs.close
			set rs = nothing
		end if
		if Forum_ID <> cLng(aryForum(1)) then
			'## Forum_SQL
			strSql = "UPDATE " & strActivePrefix & "REPLY "
			strSql = strSql & " SET CAT_ID = " & cLng(aryForum(0))
			strSql = strSql & ", FORUM_ID = " & cLng(aryForum(1))
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			
			'set rs = Server.CreateObject("ADODB.Recordset")
			
			'## if the topic hasn't been approved yet, it isn't counted either
			'## so then the topic count doesn't need to be updated

	        	if Moderation = "No" or AutoApprove = "Yes" or Topic_Status < 2 then

				'## Forum_SQL - count total number of replies in Topics table
				strSql = "SELECT T_REPLIES, T_LAST_POST, T_LAST_POST_AUTHOR "
				strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

				set rs = my_Conn.Execute (strSql)
			
				intResetCount = rs("T_REPLIES") + 1
				strT_Last_Post = rs("T_LAST_POST")
				strT_Last_Post_Author = rs("T_LAST_POST_AUTHOR")
			
				rs.Close
				set rs = nothing

				'## Forum_SQL - Get last_post and last_post_author for MoveFrom-Forum
				strSql = "SELECT TOPIC_ID, T_LAST_POST, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID "
				strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
				strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
				strSql = strSql & " ORDER BY T_LAST_POST DESC;"

				set rs = my_Conn.Execute (strSql)
			
				if not rs.eof then
					strLast_Post_Topic_ID = rs("TOPIC_ID")
					strLast_Post = rs("T_LAST_POST")
					strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
					strLast_Post_Reply_ID = rs("T_LAST_POST_REPLY_ID")
				else
					strLast_Post_Topic_ID = 0
					strLast_Post = ""
					strLast_Post_Author = 0
					strLast_Post_Reply_ID = 0
				end if
			
				rs.Close
				set rs = nothing

		        	if Moderation = "No" or Topic_Status < 2 then
					'## Forum_SQL - Update count of replies to a topic in Forum table

					strSql = "UPDATE " & strTablePrefix & "FORUM SET "
					strSql = strSql & " F_COUNT = F_COUNT - " & intResetCount
					strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
					my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL
					strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
					strSql = strSql & " F_TOPICS = F_TOPICS - 1 "
					strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
					my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords

				end if 

				strSql = "UPDATE " & strTablePrefix & "FORUM SET "
				'if strLast_Post <> "" then 
					strSql = strSql & "F_LAST_POST = '" & strLast_Post & "'"
					'if strLast_Post_Author <> "" then 
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
					'end if
				'end if
				strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & strLast_Post_Topic_ID
				strSql = strSql & ", F_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID
				strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords

				'## Forum_SQL - Get last_post and last_post_author for Forum
				strSql = "SELECT TOPIC_ID, T_LAST_POST, T_LAST_POST_AUTHOR, T_AUTHOR, T_LAST_POST_REPLY_ID "
				strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
				strSql = strSql & " WHERE FORUM_ID = " & cLng(aryForum(1))
				strSql = strSql & " ORDER BY T_LAST_POST DESC;"

				set rs = my_Conn.Execute (strSql)
			
				if not rs.eof then
					strAuthor = getMemberName(strT_Last_Post_Author)
					strLast_Post_Topic_ID = rs("TOPIC_ID")
					strLast_Post = rs("T_LAST_POST")
					strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
					strLast_Post_Reply_ID = rs("T_LAST_POST_REPLY_ID")
				else
					strAuthor = ""
					strLast_Post_Topic_ID = 0
					strLast_Post = ""
					strLast_Post_Author = ""
					strLast_Post_Reply_ID = 0
				end if
			
				rs.Close
				set rs = nothing
				'Huw -- Update member count
				if (AutoApprove = "Yes") and blnTStatus = 2 and blnTopicMoved then
					set rsFCountMP = my_Conn.Execute("SELECT F_COUNT_M_POSTS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & cLng(aryForum(1)))
					ForumCountMPosts = rsFCountMP("F_COUNT_M_POSTS")
					rsFCountMP.close
					set rsFCountMP = nothing
					if ForumCountMPosts <> 0 then
						doUCount(strT_Last_Post_Author)
					end if
					doULastPost(strT_Last_Post_Author)
				end if
				'## Forum_SQL - Update count of replies to a topic in Forum table
				strSql = "UPDATE " & strTablePrefix & "FORUM SET "
				strSql = strSql & " F_COUNT = (F_COUNT + " & intResetCount & ")"
				strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & strLast_Post_Topic_ID
				strSql = strSql & ", F_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID
				if strLast_Post <> "" then 
					strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
					if strLast_Post_Author <> "" then 
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
					end if
				end if
				strSql = strSql & " WHERE FORUM_ID = " & cLng(aryForum(1))
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords

				'## Forum_SQL
				strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
				strSql = strSql & " F_TOPICS = F_TOPICS + 1 "
				strSql = strSql & " WHERE FORUM_ID = " & cLng(aryForum(1))
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			end if
		else
			if mLev <> 4 and Moderation = "No" then
				'## Forum_SQL - Update Last Post
				strSql = " UPDATE " & strTablePrefix & "FORUM"
				strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
				strSql = strSql & ",    F_LAST_POST_AUTHOR = " & MemberID
				strSql = strSql & ",    F_LAST_POST_TOPIC_ID = " & Topic_ID
				strSql = strSql & ",    F_LAST_POST_REPLY_ID = " & 0
				strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				'## Forum_SQL - Update Last Post
				strSql = " UPDATE " & strActivePrefix & "TOPICS"
				strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
				strSql = strSql & ",    T_LAST_POST_AUTHOR = " & MemberID
				strSql = strSql & ",    T_LAST_POST_REPLY_ID = " & 0
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			end if
		end if
		err_Msg = ""
		aryForum = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "Topic" then
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_LASTPOSTDATE, " & strDBNTSQLName
	if strAuthType = "db" then
		strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(strDBNTUserName, "SQLString") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	if strAuthType = "db" then
		strSql = strSql & " AND   M_PASSWORD = '" & ChkString(strPassword, "SQLString") &"'"
		QuoteOk = (ChkQuoteOk(strDBNTUserName) and ChkQuoteOk(strPassword))
	else
		QuoteOk = ChkQuoteOk(strDBNTUserName)
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) or not (ChkQuoteOk(strPassword))then '##  Invalid Password
		rs.close
		set rs = nothing
		Go_Result "Invalid UserName or Password!", 0
		Response.End
	else
		if strPrivateForums = "1" and ForumChkSkipAllowed = 0 then
			if not(chkForumAccess(Forum_ID, MemberID,false)) then
				Go_Result "You are not allowed to post in this forum !", 0
			end if
		end if
		if strFloodCheck = 1 then
			if rs("M_LASTPOSTDATE") > DateToStr(DateAdd("s",strFloodCheckTime,strForumTimeAdjust)) and mLev < 3 then
				strTimeLimit = replace(strFloodCheckTime, "-", "")
				Go_Result "Sorry! We have flood control activated.<br />You cannot post within " & strTimeLimit & " seconds of your last post.<br />Please try again after this period of time elapses.", 0
			end if
		end if

		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("Subject"),"SQLString")
		UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		if UserIPAddress = "" then
			UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
		end if
		if strBadWordFilter = "1" then
			txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
		end if

		if txtMessage = " " then
			Go_Result "You must post a message!", 0
			Response.End
		end if
		if txtSubject = " " then
			Go_Result "You must post a subject!", 0
			Response.End
		end if         
		if Len(Request.Form("Subject")) > 50 then 
			Go_Result "The Subject can not be greater than 50 characters", 0
			Response.End
		end if
		if strSignatures = "1" and strDSignatures <> "1" then
			if Request.Form("sig") = "yes" and GetSig(strDBNTUserName) <> " " then
				txtMessage = txtMessage & vbNewline & vbNewline & ChkString(GetSig(strDBNTUserName), "signature" )
			end if
		end if

		'## Forum_SQL - Add new post to Topics Table
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS (FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_LAST_POST_REPLY_ID"
		strSql = strSql & ", T_DATE"
		strSql = strSql & ", T_STATUS"
		if strIPLogging <> "0" then
			strSql = strSql & ", T_IP"
		end if
		strSql = strSql & ", T_STICKY"
		strSql = strSql & ", T_SIG"
		strSql = strSql & ", T_ARCHIVE_FLAG"
		strSql = strSql & ", T_REPLIES"
		strSql = strSql & ", T_UREPLIES"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Forum_ID
		strSql = strSql & ", " & Cat_ID
		strSql = strSql & ", '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", 0 "
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		if Request.Form("lock") = 1 and ForumChkSkipAllowed = 1 then
			strSql = strSql & ", 0 "
		else
			if Moderation = "Yes" then
				strSql = strSql & ", 2 "
			else
				strSql = strSql & ", 1 "
			end if
		end if
		if strIPLogging <> "0" then
			strSql = strSql & ", '" & UserIPAddress & "'"
		end if
		if ForumChkSkipAllowed = 1 then
			if Request.Form("sticky") = 1 then
				strSql = strSql & ", 1 "
			else
				strSql = strSql & ", 0 "
			end if
		else
			strSql = strSql & ", 0 "
		end if
		if Request.Form("sig") = "yes" and strDSignatures = "1" then
		 	strSql = strSql & ", 1 "
		else
			strSql = strSql & ", 0 "
		end if
		if ForumChkSkipAllowed = 1 then
			if Request.Form("sticky") = 1 then
				strSql = strSql & ", 0 "
			else
				strSql = strSql & ", 1 "
			end if
		else
			strSql = strSql & ", 1 "
		end if
		strSql = strSql & ", 0 "
		strSql = strSql & ", 0 "
        	strSql = strSql & ")"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		if Err.description <> "" then 
			err_Msg = "There was an error = " & Err.description
		else
			err_Msg = "Updated OK"
		end if
                
		strSql = "SELECT Max(TOPIC_ID) as NewTopicID "
		strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
		strSql = strSql & "   and T_AUTHOR = " & rs("MEMBER_ID")
		set rs9 = my_Conn.Execute (strSql)
		NewTopicID = rs9("NewTopicId")
		rs9.close
		set rs9 = nothing

		' DEM --> Do not update forum count if topic is moderated.... Added if and end if
        	if Moderation = "No" then
			'## Forum_SQL - Increase count of topics and replies in Forum table by 1
			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    F_TOPICS = F_TOPICS + 1"
			strSql = strSql & ",    F_COUNT = F_COUNT + 1"
			strSql = strSql & ",    F_LAST_POST_AUTHOR = " & rs("MEMBER_ID") & ""
			strSql = strSql & ",    F_LAST_POST_TOPIC_ID = " & NewTopicID
			strSql = strSql & ",    F_LAST_POST_REPLY_ID = " & 0
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		end if

		ProcessSubscriptions rs("MEMBER_ID"), Cat_ID, Forum_ID, NewTopicID, Moderation

		if Moderation = "No" then
			'## Subscribe checkbox ##
			if request.form("TNotify") <> "" then
				if request.form("TNotify") = "1" then
					AddSubscription "TOPIC", rs("MEMBER_ID"), Cat_ID, Forum_ID, NewTopicID
				elseif request.form("TNotify") = "0" then
					DeleteSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, NewTopicID
				end if
			end if
			'## Subscribe checkbox end ##
		end if

		Go_Result err_Msg, 1
		Response.End
	end if	
end if

if MethodType = "Reply" or MethodType = "ReplyQuote" or MethodType = "TopicQuote" then
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_LASTPOSTDATE, " & strDBNTSQLname
	if strAuthType = "db" then
		strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(strDBNTUserName, "SQLString") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	if strAuthType = "db" then
		strSql = strSql & " AND   M_PASSWORD = '" & ChkString(strPassword, "SQLString") &"'"
		QuoteOk = (ChkQuoteOk(strDBNTUserName) and ChkQuoteOk(strPassword))
	else
		QuoteOk = ChkQuoteOk(strDBNTUserName)
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) or not(ChkQuoteOk(strPassword)) then '##  Invalid Password
		rs.close
		set rs = nothing
		err_Msg = "Invalid Password or User Name"
		Go_Result(err_Msg), 0
		Response.End
	else

		if strPrivateForums = "1" and ForumChkSkipAllowed = 0 then
			if not(chkForumAccess(Forum_ID,MemberID,false)) then
				Go_Result "You are not allowed to post in this forum !", 0
			end if
		end if

		if strFloodCheck = 1 then
			if rs("M_LASTPOSTDATE") > DateToStr(DateAdd("s",strFloodCheckTime,strForumTimeAdjust)) and mLev < 3 then
				strTimeLimit = replace(strFloodCheckTime, "-", "")
				Go_Result "Sorry! We have flood control activated.<br />You cannot post within " & strTimeLimit & " seconds of your last post.<br />Please try again after this period of time elapses.", 0
			end if
		end if

		txtMessage = ChkString(Request.Form("Message"),"message")
		UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		if UserIPAddress = "" then
			UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
		end if
		if txtMessage = " " then
			Go_Result "You must post a message!", 0
			Response.End
		end if

		if strSignatures = "1" and strDSignatures <> "1" then
			if Request.Form("sig") = "yes" and GetSig(strDBNTUserName) <> " " then
				txtMessage = txtMessage & vbNewline & vbNewline & ChkString(GetSig(strDBNTUserName), "signature" )
			end if
		end if

		'## Forum_SQL
		strSql = "INSERT INTO " & strTablePrefix & "REPLY "
		strSql = strSql & "(TOPIC_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", R_AUTHOR"
		strSql = strSql & ", R_DATE "
		if strIPLogging <> "0" then
			strSql = strSql & ", R_IP"
		end if
		strSql = strSql & ", R_STATUS"
		strSql = strSql & ", R_SIG"
		strSql = strSql & ", R_MESSAGE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Topic_ID
		strSql = strSql & ", " & Forum_ID
		strSql = strSql & ", " & Cat_ID
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		if strIPLogging <> "0" then
			strSql = strSql & ", " & "'" & UserIPAddress & "'"
		end if
		' DEM --> Added R_STATUS to allow for moderation of posts
		' Used R_STATUS = 1 to match the topic status code.
		if Moderation = "Yes" then
			strSql = strSql & ", 2"
		else
			strSql = strSql & ", 1"
		end if
		' DEM --> End of Code added
		if Request.Form("sig") = "yes" and strDSignatures = "1" then
		 	strSql = strSql & ", 1 "
		else
			strSql = strSql & ", 0 "
		end if
		strSql = strSql & ", " & "'" & txtMessage & "'"
		strSql = strSql & ")"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		' DEM --> Do not update totals on topics and forums database if post is moderated...Added if and end if
		if Moderation = "No" then
			strSql = "SELECT Max(REPLY_ID) as NewReplyID "
			strSql = strSql & " FROM " & strActivePrefix & "REPLY "
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
			strSql = strSql & "   and R_AUTHOR = " & rs("MEMBER_ID")
			set rs9 = my_Conn.Execute (strSql)
			NewReplyID = rs9("NewReplyID")
			rs9.close
			set rs9 = nothing

			'## Forum_SQL - Update Last Post and count
			strSql = "UPDATE " & strActivePrefix & "TOPICS "
			strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    T_REPLIES = T_REPLIES + 1 "
			strSql = strSql & ",    T_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
			strSql = strSql & ",    T_LAST_POST_REPLY_ID = " & NewReplyID
			if Request.Form("lock") = 1 and ForumChkSkipAllowed = 1 then
				strSql = strSql & ",        T_STATUS = 0 "
			end if
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			'## Subscribe checkbox start ##
			if request.form("TNotify") <> "" then
				if request.form("TNotify") = "1" then
					AddSubscription "TOPIC", rs("MEMBER_ID"), Cat_ID, Forum_ID, Topic_ID
				elseif request.form("TNotify") = "0" then
					DeleteSubscription "TOPIC", MemberID, Cat_ID, Forum_ID, Topic_ID
				end if
			end if
			'## Subscribe checkbox end ##

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    F_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
			strSql = strSql & ",    F_LAST_POST_TOPIC_ID = " & Topic_ID
			strSql = strSql & ",    F_LAST_POST_REPLY_ID = " & NewReplyID
			strSql = strSql & ",    F_COUNT = F_COUNT + 1 "
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		else
			'## Forum_SQL - Update Unmoderated post count
			strSql = "UPDATE " & strActivePrefix & "TOPICS "
			strSql = strSql & " SET T_UREPLIES = T_UREPLIES + 1 "
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		end if

		if Err.description <> "" then
			Go_Result  "There was an error = " & Err.description, 0
			Response.End
		else
			'if Moderation = "No" then
			ProcessSubscriptions rs("MEMBER_ID"), Cat_ID, Forum_ID, Topic_ID, Moderation
			'end if
			Go_Result  "Updated OK", 1
			Response.End
		end if
	end if
end if

if MethodType = "Forum" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorized
			Go_Result "Only an Admin can create a Forum", 0
			Response.End
		case 3 '## Moderator - Not Authorized
			Go_Result "Only an Admin can create a Forum", 0
			Response.End
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtSubject = ChkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New Forum</li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", F_STATUS"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
			if Request.Form("AuthPassword") <> " " then
				strSql = strSql & ", F_PASSWORD_NEW"
			end if
		end if
		'strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE" 
		strSql = strSql & ", F_L_ARCHIVE "
		strSql = strSql & ", F_ARCHIVE_SCHED "
		strSql = strSql & ", F_L_DELETE "
		strSql = strSql & ", F_DELETE_SCHED "
	        strSql = strSql & ", F_SUBSCRIPTION"
	      	strSql = strSql & ", F_MODERATION"
		strSql = strSql & ", F_ORDER "
		strSql = strSql & ", F_DEFAULTDAYS "
		strSql = strSql & ", F_COUNT_M_POSTS "
		strSql = strSql & ") VALUES ("
		strSql = strSql & Cat_ID
		strSql = strSql & ", 1 "
		if strPrivateForums = "1" then
			strSql = strSql & ", " & Request.Form("AuthType") & ""
			if Request.Form("AuthPassword") <> " " then
				strSql = strSql & ", '" & ChkString(Request.Form("AuthPassword"),"SQLString") & "'"
			end if
		end if
		'strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", '' "
		strSql = strSql & ", 30 "
		strSql = strSql & ", '' "
		strSql = strSql & ", 365 "
	        ' DEM --> Start of Code added for moderation and subscription
	        if strSubscription > 0 and CatSubscription > 0 and strEmail = "1" then
	                strSql = strSql & ", " & fSubscription
	        else
	                strSql = strSql & ", 0"
	        end if
	        if strModeration = 1 and CatModeration = 1 then
	                strSql = strSql & ", " & ChkString(Request.Form("Moderation"), "SQLString")
	        else
	                strSql = strSql & ", 0"
	        end if
	        ' DEM --> End of Code added for moderation and subscription
		strSql = strSql & ", 1 "
		strSql = strSql & ", " & ChkString(Request.Form("DefaultDays"), "SQLString")
		strSql = strSql & ", " & ChkString(Request.Form("ForumCntMPosts"), "SQLString")
	        strSql = strSql & ")"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		Else
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")
			newForumModerators rsCount("maxForumId")
			set rsCount = nothing
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "URL" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin can create a web link", 0
			Response.End
		case 3 '## Moderator
			Go_Result "Only an Admin can create a web link", 0
			Response.End
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtAddress = ChkString(Request.Form("Address"),"SQLString")
	txtSubject = ChkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New URL</li>"
	end if
	if txtAddress = " " or lcase(txtAddress) = "http://" or lcase(txtAddress) = "https://" or lcase(txtAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>You Must Enter an Address for the New URL</li>"
	end if
	if (left(lcase(txtAddress), 7) <> "http://" and left(lcase(txtAddress), 8) <> "https://" and left(lcase(txtAddress), 8) <> "file:///") and txtAddress <> "" then
		Err_Msg = Err_Msg & "<li>You Must prefix the Address with <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		strSql = strSql & ", F_STATUS"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
		end if
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_LAST_POST_AUTHOR"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_URL"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE"
		strSql = strSql & ", F_L_ARCHIVE "
		strSql = strSql & ", F_ARCHIVE_SCHED "
		strSql = strSql & ", F_L_DELETE "
		strSql = strSql & ", F_DELETE_SCHED "
	        strSql = strSql & ", F_SUBSCRIPTION, F_MODERATION"
		strSql = strSql & ", F_ORDER "
		strSql = strSql & ", F_DEFAULTDAYS "
		strSql = strSql & ")  VALUES ("
		strSql = strSql & Cat_ID
	        strSql = strSql & ", 1"
		if strPrivateForums = "1" then
			strSql = strSql & ", " & ChkString(Request.Form("AuthType"), "SQLString") & ""
		end if
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & MemberID & " "
		strSql = strSql & ", " & "'" & txtSubject & "'"
		strSql = strSql & ", " & "'" & txtAddress & "'"
		strSql = strSql & ", " & "'" & txtMessage & "'"
		strSql = strSql & ", 1"
	        strSql = strSql & ", ''"
	        strSql = strSql & ", 30"
	        strSql = strSql & ", ''"
	        strSql = strSql & ", 365"
	        ' DEM --> Added 0's for the subscription and moderation fields since they are ignored for URLS
	        strSql = strSql & ", 0, 0"
	        strSql = strSql & ", 1"
		strSql = strSql & ", 30"
		strSql = strSql & ") "

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")                   
			newForumModerators rsCount("maxForumId")                   
			set rsCount = nothing
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "EditForum" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin or a Moderator can change this Forum", 0
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, strDBNTUserName) = "0" then
				Go_Result "Only an Admin or a Moderator can change this Forum", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtSubject = ChkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Forum</li>"
	end if

        if strModeration <> 0 and Request.Form("Moderation") = 0 then
	       	if CheckForUnModeratedPosts("FORUM", Cat_ID, Forum_ID, 0) > 0 then
			Err_Msg = Err_Msg & "<li>Please Approve or Delete all UnModerated/Held posts in this Forum before turning Moderation off</li>"
		end if
	end if

	if Err_Msg = "" then
		'## Forum_SQL - Check if CAT_ID changed
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM " 
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		set rsCatIDCheck = my_Conn.execute(strSql)
		bolCatIDChanged = (cSTr(rsCatIDCheck("CAT_ID")) <> ChkString(Request.Form("Category"), "SQLString"))
		rsCatIDCheck.Close
		set rsCatIDCheck = Nothing

		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & cLng("0" & Request.Form("Category"))
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS = " & cLng("0" & Request.Form("AuthType"))
			if Request.Form("AuthPassword") <> " " then
				strSql = strSql & ", F_PASSWORD_NEW = '" & ChkString(Request.Form("AuthPassword"),"SQLString") & "'"
			end if
		end if
		strSql = strSql & ", F_SUBJECT = '" & txtSubject & "'"
		strSql = strSql & ", F_DESCRIPTION = '" & txtMessage & "'"
		if Request.Form("Moderation") <> "" then
		        strSql = strSql & ",    F_MODERATION = " & cLng("0" & Request.Form("Moderation"))
		end if
		if fSubscription <> "" then
		        strSql = strSql & ",    F_SUBSCRIPTION = " & cLng("0" & fSubscription)
		end if
		strSql = strSql & ",   F_DEFAULTDAYS = " & cLng(Request.Form("DefaultDays"))
		strSql = strSql & ",   F_COUNT_M_POSTS = " & cLng("0" & Request.Form("ForumCntMPosts"))
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		if bolCatIDChanged then
			'## Update category CAT_SUBSCRIPTION/CAT_MODERATION if required
			strSQL = "SELECT " & strTablePrefix & "CATEGORY.CAT_SUBSCRIPTION, " & strTablePrefix & "CATEGORY.CAT_MODERATION FROM " & strTablePrefix & "CATEGORY "
			strSQL = strSQL & " WHERE CAT_ID=" & cLng("0" & Request.Form("Category"))
			set rs = my_conn.execute(strSQL)
			intCatSubs = rs("CAT_SUBSCRIPTION")
			intCatMod = rs("CAT_MODERATION")
			rs.close
			set rs = nothing
			if intCatSubs < fSubscription then
				strSQL = "UPDATE " & strTablePrefix & "CATEGORY SET " & strTablePrefix & "CATEGORY.CAT_SUBSCRIPTION = " & fSubscription
				
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			end if
			if intCatMod = 0 and Request.Form("Moderation") > 0 then
				strSQL = "UPDATE " & strTablePrefix & "CATEGORY SET " & strTablePrefix & "CATEGORY.CAT_MODERATION = " & 1
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			end if 
			'## Forum_SQL - Do DB Update
			strSql = "UPDATE " & strActivePrefix & "TOPICS "
			strSql = strSql & " SET CAT_ID = " & cLng("0" & Request.Form("Category"))
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			'## Forum_SQL - Do DB Update
			strSql = "UPDATE " & strActivePrefix & "REPLY "
			strSql = strSql & " SET CAT_ID = " & cLng("0" & Request.Form("Category"))
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			' DEM --> Added _SUBSCRIPTIONS table
			'## Forum_SQL - Do DB Update
			strSql = "UPDATE " & strTablePrefix & "SUBSCRIPTIONS "
			strSql = strSql & " SET CAT_ID = " & cLng("0" & Request.Form("Category"))
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		end if

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			updateForumMembers Forum_ID
			if mLev = 4 then
				updateForumModerators Forum_ID
			end if
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "EditURL" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin or a Moderator can change this web link", 0
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, strDBNTUserName) = "0" then
				Go_Result "Only an Admin or a Moderator can change this web link", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtAddress = ChkString(Request.Form("Address"),"SQLString")
	txtSubject = ChkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New URL</li>"
	end if
	if txtAddress = " " or lcase(txtAddress) = "http://" or lcase(txtAddress) = "https://" or lcase(txtAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>You Must Enter an Address for the New URL</li>"
	end if
	if (left(lcase(txtAddress), 7) <> "http://" and left(lcase(txtAddress), 8) <> "https://" and left(lcase(txtAddress), 8) <> "file:///") and (txtAddress <> "") then
		Err_Msg = Err_Msg & "<li>You Must prefix the Address with <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if Err_Msg = "" then

		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & cLng("0" & Request.Form("Category"))
		if strPrivateForums = "1" then
			strSql = strSql & ",    F_PRIVATEFORUMS = " & cLng("0" & Request.Form("AuthType"))
		end if
		strSql = strSql & ",    F_SUBJECT = '" & txtSubject & "'"
		strSql = strSql & ",    F_URL = '" & txtAddress & "'"
		strSql = strSql & ",    F_DESCRIPTION = '" & txtMessage & "'"
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			updateForumMembers Forum_ID
			if mLev = 4 then
				updateForumModerators Forum_ID
			end if
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "Category" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member 
		case 0  '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin can create a category", 0
			Response.End
		case 3 '## Moderator
			Go_Result "Only an Admin can create a category", 0
			Response.End
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select
	txtSubject = chkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if

	Err_Msg = ""
	if txtSubject = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New Category</li>"
	end if
	if Err_Msg = "" then

		'## Forum_SQL - Do DB Update
	        ' DEM --> Insert replaced to add subscription and moderation capabilities
	        strSql = "INSERT INTO " & strTablePrefix & "CATEGORY (CAT_NAME, CAT_STATUS, CAT_SUBSCRIPTION, CAT_MODERATION, CAT_ORDER) "
        	strSql = strSql & " VALUES ('" & txtSubject & "'"
       	        strSql = strSql & ", 1"
	        if strSubscription <> 0 and strEmail = "1" then
        	       strSql = strSql & ", " & fSubscription
	        else
                        strSql = strSql & ", 0"
        	end if
	        if strModeration <> 0 then
        	        strSql = strSql & ", " & ChkString(Request.Form("Moderation"), "SQLString")
	        else
        	        strSql = strSql & ", 0"
	        end if
       	        strSql = strSql & ", 1"
        	strSql = strSql & ")"

	        my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			Go_Result  "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if

if MethodType = "EditCategory" then
	member = cLng(ChkUser(strDBNTUserName, strPassword,-1))
	select case Member 
		case 0  '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an Admin can change a category", 0
			Response.End
		case 3 '## Moderator
			'## Do Nothing
			Go_Result "Only an Admin can change a category", 0
			Response.End
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
			Response.End
	end select
	txtSubject = chkString(Request.Form("Subject"),"SQLString")
	if strBadWordFilter = "1" then
		txtSubject = chkString(ChkBadWords(Request.Form("Subject")),"SQLString")
	end if

	Err_Msg = ""
	if txtSubject = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Category</li>"
	end if
        if strModeration <> 0 and Request.Form("Moderation") = 0 then
	        if CheckForUnmoderatedPosts("CAT", Cat_ID, 0, 0) > 0 then
			Err_Msg = Err_Msg & "<li>Please Approve or Delete all UnModerated/Held posts in this Category before turning Moderation off</li>"
		end if
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "CATEGORY "
		strSql = strSql & " SET CAT_NAME = '" & txtSubject & "'"
        	' DEM --> Start of Code added for moderation and subscription functionality
	        if strModeration <> 0 then
        	        strSql = strSql & ",   CAT_MODERATION = " & cLng("0" & Request.Form("Moderation"))
	        end if
	        if strSubscription <> 0 and strEmail = "1" then
	                 strSql = strSql & ",   CAT_SUBSCRIPTION = " & cLng("0" & Request.Form("Subscription"))
        	end if
	        ' DEM --> End of code added for moderation and subscription functionality
		strSql = strSql & " WHERE CAT_ID = " & Cat_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		Application.Lock
		Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
		Application.UnLock

		err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
			Response.End
		else
			Go_Result "Updated OK", 1
		end if
	else 
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
end if
'set rs = nothing
WriteFooter
Response.End

sub Go_Result(str_err_Msg, boolOk)

	select case MethodType
		case "Topic", "TopicQuote", "Reply", "ReplyQuote"
			set rsFCountMP = my_Conn.Execute("SELECT F_COUNT_M_POSTS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & Forum_ID)
			ForumCountMPosts = rsFCountMP("F_COUNT_M_POSTS")
			rsFCountMP.close
			set rsFCountMP = nothing
	end select

	Response.write 	"      <table border=""0"" width=""100%"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td width=""33%"" align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine
	if MethodType = "Topic" or _
		MethodType = "TopicQuote" or _
		MethodType = "Reply" or _
		MethodType = "ReplyQuote" or _
		MethodType = "Edit" or _ 
		MethodType = "EditTopic" then 
			Response.Write	"          " & getCurrentIcon(strIconBar,"","align=""absmiddle""")
			if blnCStatus <> 0 then
				Response.Write	getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""")
			else
				Response.Write	getCurrentIcon(strIconFolderClosed,"","align=""absmiddle""")
			end if
			Response.Write	"&nbsp;<a href=""default.asp?CAT_ID=" & Cat_ID & """>" & ChkString(strCatTitle, "title") & "</a><br />" & vbNewLine
			Response.Write	"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""")
			if blnFStatus <> 0 and blnCStatus <> 0 then
				Response.Write	getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""")
			else
				Response.Write	getCurrentIcon(strIconFolderClosed,"","align=""absmiddle""")
			end if
			Response.Write	"&nbsp;<a href=""forum.asp?FORUM_ID=" & Forum_ID & """>" & ChkString(strForum_Title, "title") & "</a><br />" & vbNewLine
	end if 
	if MethodType = "Reply" or _
		MethodType = "ReplyQuote" or _
		MethodType = "TopicQuote" or _
		MethodType = "Edit" or _ 
		MethodType = "EditTopic" then 
	   		Response.Write "          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""")
	   		if blnTStatus <> 0 and blnFStatus <> 0 and blnCStatus <> 0 then
	   			Response.Write	getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""")
	   		else
	   			Response.Write	getCurrentIcon(strIconFolderClosedTopic,"","align=""absmiddle""")
	   		end if
	   		Response.Write	"&nbsp;<a href=""" & chkString(Request.Form("refer"),"refer") & """>" & ChkString(strTopicTitle,"title") & "</a>" & vbNewLine
	end if 
	Response.write	"          </font></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine
	if boolOk = 1 then 
		Response.write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>"
		select case MethodType
			case "Edit"
			        Response.Write("Your Reply Was Changed Successfully!")
			case "EditCategory"
			        ' DEM --> Added if statement to handle if subscriptions or moderation is allowed
			        if strSubscription > 0 or strModeration > 0 then
			                Response.Write("Category Information Changed Successfully")
			        else
			                Response.Write("Category Name Changed Successfully!")
			        end if
			case "EditForum"
			        Response.Write("FORUM Information Updated Successfully!")
			case "EditTopic"
			        Response.Write("Topic Changed Successfully!")
			case "EditURL"
			        Response.Write("URL Information Updated Successfully!")
			case "Reply", "ReplyQuote", "TopicQuote"
			        ' DEM --> If moderated post, the counts should not be updated until after approval
			        ' Combined the Reply, ReplyQuote and TopicQuote because the basic code was the same.
			        if Moderation = "Yes" then
			                Response.Write("New Reply Posted!  It will appear once approved by a moderator")
			        else
			                Response.Write("New Reply Posted!")
			                DoPCount
					if ForumCountMPosts <> 0 then
						DoUCount MemberID
					end if
			        end if
					DoULastPost MemberID
			case "Topic"
			        ' DEM --> If moderated post, the counts should not be updated until after approval
			        if Moderation = "Yes" then
			                Response.Write("New Topic Posted!  It will appear once approved by a moderator")
			        else
			                Response.Write("New Topic Posted!")
			                DoTCount
			                DoPCount
					if ForumCountMPosts <> 0 then
						DoUCount MemberID
					end if
			        end if
				DoULastPost MemberID
			case "Forum"
			        Response.Write("New Forum Created!")
			case "URL"
			        Response.Write("New URL Created!")
			case "Category"
			        Response.Write("New Category Created!")
			case else
			        Response.Write("Complete!")
			        'DoPCount
			        'DoUCount Request.Form("UserName")
			        'DoULastPost Request.Form("UserName")
		end select
		if MethodType = "Topic" then
			strReturnURL = "topic.asp?TOPIC_ID=" & NewTopicID
			strReturnTxt = "Go to new topic"
		elseif MethodType = "Reply" or MethodType = "ReplyQuote" or MethodType = "TopicQuote" then
			strReturnURL = "topic.asp?whichpage=-1&TOPIC_ID=" & Topic_ID & "&REPLY_ID=" & NewReplyID
			strReturnTxt = "Back to the topic"
		elseif MethodType = "EditTopic" then
			strReturnURL = "topic.asp?TOPIC_ID=" & Topic_ID
			strReturnTxt = "Back to the topic"
		elseif MethodType = "Edit" then
			strReturnURL = "topic.asp?whichpage=-1&TOPIC_ID=" & Topic_ID & "&REPLY_ID=" & Reply_ID
			strReturnTxt = "Back to the topic"
		else
			strReturnURL = chkString(Request.Form("refer"),"refer")
			strReturnTxt = "Back To Forum"
		end if
		Response.write	"</font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=" & strReturnURL & """>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
		select case MethodType
			case "Category"
				Response.Write("Remember to create at least one new forum in this category.")
			case "Forum"
				Response.Write("The new forum is ready for users to begin posting!")
			case "EditForum", "EditCategory"
				Response.Write("Thank you for your contribution!")
			case "URL"
				Response.Write("The new URL is in place!")
			case "EditURL"
				Response.Write("Cheers! Have a nice day!")
			case "Topic", "TopicQuote", "EditTopic", "Reply", "ReplyQuote", "Edit" 
				Response.Write("Thank you for your contribution!")
			case else
				Response.Write("Have a nice day!")
		end select
		Response.write	"</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strReturnURL & """>" & strReturnTxt & "</a></font></p>" & vbNewLine
	else 
		Response.write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There has been a problem!</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & str_err_Msg & "</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go back to correct the problem.</a></font></p>" & vbNewLine
	end if 
	WriteFooter
  	Response.End
end sub

sub newForumMembers(fForumID)
	on error resume next
	if Request.Form("AuthUsers") = "" then
		exit Sub
	end if
	Users = split(Request.Form("AuthUsers"),",")
	for count = Lbound(Users) to Ubound(Users)
		strSql = "INSERT INTO " & strTablePrefix & "ALLOWED_MEMBERS ("
		strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& Users(count) & ", " & fForumID & ")"

		my_conn.execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 then
			Go_REsult err.description, 0
		end if
	next
	on error goto 0
end sub

sub updateForumMembers(fForumID)
	my_Conn.execute ("DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID = " & fForumId),,adCmdText + adExecuteNoRecords
	newForumMembers(fForumID)
end sub

sub newForumModerators(fForumID)
	on error resume next
	if Request.Form("ForumMod") = "" then
		exit Sub
	end if
	Users = split(Request.Form("ForumMod"),",")
	for count = Lbound(Users) to Ubound(Users)
		strSql = "INSERT INTO " & strTablePrefix & "MODERATOR ("
		strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& Users(count) & ", " & fForumID & ")"

		my_conn.execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 then
			Go_REsult err.description, 0
		end if
	next
	on error goto 0
end sub

sub updateForumModerators(fForumID)
	my_Conn.execute ("DELETE FROM " & strTablePrefix & "MODERATOR WHERE FORUM_ID = " & fForumId),,adCmdText + adExecuteNoRecords
	newForumModerators(fForumID)
end sub

sub DoAutoMoveEmail(TopicNum)
	'## Emails Topic Author if Topic Moved.  
	strSql  = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID," & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strActivePrefix & "TOPICS.FORUM_ID, " & strActivePrefix & "TOPICS.T_SUBJECT "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strActivePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strActivePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strActivePrefix & "TOPICS.TOPIC_ID = " & TopicNum

	set rs2 = my_Conn.Execute (strSql)
	
	email = rs2("M_EMAIL")
	user_name = rs2("M_NAME")
	Topic_Title = rs2("T_SUBJECT")
	ForumId = rs2("FORUM_ID")
	Usernum = rs2("MEMBER_ID")
	
	rs2.close
	set rs2 = nothing
	if lcase(strEmail) = "1" then
		strRecipientsName = user_name
		strRecipients = email
		strSubject = strForumTitle & " - Topic Moved"
		strMessage = "Hello " & user_name & vbNewLine & vbNewLine
		strMessage = strMessage & "Your posting on " & strForumTitle & "." & vbNewLine
		strMessage = strMessage & "Regarding the subject - " & Topic_Title & "." & vbNewLine & vbNewLine
		
		if not(chkForumAccess(ForumID,Usernum,false)) then
			strMessage = strMessage & "Has been removed from public display, If you have any questions regarding this, please contact the Administrator of the forum" & vbNewLine
		else
			strMessage = strMessage & "Has been moved to a new forum, You can view it at " & vbNewLine & Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & "topic.asp?TOPIC_ID=" & TopicNum & vbNewLine
		end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
	end if
end sub

'## Subscribe checkbox start ##
sub DeleteSubscription(Level, MemberID, CatID, ForumID, TopicID)
         ' --- Delete the appropriate sublevel of subscriptions
         StrSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS"
         StrSql = StrSql & " WHERE " & strTablePrefix & "SUBSCRIPTIONS.MEMBER_ID = " & MemberID
         if sublevel = "CAT" then
                 StrSql = StrSQL & " AND " & strTablePrefix & "SUBSCRIPTIONS.CAT_ID = " & CatID
         elseif sublevel = "FORUM" then
                 StrSql = StrSQL & " AND " & strTablePrefix & "SUBSCRIPTIONS.FORUM_ID = " & ForumID
         elseif sublevel = "TOPIC" then
                 StrSql = StrSQL & " AND " & strTablePrefix & "SUBSCRIPTIONS.TOPIC_ID = " & TopicID
         end if
         my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

sub AddSubscription(SubLevel, MemberID, CatID, ForumID, TopicID)
         ' --- Insert the appropriate sublevel subscription
         StrSql = "INSERT INTO " & strTablePrefix & "SUBSCRIPTIONS"
         StrSql = StrSql & "(MEMBER_ID, CAT_ID, FORUM_ID, TOPIC_ID) VALUES (" & MemberID & ", "
         if sublevel = "BOARD" then
			StrSql = StrSql & "0, 0, 0)"
		 elseif sublevel = "CAT" then
			StrSql = StrSql & CatID & ", 0, 0)"
         elseif sublevel = "FORUM" then
			StrSql = StrSql & CatID & ", " & ForumID & ", 0)"
         else
			StrSql = StrSql & CatID & ", " & ForumID & ", " & TopicID & ")"
         end if
         my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

'## Subscribe checkbox end ##
%>
