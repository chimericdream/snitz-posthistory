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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<% 
if Request("CAT_ID") <> "" then
	if IsNumeric(Request("CAT_ID")) = True then Cat_ID = cLng(Request("CAT_ID")) else Cat_ID = 0
end if
if Request("FORUM_ID") <> "" then
	if IsNumeric(Request("FORUM_ID")) = True then Forum_ID = cLng(Request("FORUM_ID")) else Forum_ID = 0
end if
if Request("TOPIC_ID") <> "" then
	if IsNumeric(Request("TOPIC_ID")) = True then Topic_ID = cLng(Request("TOPIC_ID")) else Topic_ID = 0
end if
if Request("REPLY_ID") <> "" then
	if IsNumeric(Request("REPLY_ID")) = True then Reply_ID = cLng(Request("REPLY_ID")) else Reply_ID = 0
end if
if Request("MEMBER_ID") <> "" then
	if IsNumeric(Request("MEMBER_ID")) = True then Member_ID = cLng(Request("MEMBER_ID")) else Member_ID = 0
end if

if (Cat_ID + Forum_ID + Topic_ID + Reply_ID + Member_ID) < 1 then
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>The URL has been modified!</b></font></p>" & vbNewLine & _
			"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><b>Possible Hacking Attempt!</b></font></p>" & vbNewLine
	WriteFooterShort
	Response.End
end if

Mode_Type = ChkString(Request("mode"), "SQLString")
strPassword = trim(Request.Form("Pass"))

if request("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
	ArchiveView = "true"
	ArchiveLink = "ARCHIVE=true&"
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
	ArchiveLink = ""
end if

select case Mode_Type
	case "DeleteReply"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser3(strDBNTFUserName, strEncodedPassword, Reply_ID)) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 1) or (mLev = 4) then '## is Allowed
				strSql = "SELECT R_STATUS"
				strSql = strSql & " FROM " & strActivePrefix & "REPLY "
				strSql = strSql & " WHERE REPLY_ID = " & Reply_ID

				set rs = my_Conn.Execute (strSql)

				Reply_Status = rs("R_STATUS")

				rs.close
				set rs = nothing

				'## Forum_SQL - Delete reply
				strSql = "DELETE FROM " & strActivePrefix & "REPLY "
				strSql = strSql & " WHERE REPLY_ID = " & Reply_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				'## Forum_SQL - Delete all past history entries for this reply
				strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
				strSql = strSql & " WHERE R_ID = " & Reply_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				'## Forum_SQL - Get last_post and last_post_author for Topic
				strSql = "SELECT REPLY_ID, R_DATE, R_AUTHOR, R_STATUS"
				strSql = strSql & " FROM " & strActivePrefix & "REPLY "
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID & " "
				strSql = strSql & " AND R_STATUS <= 1 "
				strSql = strSql & " ORDER BY R_DATE DESC"

				set rs = my_Conn.Execute (strSql)

				if not(rs.eof or rs.bof) then
					strLast_Post_Reply_ID = rs("REPLY_ID")
					strLast_Post = rs("R_DATE")
					strLast_Post_Author = rs("R_AUTHOR")
				end if
				if (rs.eof or rs.bof) or IsNull(strLast_Post) or IsNull(strLast_Post_Author) then  'topic has no replies
					set rs2 = Server.CreateObject("ADODB.Recordset")

					'## Forum_SQL - Get post_date and author from Topic
					strSql = "SELECT T_AUTHOR, T_DATE "
					strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
					strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID & " "

					set rs2 = my_Conn.Execute (strSql)

					strLast_Post_Reply_ID = 0
					strLast_Post = rs2("T_DATE")
					strLast_Post_Author = rs2("T_AUTHOR")

					rs2.Close
					set rs2 = nothing

				end if

				rs.Close
				set rs = nothing

				'## FORUM_SQL - Decrease count of replies to individual topic by 1
				'## Only if R_STATUS <= 1

				if Reply_Status <= 1 then
					strSql = "UPDATE " & strActivePrefix & "TOPICS "
					strSql = strSql & " SET T_REPLIES = T_REPLIES - 1 "
					if strLast_Post <> "" then
						strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
						if strLast_Post_Author <> "" then
							strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author & ""
						end if
					end if
					strSql = strSql & ", T_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID & ""
					strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Get last_post and last_post_author for Forum
					strSql = "SELECT TOPIC_ID, T_LAST_POST, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID "
					strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
					strSql = strSql & " ORDER BY T_LAST_POST DESC"

					set rs = my_Conn.Execute (strSql)

					if not rs.eof then
						strLast_Post = rs("T_LAST_POST")
						strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
						strLast_Post_Topic_ID = rs("TOPIC_ID")
						strLast_Post_Reply_ID = rs("T_LAST_POST_REPLY_ID")
					else
						strLast_Post = ""
						strLast_Post_Author = "NULL"
						strLast_Post_Topic_ID = 0
						strLast_Post_Reply_ID = 0
					end if

					rs.Close
					set rs = nothing

					'## Forum_SQL - Decrease count of total replies in Forum by 1
					'## Only if deleted reply wasn't archived

					if ArchiveView = "" then
						strSql =  "UPDATE " & strTablePrefix & "FORUM "
						strSql = strSql & " SET F_COUNT = F_COUNT - 1 "
						strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
						strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & strLast_Post_Topic_ID
						strSql = strSql & ", F_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID
						strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

						'## FORUM_SQL - Decrease count of total replies in Totals table by 1
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET P_COUNT = P_COUNT - 1 "

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					end if
				else
					strSql = "UPDATE " & strActivePrefix & "TOPICS "
					strSql = strSql & " SET T_UREPLIES = T_UREPLIES - 1 "
					strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				end if
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Reply Deleted!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Reply</b></p>" & vbNewLine & _
						"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Reply</b></p>" & vbNewLine & _
					"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if
	case "DeleteTopic"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser5(strDBNTFUserName, strEncodedPassword, Topic_ID))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 1) or (mLev = 4) then
				delAr = split(Topic_ID, ",")
				for i = 0 to ubound(delAr) 

					'## Forum_SQL - count total number of replies of TOPIC_ID  in Reply table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strActivePrefix & "REPLY "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))

					rs.Open strSql, my_Conn
					risposte = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - get topic status so you know if the counts need to be updated
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT T_STATUS "
					strSql = strSql & " FROM " & strActivePrefix & "TOPICS "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))

					rs.Open strSql, my_Conn

					Topic_Status = rs("T_STATUS")

					rs.close
					set rs = nothing

					'## Forum_SQL - Delete the actual topics
					strSql = "DELETE FROM " & strActivePrefix & "TOPICS "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for this topic
					strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
					strSql = strSql & " WHERE T_ID = " & cLng(delAr(i))
	
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for replies to this topic
					strSql = "SELECT REPLY_ID FROM " & strActivePrefix & "REPLY "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE R_ID = " & phRS("REPLY_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all replys related to the topics
					strSql = "DELETE FROM " & strActivePrefix & "REPLY "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete any subscriptions to this topic
					strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
					strSql = strSql & " WHERE TOPIC_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Don't update if topic was in archive
					if (Topic_Status <= 1) and (ArchiveView = "") then
						'## Forum_SQL - Get last_post and last_post_author for Forum
						strSql = "SELECT TOPIC_ID, T_LAST_POST, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID"
						strSql = strSql & " FROM " & strTablePrefix & "TOPICS "			
						strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
						strSql = strSql & " ORDER BY T_LAST_POST DESC"

						set rs = my_Conn.Execute (strSql)

						if not rs.eof then
							rs.movefirst
							strLast_Post = rs("T_LAST_POST")
							strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
							strLast_Post_Topic_ID = rs("TOPIC_ID")
							strLast_Post_Reply_ID = rs("T_LAST_POST_REPLY_ID")
						else
							strLast_Post = ""
							strLast_Post_Author = "NULL"
							strLast_Post_Topic_ID = 0
							strLast_Post_Reply_ID = 0
						end if

						rs.Close
						set rs = nothing

						'## Forum_SQL - Update count of replies to a topic in Forum table
						strSql = "UPDATE " & strTablePrefix & "FORUM "
						strSql = strSql & " SET F_COUNT = F_COUNT - " & cLng(risposte) + 1
						strSql = strSql & ", F_TOPICS = F_TOPICS - " & 1				
						strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
						strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & strLast_Post_Topic_ID
						strSql = strSql & ", F_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID
						strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

						'## Forum_SQL - Update total TOPICS in Totals table
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET T_COUNT = T_COUNT - " & 1
						strSql = strSql & ",    P_COUNT = P_COUNT - " & cLng(risposte) + 1
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					end if
				next
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Topic Deleted!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Topic</b></font><br />" & vbNewLine & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Topic</b></font><br />" & vbNewLine & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case "DeleteForum"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				delAr = split(Forum_ID, ",")
				for i = 0 to ubound(delAr) 
					'## Forum_SQL - Delete all past history entries for replies to topics in this forum
					strSql = "SELECT REPLY_ID FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE R_ID = " & phRS("REPLY_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all replys in this forum
					strSql = "DELETE FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for topics in this forum
					strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE T_ID = " & phRS("TOPIC_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all topics in this forum
					strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for archived replies to topics in this forum
					strSql = "SELECT REPLY_ID FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE R_ID = " & phRS("REPLY_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all archived replys in this forum
					strSql = "DELETE FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for archived topics in this forum
					strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE T_ID = " & phRS("TOPIC_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all archived topics in this forum
					strSql = "DELETE FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete the moderators of this forum
					strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete the Allowed Members of this forum
					strSql = "DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all subscriptions to this forum
					strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete the actual forums
					strSql = "DELETE FROM " & strTablePrefix & "FORUM "
					strSql = strSql & " WHERE FORUM_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - count total number of replies in Reply table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE R_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risreply = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of Topics in Topics table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(TOPIC_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE T_STATUS <= 1 "

					rs.Open strSql, my_Conn
					rispost = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of archived replies in Archived Reply table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE R_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risareply = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of Archived Topics in Archived Topics table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(TOPIC_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE T_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risapost = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - Update total topics and posts in Totals table
					strSql = "UPDATE " & strTablePrefix & "TOTALS "
					strSql = strSql & " SET P_COUNT = " & cLng(risreply + rispost)
					strSql = strSql & ",    T_COUNT = " & cLng(rispost)
					strSql = strSql & ",    P_A_COUNT = " & cLng(risareply + risapost)
					strSql = strSql & ",    T_A_COUNT = " & cLng(risapost)

					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					Application.Lock
					Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
					Application.UnLock
				next
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Forum Deleted!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Forum</b></font><br />" & vbNewLine & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Forum</b></font><br />" & vbNewLine & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case "DeleteCategory"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				delAr = split(Cat_ID, ",")
				for i = 0 to ubound(delAr) 
					'## Forum_SQL - Delete all past history entries for replies to topics in this category
					strSql = "SELECT REPLY_ID FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE R_ID = " & phRS("REPLY_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all replys in this category
					strSql = "DELETE FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for replies to topics in this category
					strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE T_ID = " & phRS("TOPIC_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all topics in this category
					strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for replies to topics in this category
					strSql = "SELECT REPLY_ID FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE R_ID = " & phRS("REPLY_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all archived replys in this category
					strSql = "DELETE FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all past history entries for replies to topics in this category
					strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					Set phRS = Server.CreateObject("ADODB.RecordSet")
					phRS.Open strSql, my_Conn
					Do
						If phRS.EOF or phRS.BOF Then Exit Do
						strSql = "DELETE FROM " & strTablePrefix & "POST_HISTORY "
						strSql = strSql & " WHERE T_ID = " & phRS("TOPIC_ID")

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					Loop
					phRS.Close
					Set phRS = Nothing

					'## Forum_SQL - Delete all archived topics in this category
					strSql = "DELETE FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all moderators and Allowed Members of the forums in this category
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT FORUM_ID "
					strSql = strSql & " FROM " & strTablePrefix & "FORUM "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))

					rs.Open strSql, my_Conn
						do until rs.EOF
							my_Conn.Execute ("DELETE FROM " & strTablePrefix & "MODERATOR WHERE FORUM_ID = " & cLng(rs("FORUM_ID"))),,adCmdText + adExecuteNoRecords
							my_Conn.Execute ("DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID = " & cLng(rs("FORUM_ID"))),,adCmdText + adExecuteNoRecords
							rs.movenext
						loop
					rs.close
					set rs = nothing

					'## Forum_SQL - Delete this Category from any Group Categories
					strSql = "DELETE FROM " & strTablePrefix & "GROUPS "
					strSql = strSql & " WHERE GROUP_CATID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all subscriptions to this Category
					strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete all forums in this category
					strSql = "DELETE FROM " & strTablePrefix & "FORUM "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - Delete the actual category
					strSql = "DELETE FROM " & strTablePrefix & "CATEGORY "
					strSql = strSql & " WHERE CAT_ID = " & cLng(delAr(i))
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					'## Forum_SQL - count total number of replies in Reply table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE R_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risreply = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of Topics in Topics table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(TOPIC_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE T_STATUS <= 1 "

					rs.Open strSql, my_Conn
					rispost = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of archived replies in Archived Reply table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "A_REPLY "
					strSql = strSql & " WHERE R_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risareply = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - count total number of Archived Topics in Archived Topics table
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(TOPIC_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
					strSql = strSql & " WHERE T_STATUS <= 1 "

					rs.Open strSql, my_Conn
					risapost = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - Update total topics and posts in Totals table
					strSql = "UPDATE " & strTablePrefix & "TOTALS "
					strSql = strSql & " SET P_COUNT = " & cLng(risreply + rispost)
					strSql = strSql & ",    T_COUNT = " & cLng(rispost)
					strSql = strSql & ",    P_A_COUNT = " & cLng(risareply + risapost)
					strSql = strSql & ",    T_A_COUNT = " & cLng(risapost)

					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					Application.Lock
					Application(strCookieURL & "JumpBoxChanged")= DateToStr(strForumTimeAdjust)
					Application.UnLock
				next
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Category Deleted!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Category</b></font><br />" & vbNewLine & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete Category</b></font><br />" & vbNewLine & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case "DeleteMember"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				intDeleted = 0
				delAr = split(Member_ID, ",")
				for i = 0 to ubound(delAr) 
					canDelete = cLng(chkCanDelete(MemberID,cLng(delAr(i))))
					if canDelete = 1 then
						'## Forum_SQL - Remove the member from the moderator table
						strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
						strSql = strSql & " WHERE MEMBER_ID = " & cLng(delAr(i))
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

						'## Forum_SQL - Remove any subscriptions this member has in the Subscriptions table
						strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
						strSql = strSql & " WHERE MEMBER_ID = " & cLng(delAr(i))
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

						'## Forum_SQL - Remove the member from the Allowed Members table
						strSql = "DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS "
						strSql = strSql & " WHERE MEMBER_ID = " & cLng(delAr(i))
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

						'## Forum_SQL - Select postcount
						strSql = "SELECT COUNT(T_AUTHOR) AS POSTCOUNT "
						strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
						strSql = strSql & " WHERE T_AUTHOR = " & cLng(delAr(i))

						set rs = my_Conn.Execute (strSql)
						if not rs.eof then
							intPostcount = rs("POSTCOUNT")
						else
							intPostcount = 0
						end if
						rs.close
						set rs = nothing

						'## Forum_SQL - Select postcount
						strSql = "SELECT COUNT(R_AUTHOR) AS REPLYCOUNT "
						strSql = strSql & " FROM " & strTablePrefix & "REPLY "
						strSql = strSql & " WHERE R_AUTHOR = " & cLng(delAr(i))

						set rs = my_Conn.Execute (strSql)
						if not rs.eof then
							intReplycount = rs("REPLYCOUNT")
						else
							intReplycount = 0
						end if

						rs.close
						set rs = nothing							

						'## Forum_SQL - Select Archived postcount
						strSql = "SELECT COUNT(T_AUTHOR) AS POSTCOUNT "
						strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
						strSql = strSql & " WHERE T_AUTHOR = " & cLng(delAr(i))

						set rs = my_Conn.Execute (strSql)
						if not rs.eof then
							intA_Postcount = rs("POSTCOUNT")
						else
							intA_Postcount = 0
						end if

						rs.close
						set rs = nothing

						'## Forum_SQL - Select postcount
						strSql = "SELECT COUNT(R_AUTHOR) AS REPLYCOUNT "
						strSql = strSql & " FROM " & strTablePrefix & "A_REPLY "
						strSql = strSql & " WHERE R_AUTHOR = " & cLng(delAr(i))

						set rs = my_Conn.Execute (strSql)
						if not rs.eof then
							intA_Replycount = rs("REPLYCOUNT")
						else
							intA_Replycount = 0
						end if

						rs.close
						set rs = nothing

						if ((intReplycount + intPostCount + intA_Replycount + intA_PostCount) = 0) then
							'## Forum_SQL - Delete the Member - Member has no posts
							strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS "
							strSql = strSql & " WHERE MEMBER_ID = " & cLng(delAr(i))

							my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
						else
							'## Forum_SQL - disable account - Member has posts, cannot delete just disable account
							strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
							strSql = strSql & " SET M_STATUS = " & 0
							strSql = strSql & ",    M_EMAIL = ' '"
							strSql = strSql & ",    M_LEVEL = " & 1
							strSql = strSql & ",    M_NAME = 'n/a'"
							strSql = strSql & ",    M_COUNTRY = ' '"
							strSql = strSql & ",    M_TITLE = 'deleted'"
							strSql = strSql & ",    M_HOMEPAGE = ' '"
							strSql = strSql & ",    M_AIM = ' '"
							strSql = strSql & ",    M_ICQ = ' '"
							strSql = strSql & ",    M_MSN = ' '"
							strSql = strSql & ",    M_YAHOO = ' '"
							strSql = strSql & " WHERE MEMBER_ID = " & cLng(delAr(i))

							my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
						end if

						'## Forum_SQL - Update total of Members in Totals table
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET U_COUNT = U_COUNT - " & 1

						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
						intDeleted = intDeleted + 1
					end if
				next
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>"
				if intDeleted > 0 then
					Response.Write("<b>Member Deleted!</b>")
				else
					Response.Write("<b>No Members Deleted!</b>")
				end if
				Response.Write	"</font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete a Member</b></font><br />" & vbNewLine & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Delete a Member</b></font><br />" & vbNewLine & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case else
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Delete "
		select case Mode_Type
			case "Member"
				Response.Write("Member")
			case "Category"
				Response.Write("Category")
			case "Forum"
				Response.Write("Forum")
			case "Topic"
				Response.Write("Topic")
			case "Reply"
				Response.Write("Reply")
		end select
		Response.Write	"</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><font color=""" & strHiLiteFontColor & """>NOTE:&nbsp;</font></b>"
		select case Mode_Type
			case "Member"
				Response.Write("Only Administrators can delete a Member.")
			case "Category"
				Response.Write("Only Administrators can delete a Category.")
			case "Forum"
				Response.Write("Only Administrators can delete Forums.")
			case "Topic"
				Response.Write("Only Moderators and Administrators, or the Author of a Topic (if no Replies have been made to it) can delete Topics.")
			case "Reply"
				Response.Write("Only the Author, Moderators and Administrators can delete Replies.")
		end select
		Response.Write	"</font></p>" & vbNewLine & _
				"      <form action=""pop_delete.asp?mode="
		select case Mode_Type
			case "Member"
				Response.Write("DeleteMember")
			case "Category"
				Response.Write("DeleteCategory")
			case "Forum"
				Response.Write("DeleteForum")
			case "Topic"
				Response.Write("DeleteTopic")
			case "Reply"
				Response.Write("DeleteReply")
		end select
		Response.Write	""" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"      <input type=""hidden"" name=""ARCHIVE"" value=""" & ArchiveView & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""REPLY_ID"" value=""" & Reply_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""TOPIC_ID"" value=""" & Topic_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""FORUM_ID"" value=""" & Forum_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""CAT_ID"" value=""" & Cat_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""MEMBER_ID"" value=""" & Member_ID & """>" & vbNewLine & _
				"      <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine
		if strAuthType = "db" then
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:</font></b></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" maxLength=""25"" name=""User"" value=""" & chkString(strDBNTUserName,"display") & """ style=""width:150px;""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:</font></b></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """><input type=""Password"" maxLength=""25"" name=""Pass"" value="""" style=""width:150px;""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		else
			if strAuthType = "nt" then
				Response.Write	"              <tr>" & vbNewLine & _
      						"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>NT Account:</font></b></td>" & vbNewLine & _
						"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & strDBNTUserName & "</font></td>" & vbNewLine & _
						"              </tr>" & vbNewLine
			end if
		end if
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      </form>" & vbNewLine
end select
WriteFooterShort
Response.End

function chkUser5(fName, fPassword, fTopic)
	'## Forum_SQL
	strSql = "SELECT M.MEMBER_ID, M.M_LEVEL, M.M_NAME, M.M_PASSWORD, T.T_AUTHOR, T.T_REPLIES "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "TOPICS T "
	StrSql = strSql & " WHERE M." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then
		strSql = strSql & " AND M.M_PASSWORD = '" & fPassword &"' "
	End If
	strSql = strSql & " AND T.TOPIC_ID = " & fTopic
	strSql = strSql & " AND M.M_STATUS = " & 1
 
	set rsCheck = my_Conn.Execute (strSql)
 
	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		chkUser5 = 0 '## Invalid Password
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(rsCheck("T_AUTHOR")) and cLng(rsCheck("T_REPLIES")) < 1 then 
			chkUser5 = 1 '## Author
		else
			Select case cLng(rsCheck("M_LEVEL"))
				case 1
					chkUser5 = 2 '## Normal User
				case 2
					chkUser5 = 3 '## Moderator
				case 3
					chkUser5 = 4 '## Admin
				case else
					chkUser5 = cLng(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if
 
	rsCheck.close	
	set rsCheck = nothing
end function

function chkUser3(fName, fPassword, fReply)
	'## Forum_SQL
	strSql = "SELECT M.MEMBER_ID, M.M_LEVEL, M.M_NAME, M.M_PASSWORD, R.R_AUTHOR "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "REPLY R "
	StrSql = strSql & " WHERE M." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then	
		strSql = strSql & " AND M.M_PASSWORD = '" & fPassword &"' "
	End If
	strSql = strSql & " AND R.REPLY_ID = " & fReply
	strSql = strSql & " AND M.M_STATUS = " & 1

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		chkUser3 = 0 '## Invalid Password
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(rsCheck("R_AUTHOR")) then 
			chkUser3 = 1 '## Author
		else
			Select case cLng(rsCheck("M_LEVEL"))
				case 1
					chkUser3 = 2 '## Normal User
				case 2
					chkUser3 = 3 '## Moderator
				case 3
					chkUser3 = 4 '## Admin
				case else
					chkUser3 = cLng(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing
end function

function chkCanDelete(fAM_ID, fM_ID)
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	StrSql = strSql & " WHERE MEMBER_ID = " & fM_ID

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF then
		chkCanDelete = 0 '## No Members Found
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(fAM_ID) then 
			chkCanDelete = 0 '## Can't delete self
		else
			Select case cLng(rsCheck("M_LEVEL"))
				case 1
					chkCanDelete = 1 '## Can delete Normal User
				case 2
					chkCanDelete = 1 '## Can delete Moderator
				case 3
					if fAM_ID <> intAdminMemberID then
						chkCanDelete = 0 '## Only the Forum Admin can delete other Administrators
					else
						chkCanDelete = 1 '## Forum Admin is ok to delete other Administrators
					end if
				case else
					chkCanDelete = 0 '## Member doesn't have a Member Level?
			End Select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing
end function
%>
