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
<%
If mlev <> 4 and mlev <> 3 Then
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>Sorry!</font></p>" & vbNewLine & _
			"      <table align=""center"" border=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>You do not have permission to access this page.</font></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strForumURL & """>Back to the forum.</a></font></p><br />" & vbNewLine
Else
	if request("ARCHIVE") = "true" then
		strActivePrefix = strTablePrefix & "A_"
	else
		strActivePrefix = strTablePrefix
	end if

	strSQL = "SELECT * FROM " & strTablePrefix & "POST_HISTORY WHERE "

	If Request.QueryString("T_ID") <> "" Then
		T_ID = cLng(Request.QueryString("T_ID"))
		strSQL = strSQL & "T_ID = " & T_ID
	End If
	If Request.QueryString("R_ID") <> "" Then
		R_ID = cLng(Request.QueryString("R_ID"))
		strSQL = strSQL & "R_ID = " & R_ID
	End If

	strSQL = strSQL & " ORDER BY P_DATE ASC"
	Set phRS = Server.CreateObject("ADODB.RecordSet")
	phRS.Open strSQL, my_Conn

	If phRS.EOF or phRS.BOF Then
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There was an error.</font></p>" & vbNewLine & _
			"      <table align=""center"" border=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>There are no history entries for the post you selected.</font></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strForumURL & """>Back to the forum.</a></font></p><br />" & vbNewLine
	Else
		Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
				"          " & getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
				"          " & getCurrentIcon(strIconBar,"","align=""absmiddle""") & vbNewline & _
				"&nbsp;Post History<br />" & vbNewLine & _
				"</font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
		Response.Write	"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine & _
	  			"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Post History</font></b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine

		arrPostData = phRS.GetRows(adGetRowsRest)
		iHistCount = UBound(arrPostData, 2)

		intI = 0

		phH_ID = 0
		phT_ID = 1
		phT_SUBJECT = 2
		phR_ID = 3
		phP_MESSAGE = 4
		phP_AUTHOR = 5
		phP_LAST_EDITBY = 6
		phP_DATE = 7

		for iForum = 0 to iHistCount
			phHistory_ID = arrPostData(phH_ID, iForum)
			phHistory_TopicID = arrPostData(phT_ID, iForum)
			phHistory_TopicSub = arrPostData(phT_SUBJECT, iForum)
			phHistory_ReplyID = arrPostData(phR_ID, iForum)
			phHistory_Message = arrPostData(phP_MESSAGE, iForum)
			phHistory_Author = arrPostData(phP_AUTHOR, iForum)
			phHistory_LastEditBy = arrPostData(phP_LAST_EDITBY, iForum)
			phHistory_Date = arrPostData(phP_DATE, iForum)

			if intI = 0 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if

			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgcolor=""" & CColor & """ height=""100%"" valign=""top"">" & vbNewLine & _
					"                  <table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
					"                    <tr>" & vbNewLine & _
					"                      <td valign=""top"">" & vbNewLine
			Response.Write  "                      " & getCurrentIcon(strIconPosticon,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>Posted&nbsp;-&nbsp;" & ChkDate(phHistory_Date, "&nbsp;:&nbsp;" ,true) & "</font>" & vbNewline
			Response.Write	"                      <hr noshade size=""" & strFooterFontSize & """></td>" & vbNewLine & _
					"                    </tr>" & vbNewLine & _
					"                    <tr>" & vbNewLine & _
					"                      <td valign=""top"" height=""100%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"" id=""msg"">"
			Response.Write	formatStr(phHistory_Message)
			Response.Write	"</span id=""msg""></font></td>" & vbNewLine & _
					"                    </tr>" & vbNewLine
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

		If Request.QueryString("T_ID") <> "" Then
			T_ID = cLng(Request.QueryString("T_ID"))
			strSQL = "SELECT TOPIC_ID, T_SUBJECT, T_MESSAGE AS phMsg, T_AUTHOR AS phAuth, T_LAST_EDITBY AS phLastEditBy, T_DATE AS phDate FROM " & strActivePrefix & "TOPICS WHERE TOPIC_ID = " & T_ID
		End If
		If Request.QueryString("R_ID") <> "" Then
			R_ID = cLng(Request.QueryString("R_ID"))
			strSQL = "SELECT REPLY_ID, R_AUTHOR AS phAuth, R_MESSAGE AS phMsg, R_LAST_EDITBY AS phLastEditBy, R_DATE AS phDate FROM " & strActivePrefix & "REPLY WHERE REPLY_ID = " & R_ID
		End If

		Set phRS2 = Server.CreateObject("ADODB.RecordSet")
		phRS2.Open strSQL, my_Conn

		If Request.QueryString("T_ID") <> "" Then
			phHistory_TopicID = phRS2("TOPIC_ID")
			phHistory_TopicSub = phRS2("T_SUBJECT")
		End If
		If Request.QueryString("R_ID") <> "" Then
			phHistory_ReplyID = phRS2("REPLY_ID")
		End If
		phHistory_Message = phRS2("phMsg")
		phHistory_Author = phRS2("phAuth")
		phHistory_LastEditBy = phRS2("phLastEditBy")
		phHistory_Date = phRS2("phDate")

		phRS2.Close
		Set phRS2 = Nothing

		if intI = 0 then
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if

		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ height=""100%"" valign=""top"">" & vbNewLine & _
				"                  <table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"                    <tr>" & vbNewLine & _
				"                      <td valign=""top""><b>Current Version</b><br />" & vbNewLine
		Response.Write  "                      " & getCurrentIcon(strIconPosticon,"","hspace=""3""") & "<font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize  & """>Posted&nbsp;-&nbsp;" & ChkDate(phHistory_Date, "&nbsp;:&nbsp;" ,true) & "</font>" & vbNewline
		Response.Write	"                      <hr noshade size=""" & strFooterFontSize & """></td>" & vbNewLine & _
				"                    </tr>" & vbNewLine & _
				"                    <tr>" & vbNewLine & _
				"                      <td valign=""top"" height=""100%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"" id=""msg"">"
		Response.Write	formatStr(phHistory_Message)
		Response.Write	"</span id=""msg""></font></td>" & vbNewLine & _
				"                    </tr>" & vbNewLine
		Response.Write	"                    <tr>" & vbNewLine & _
				"                      <td valign=""bottom"" align=""right"" height=""20""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go to Top of Page","align=""right""") & "</a></td>" & vbNewLine & _
				"                    </tr>" & vbNewLine & _
				"                  </table>" & vbNewLine & _
				"                </td>" & vbNewLine & _
				"              </tr>" & vbNewLine

		Response.write	"			</table>" & vbNewline & _
				"				</td>" & vbNewline & _
				"			</tr>" & vbNewline & _
				"		</table>" & vbNewline & _
				"	</td>" & vbNewline & _
				"</tr>" & vbNewline & _
				"</table><br />" & vbNewline
	End If

	phRS.Close
	Set phRS = Nothing
End If

WriteFooter
%>