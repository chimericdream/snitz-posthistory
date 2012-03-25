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

'## Const variable_name = "icon_filename|width|height"

Const strIconAIM = "icon_aim.gif|15|15"
Const strIconBar = "icon_bar.gif|15|15"
Const strIconBlank = "icon_blank.gif|15|15"
Const strIconClosedTopic = "icon_closed_topic.gif|15|15"
Const strIconDeleteReply = "icon_delete_reply.gif|15|15"
Const strIconEditTopic = "icon_edit_topic.gif|15|15"
Const strIconEditorBold = "icon_editor_bold.gif|23|22"
Const strIconEditorCenter = "icon_editor_center.gif|23|22"
Const strIconEditorCode = "icon_editor_code.gif|23|22"
Const strIconEditorEmail = "icon_editor_email.gif|23|22"
Const strIconEditorHR = "icon_editor_hr.gif|23|22"
Const strIconEditorLeft = "icon_editor_left.gif|23|22"
Const strIconEditorImage = "icon_editor_image.gif|23|22"
Const strIconEditorItalicize = "icon_editor_italicize.gif|23|22"
Const strIconEditorList = "icon_editor_list.gif|23|22"
Const strIconEditorQuote = "icon_editor_quote.gif|23|22"
Const strIconEditorRight = "icon_editor_right.gif|23|22"
Const strIconEditorSmilie = "icon_editor_smilie.gif|23|22"
Const strIconEditorStrike = "icon_editor_strike.gif|23|22"
Const strIconEditorUnderline = "icon_editor_underline.gif|23|22"
Const strIconEditorUrl = "icon_editor_url.gif|23|22"
Const strIconEmail = "icon_email.gif|15|15"
Const strIconFolder = "icon_folder.gif|15|15"
Const strIconFolderArchive = "icon_folder_archive.gif|16|16"
Const strIconFolderArchived = "icon_folder_archived.gif|15|15"
Const strIconFolderClosed = "icon_folder_closed.gif|15|15"
Const strIconFolderClosedTopic = "icon_folder_closed_topic.gif|15|15"
Const strIconFolderDelete = "icon_folder_delete.gif|15|15"
Const strIconFolderHold = "icon_folder_hold.gif|15|15"
Const strIconFolderHot = "icon_folder_hot.gif|15|17"
Const strIconFolderLocked = "icon_folder_locked.gif|15|15"
Const strIconFolderModerate = "icon_folder_moderate.gif|15|15"
Const strIconFolderNew = "icon_folder_new.gif|15|15"
Const strIconFolderNewHot = "icon_folder_new_hot.gif|15|17"
Const strIconFolderNewLocked = "icon_folder_new_locked.gif|15|15"
Const strIconFolderNewSticky = "icon_folder_new_sticky.gif|15|15"
Const strIconFolderNewStickyLocked = "icon_folder_new_sticky_locked.gif|15|15"
Const strIconFolderNewTopic = "icon_folder_new_topic.gif|15|15"
Const strIconFolderOpen = "icon_folder_open.gif|15|15"
Const strIconFolderOpenTopic = "icon_folder_open_topic.gif|15|15"
Const strIconFolderPencil = "icon_folder_pencil.gif|15|15"
Const strIconFolderSticky = "icon_folder_sticky.gif|15|15"
Const strIconFolderStickyLocked = "icon_folder_sticky_locked.gif|15|15"
Const strIconFolderUnlocked = "icon_folder_unlocked.gif|15|15"
Const strIconFolderUnmoderated = "icon_folder_unmoderated.gif|15|15"
Const strIconGoDown = "icon_go_down.gif|15|15"
Const strIconGoLeft = "icon_go_left.gif|15|15"
Const strIconGoRight = "icon_go_right.gif|15|15"
Const strIconGoUp = "icon_go_up.gif|15|15"
Const strIconGroup = "icon_group.gif|15|15"
Const strIconGroupCategories = "icon_group_categories.gif|21|22"
Const strIconHomepage = "icon_homepage.gif|15|15"
Const strIconICQ = "icon_icq.gif|15|15"
Const strIconIP = "icon_ip.gif|15|15"
Const strIconLastpost = "icon_lastpost.gif|12|10"
Const strIconLock = "icon_lock.gif|12|12"
Const strIconMinus = "icon_minus.gif|10|10"
Const strIconMSNM = "icon_msnm.gif|15|15"
Const strIconPencil = "icon_pencil.gif|12|12"
Const strIconPhotoNone = "icon_photo_none.gif|150|150"
Const strIconPlus = "icon_plus.gif|10|10"
Const strIconPosticon = "icon_posticon.gif|15|15"
Const strIconPosticonHold = "icon_posticon_hold.gif|15|15"
Const strIconPosticonUnmoderated = "icon_posticon_unmoderated.gif|15|15"
Const strIconPrint = "icon_print.gif|16|17"
Const strIconPrivateAdd = "icon_private_add.gif|23|22"
Const strIconPrivateAddAll = "icon_private_addall.gif|23|22"
Const strIconPrivateRemAll = "icon_private_remall.gif|23|22"
Const strIconPrivateRemove = "icon_private_remove.gif|23|22"
Const strIconProfile = "icon_profile.gif|15|15"
Const strIconProfileLocked = "icon_profile_locked.gif|15|15"
Const strIconReplyTopic = "icon_reply_topic.gif|15|15"
Const strIconSendTopic = "icon_send_topic.gif|15|15"
Const strIconSmile = "icon_smile.gif|15|15"
Const strIconSmile8ball = "icon_smile_8ball.gif|15|15"
Const strIconSmileAngry = "icon_smile_angry.gif|15|15"
Const strIconSmileApprove = "icon_smile_approve.gif|15|15"
Const strIconSmileBig = "icon_smile_big.gif|15|15"
Const strIconSmileBlackeye = "icon_smile_blackeye.gif|15|15"
Const strIconSmileBlush = "icon_smile_blush.gif|15|15"
Const strIconSmileClown = "icon_smile_clown.gif|15|15"
Const strIconSmileCool = "icon_smile_cool.gif|15|15"
Const strIconSmileDead = "icon_smile_dead.gif|15|15"
Const strIconSmileDisapprove = "icon_smile_disapprove.gif|15|15"
Const strIconSmileEvil = "icon_smile_evil.gif|15|15"
Const strIconSmileKisses = "icon_smile_kisses.gif|15|15"
Const strIconSmileQuestion = "icon_smile_question.gif|15|15"
Const strIconSmileSad = "icon_smile_sad.gif|15|15"
Const strIconSmileShock = "icon_smile_shock.gif|15|15"
Const strIconSmileShy = "icon_smile_shy.gif|15|15"
Const strIconSmileSleepy = "icon_smile_sleepy.gif|15|15"
Const strIconSmileTongue = "icon_smile_tongue.gif|15|15"
Const strIconSmileWink = "icon_smile_wink.gif|15|15"
Const strIconSort = "icon_sort.gif|15|15"
Const strIconStarBlue = "icon_star_blue.gif|13|12"
Const strIconStarBronze = "icon_star_bronze.gif|13|12"
Const strIconStarCyan = "icon_star_cyan.gif|13|12"
Const strIconStarGold = "icon_star_gold.gif|13|12"
Const strIconStarGreen = "icon_star_green.gif|13|12"
Const strIconStarOrange = "icon_star_orange.gif|13|12"
Const strIconStarPurple = "icon_star_purple.gif|13|12"
Const strIconStarRed = "icon_star_red.gif|13|12"
Const strIconStarSilver = "icon_star_silver.gif|13|12"
Const strIconSubscribe = "icon_subscribe.gif|15|15"
Const strIconTopicAllRead = "icon_topic_all_read.gif|15|15"
Const strIconTrashcan = "icon_trashcan.gif|12|12"
Const strIconUnlock = "icon_unlock.gif|12|12"
Const strIconUnsubscribe = "icon_unsubscribe.gif|15|15"
Const strIconUrl = "icon_url.gif|16|16"
Const strIconYahoo = "icon_yahoo.gif|16|15"
Const strIconPostHistory = "icon_post_history.gif|16|16"

function getCurrentIcon(fIconName,fAltText,fOtherTags)
	if fIconName = "" then exit function
	if fOtherTags <> "" then fOtherTags = " " & fOtherTags
	if Instr(fIconName,"http://") > 0 then strTempImageUrl = "" else strTempImageUrl = strImageUrl
	tmpicons = split(fIconName,"|")
	if tmpicons(1) <> "" then fWidth = " width=""" & tmpicons(1) & """"
	if tmpicons(2) <> "" then fHeight = " height=""" & tmpicons(2) & """"
	getCurrentIcon = "<img src=""" & strTempImageUrl & tmpicons(0) & """" & fWidth & fHeight & " border=""0"" alt=""" & fAltText & """ title=""" & fAltText & """" & fOtherTags & " />"
end function
%>
