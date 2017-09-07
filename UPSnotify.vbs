'################################## SCRIPT HEADER #########################################
'##  This script will take UPS shipping notifications and take selected information,
'##    from them and place them into a template email and save that email as a draft.
'## 
'##
'## This will be broken into to parts:   Select unread emails from a folder in Outlook
'##                                      Parse the email for selected information                                    					                        
'##                                      Save template (w/ info from UPS) in drafts folder
'##                                                 in Outlook										
'##
'##  <variable_name>     = Reference in a comment to a variable in code
'##  '       = Line of code not approved for live use
'##  '##     = Comment
'##  '####   = Debugger
'##  '!!!!!! = Update
'##                             [Date Edited: 3-21-2014]
'##                            [Edited By: Michael Kracke]
'##                                   [Version 3.0]
'#################################### END HEADER ##########################################
'------------------------------------- Version 3.0 ----------------------------------------
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!                            UPDATES FROM PREVIOUS VERSION
'!  Changes in Code
'!  - Added possible call instructions to Function getEmail
'!  - Updated Function sendEmail to include changes from ^^
'!  - Added <soNum> to the subject line 
'! 
'!  Changes in Formatting
'!  - Added comments and spacing for easy reading   
'! 
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'------------------------------------------------------------------------------------------
'################################# INSTRUCTIONS ###########################################
'##            Below is a brief set of instructions on how to use this script
'## 
'##     This script is written to run regardless of whose account the script is run on.
'##    For the script to run correctly a structure of folders must be set and followed.
'##    The follows folders and locations must exist for the script to run:
'##   - This script should live in C:\Scripts\NESNS\UPSnotify.vbs
'##   - Outlook should have the folder structure "Inbox\Shipped\UPS Tracking"
'##        This should be where a rule places all UPS Tracking emails 
'##                     (NOT DELIEVERY NOTIFICATIONS)
'##   - Outlook should also contain the folder structure "Drafts\UPS Notifications"
'##        This is to separate the drafts create by this script from any other drafts 
'############################### END OF INSTRUCTIONS ######################################
'------------------------------------------------------------------------------------------
'################################# START OF SCRIPT ########################################
'## Connect to Outlook and Excel
'## Setup current users outlook folder system
DIM objApp, objWbs, objWorkbook, solSheet
SET objApp = CreateObject("Excel.Application")
SET objWbs = objApp.WorkBooks
objApp.Visible = FALSE
SET objWorkbook = objWbs.Open("\\milkyway\Sourcebook\_Order Processing\Daily Dimensions Sourcebook Orders - Morton - NatExplr Company.xlsx")
SET solSheet = objWorkbook.Sheets("Sales Order Line")
CONST olFolderInbox = 6
CONST olFolderDrafts = 16
CONST olTxt = 0
SET objOutlook = CreateObject("Outlook.Application")
SET objNamespace = objOutlook.GetNamespace("MAPI")
SET objEmail = objOutlook.CreateItem(olMailItem)
SET objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
'## Next 2 lines are specific to the current user
SET objShipFolder = objFolder.Folders("Ed Merch")
SET objUPSFolder = objShipFolder.Folders("UPS Quantum")
SET allEmails = objUPSFolder.Items
DIM email, emailAdd, shipTo


'## Snarfing information from <objUPSFolder> to add to email
FOR EACH email IN objUPSFolder.Items
ON ERROR RESUME NEXT
	IF email.Unread = TRUE THEN
		'## Get # of packages
		DIM Reg1, pack, packNum, catch1
		SET Reg1 = NEW RegExp
		WITH Reg1 
			.Pattern = "Number of Packages:\s[0-9]*"
			.IgnoreCase = TRUE
			.Global = TRUE
		END WITH
				IF Reg1.test(email.body) THEN
			SET pack = Reg1.Execute(email.body)
			FOR EACH catch1 IN pack
				packNum = Split(catch1, ":")
			NEXT	
		END IF
		'## Obtain the SO # 
		DIM Reg2, ref, refNum, catch2
		SET Reg2 = NEW RegExp
		WITH Reg2 
			.Pattern = "Reference Number 1:(\s(\w)*)*\n"
			.IgnoreCase = TRUE
			.Global = TRUE
		END WITH
		IF Reg2.test(email.body) THEN
			SET ref = Reg2.Execute(email.body)
			FOR EACH catch2 IN ref
				refNum = Split(UCase(catch2), ":")
			NEXT
		END IF	
		'## Cut order number out of <refNum>
		DIM Reg3, so, soNum, catch3
		SET Reg3 = NEW RegExp
		WITH Reg3 
			.Pattern = "SO[0-9]{8}"
			.IgnoreCase = TRUE
			.Global = TRUE
		END WITH
		IF Reg3.test(refNum(1)) THEN
			SET so = Reg3.Execute(refNum(1))
			FOR EACH catch3 IN so
				soNum = UCASE(catch3)
			NEXT
		END IF
		'## Snarf the shipping number
		DIM Reg4, sh, shNum, catch4
		SET Reg4 = NEW RegExp
		WITH Reg4 
			.Pattern = "SH[0-9]{9}"
			.IgnoreCase = TRUE
			.Global = TRUE
		END WITH
		IF Reg3.test(email.body) THEN
			SET sh = Reg4.Execute(email.body)
			FOR EACH catch4 IN sh
				shNum = UCASE(catch4)
			NEXT
		END IF
		'## Find order in Excel sheet by looking for <soNum>
		DIM i, flip, row 
		flip = true
		i = 5
		DO WHILE flip = TRUE
			IF solSheet.Range("A" & i).value = soNum THEN
				row = solSheet.Range("A" & i).row
				flip = false
			End IF
			IF solSheet.Range("A" & i).value = "" THEN
				flip = false
				row = solSheet.Range("A" & i).row
			END IF
			i = i + 1
		LOOP
		'## Snarf tracking #
		DIM Reg5, track, trackNum, catch5
		SET Reg5 = NEW RegExp
		WITH Reg5 
			.Pattern = "1Z(.{16})"
			.IgnoreCase = TRUE
			.Global = TRUE
		END WITH
		IF Reg5.test(email.body) THEN
			SET track = Reg5.Execute(email.body)
			FOR EACH catch5 IN track
				trackNum = catch5
			NEXT
		END IF
		'## Change Location code to a string
		SELECT CASE loc
			CASE loc = InStr(refNum(1), "NE") 
				locCode = "MAIN"		
			CASE loc = InStr(refNum(1), "AW")
				locCode = "AKERWOODS"
			CASE loc = InStr(refNum(1), "BA")
				locCode = "BEINGART"
			CASE loc = InStr(refNum(1), "CCW")
				locCode = "CEDARCREEK"
			CASE loc = InStr(refNum(1), "FN")
				locCode = "FREENOTES"
			CASE loc = InStr(refNum(1), "FNHP")
				locCode = "FREENOTEHP"
			CASE loc = InStr(refNum(1), "GUSA")
				locCode = "GARDMAN"
			CASE loc = InStr(refNum(1), "GP")
				locCode = "GREATPLAIN"
			CASE loc = InStr(refNum(1), "KK")
				locCode = "CHALKKODO"
			CASE loc = InStr(refNum(1), "NY")
				locCode = "NATURALYD"
			CASE loc = InStr(refNum(1), "SC")
				locCode = "SOYCLEAN"
			CASE loc = InStr(refNum(1), "RH")
				locCode = "RUSTICHOME"
			CASE loc = InStr(refNum(1), "SD")
				locCode = "SLAPDRUM"
			CASE loc = InStr(refNum(1), "WC")
				locCode = "WOODCNTRY"
			CASE loc = InStr(refNum(1), "CB")
				locCode = "CHILDBRITE"
			CASE loc = InStr(refNum(1), "RG")
				locCode = "MAIN"
		END SELECT
		'## Comment out during testing
		email.Unread = FALSE
		'## Call Functions
		emailAdd = solSheet.Range("M" & i).value
		getEmail soNum, emailAdd, phone
		ShipAddress shipTo
		ProductList row, item, pending, locCode, soNum, shNum
		SendEmail packNum, refNum, shipTo, trackNum, soNum, emailAdd, phone, item, pending
		'## Clear variables
		packNum = " "
		refNum = " "
		shipTo = " "
		trackNum = " "
		emailAdd = " "
		phone = " "
		locCode = " "
		item = " "
		pending = " "
	'## end of IF email.Unread
	END IF
'## Get next email
NEXT

'## Get the email address, possible special email addresses & phone #'s
FUNCTION getEmail(soNum, emailAdd, phone)
SET objWorkbook = objWbs.Open("\\milkyway\Sourcebook\_Order Processing\Special Shipment Tracking Notifications.xlsx")
SET emailSheet = objWorkbook.Sheets("Sheet1")
emailAdd = solSheet.Range("M" & i).value
DIM n, theEnd, tmp
		theEnd = true
		n = 3
		DO WHILE theEnd = TRUE
			IF emailSheet.Range("A" & n).value = soNum THEN
				emailAdd = emailAdd & "; " & emailSheet.Range("B" & n).value
				phone = emailSheet.Range("C" & n).value
			End IF
			IF emailSheet.Range("A" & n).value = "" THEN
				theEnd = false
			END IF
			n = n + 1
		LOOP
'## End Function getEmail
END FUNCTION

'## Add the shipping address from Excel
FUNCTION ShipAddress(shipTo)
	IF NOT solSheet.Range("AH" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AH" & i - 1).value & "<br />&emsp;"
	END IF
	IF NOT solSheet.Range("AI" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AI" & i - 1).value & "<br />&emsp;"
	END IF
	IF NOT solSheet.Range("AJ" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AJ" & i - 1).value & "<br />&emsp;"
	END IF
	IF NOT solSheet.Range("AK" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AK" & i - 1).value & "<br />&emsp;"
	END IF 
	IF NOT solSheet.Range("AL" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AL" & i - 1).value & "<br />&emsp;"
	END IF
	IF NOT solSheet.Range("AM" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AM" & i - 1).value & "<br />&emsp;"
	END IF
	IF NOT solSheet.Range("AN" & i - 1).value = "" THEN
		shipTo = shipTo & solSheet.Range("AN" & i - 1).value  & "<br />"
	END IF
'## End Function ShipAdresses
END FUNCTION

'## Add items to the shipped and pending lists
FUNCTION ProductList (row, item, pending, locCode, soNum, shNum)		
	tmp = " "
	item = " "
	DO WHILE solSheet.Range("A" & row).value = soNum
		where = (solSheet.Range("T" & row).value = locCode)
		what = (solSheet.Range("R" & row).value = "SHIP")
		how = (solSheet.Range("AR" & row).value = shNum) 
		here = (solSheet.Range("C" & row).value = "In Progress") 
		IF where AND (how OR here) AND NOT what THEN
				item = item & "<tr><td>" & solSheet.Range("R" & row).value &_
					"</td><td>" & solSheet.Range("S" & row).value &_
					"</td><td>" & solSheet.Range("W" & row).value & "</td></tr>"
		END IF
		IF NOT where AND NOT how AND NOT what AND here	THEN
			tmp = tmp & "<tr><td>" & solSheet.Range("R" & row).value &_
					"</td><td>" & solSheet.Range("S" & row).value &_
					"</td><td>" & solSheet.Range("W" & row).value & "</td></tr>"
			pending = "Items pending shipment from this order: <br /><br />" &_ 
				"<table border='1' text-align='center' cellpadding='2'>" &_ 
				"<tr><td><b>Item #</b></td><td><b>Description</b></td><td><b>Qty</b></td></b></tr>" &_
				tmp &_
				"</table></span><br />" &_ 
				"As specified in our Resource Guide, drop-shipped items can take up to 4-weeks to ship. <br /><br /> " 
		END IF	
		row = row + 1
	LOOP
'## End Function ProductList	
END FUNCTION

'## Create and save the email as a draft
FUNCTION SendEmail(packNum, refNum, shipTo, trackNum, soNum, emailAdd, phone, item, pending)
	SET objEmail = objOutlook.CreateItem(olMailItem)
	'## Add special emails iff they exist
	IF NOT emailAdd = "" THEN
		objEmail.Recipients.add(emailAdd)
		noEmail = ""
	ELSE 
		noEmail = "<font color='RED'><b>*NO EMAIL*</b></font><br /><br />"
	END IF
	'## Add "Please Call" instructions iff they exist
	IF phone = "" THEN
		toCall = ""
	ELSE 
		toCall = "<font color='RED'><b>Please call when ships: " & phone & "</b></font><br /><br />"
	END IF
	'## Copy service@... to email
	objEmail.cc = "service@natureexplore.org"
	'## Standard subject line
	objEmail.Subject = "Nature Explore Shipping Information - Order: " & soNum
	'## Compose the HTML email draft
	objEmail.HTMLBody = "<font face='Calibri' size='3'>" & noEmail & "Hello,<br /><br />" &_ 
	"Items from your recent Nature Explore order have shipped. Below is your tracking information: <br /><br />" &_ 
	"<b>Shipping To:</b><br />" &_ 
	"&emsp;" & shipTo  & "<br />" &_ 
	"<b>Order #:</b>" & refNum(1) & "<br /><br />" &_ 
	"<table border='1' text-align='center'>" &_ 
	"<tr><td><b>Item #</b></td><td><b>Description</b></td><td><b>Qty</b></td></b></tr>" &_ 
	item &_ 
	"</table><br />" &_ 
	"<span style='text-align:center'><b>Shipment Method:</b> UPS <br />" &_ 
	"<b>Tracking number: </b>" &_ 
	"<a href='http://wwwapps.ups.com/WebTracking/processInputRequest?HTMLVersion=5.0&loc=en_US&Requester=UPSHome&tracknum=" &_
	trackNum & "&track.x=36&track.y=13'>" & trackNum & "</a><br />" &_ 
	"<b>Number of Packages:</b> " & packNum(1) & " <br /><br />" &_ 
	pending &_
	"Please open and inspect all items immediately upon receipt: "  &_
	"Carefully unpack and inspect the contents of all cartons and make sure all parts are there. " &_ 
	"If any parts are missing or damaged, contact Nature Explore immediately at 1-888-908-8733 or service@natureexplore.org. <br /><br />" &_ 
	"If you have any questions please feel free to call or email. <br /><br />" &_ 
	"Thanks!<br />" &_
	"<br /> <b>Natural Products Specialist Team</b><br />"  &_
	"Nature Explore<br />" &_
	"<small>www.natureexplore.org<br />" &_
	"P: 1-888-908-8733 x1</small></font><br />"
	'## Save the draft and put in specific folder for later review 
	objEmail.Save
	DIM UPSdraft
	SET objDraft = objNamespace.GetDefaultFolder(olFolderDrafts)
	SET UPSdraft = objDraft.Folders("UPS Notifications")
	objEmail.Move UPSdraft
END FUNCTION

'## Close Files and exit the program
objWorkbook.Close 
objWbs.Close 
objApp.Quit 
SET solSheet = NOTHING
SET	objWorkbook = NOTHING
SET objWbs = NOTHING
SET objApp = NOTHING
MsgBox "DONE!"
'#################################### END SCRIPT ###########################################
'-------------------------------------------------------------------------------------------
'################################### SCRIPT FOOTER #########################################
'## Nature Explore shipment notification script of packages shipped via UPS shipping service
'##    Created, Edited, Maintained By: Michael Kracke - Nature Explore 2013 - Present
'#################################### END FOOTER ###########################################