'################################## SCRIPT HEADER #########################################
'##  This script will take SO #'s along with Excel spreasheets take selected information
'##    from them and place them into a template email and save that email as a draft.
'## 
'##
'## This will be broken into to parts: Scan excel spreadsheet for information                                    					                        
'##                                    Save template (w/ info from UPS) in drafts folder
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
'##   - This script should live in C:\Scripts\NES\Shipnotify.vbs
'##   - Outlook should also contain the folder structure "Drafts\UPS Notifications"
'##        This is to separate the drafts create by this script from any other drafts 
'############################### END OF INSTRUCTIONS ######################################
'------------------------------------------------------------------------------------------
'################################# START OF SCRIPT ########################################
'## Connect to Outlook and Exel
DIM objApp, objWbs, objWorkbook, solSheet
SET objApp = CreateObject("Excel.Application")
SET objWbs = objApp.WorkBooks
objApp.Visible = FALSE
SET objWorkbook = objWbs.Open("\\milkyway\Sourcebook\_Order Processing\Daily Dimensions Sourcebook Orders - Morton - NatExplr Company.xlsx")
SET solSheet = objWorkbook.Sheets("Sales Order Line")
CONST olFolderDrafts = 16
SET objOutlook = CreateObject("Outlook.Application")
SET objNamespace = objOutlook.GetNamespace("MAPI")
SET objEmail = objOutlook.CreateItem(olMailItem)
DIM email, emailAdd, phone, create, shipTo
shipTo = ""
phone = ""
create = TRUE

'## Take input from Excel
DO WHILE create = TRUE
ON ERROR RESUME NEXT
	strInput = InputBox("Enter order #, location, shipping method and tracking #: (ex. 'SO90010934, NE, UPS 1ZX4900GD472'):")
	'## Snarf SO #
	DIM Reg3, so, soNum, catch3
	SET Reg3 = NEW RegExp
	WITH Reg3 
		.Pattern = "SO[0-9]{8}"
		.IgnoreCase = TRUE
		.Global = TRUE
	END WITH
	IF Reg3.test(strInput) THEN
		SET so = Reg3.Execute(strInput)
		FOR EACH catch3 IN so
			soNum = UCASE(catch3)
		NEXT
	END IF
	'## Find order in Excel from <soNum>
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
	'## Get the location code, track # and ship method
	DIM loc, trackNum, shipMethod
	refNum = Split(UCase(strInput), ",")
	trackNum = refNum(3)
	shipMethod = refNum(2)
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
	END SELECT
	'## Call functions
	emailAdd = solSheet.Range("M" & i).value
	getEmail soNum, emailAdd, phone
	ShipAddress shipTo
	ProductList row, item, pending, locCode, soNum, shNum
	Tracking trackNum, shipMethod
	SendEmail packNum, refNum, shipTo, trackNum, shipMethod, emailAdd, phone, item, pending
	'## Clear variables
	packNum = " "
	refNum = " "
	shipTo = " "
	trackNum = " "
	shipMethod = " "
	emailAdd = " "
	phone = " "
	locCode = " "
	item = " "
	pending = " "
	'## Ask to create another email or to quit program
	click = MsgBox ("Create another email?", 4)
	IF NOT click = vbYes THEN
		create = FALSE
	END IF
'## Create email iff vbYes
LOOP		
	
'!! - DOES NOT WORK 
'## Add hyperlinks to main shipping methods 
FUNCTION Tracking(trackNum, shipMethod)
	'msgbox "inTrack"
	'msgbox shipMethod & ": " & UCASE(shipMethod)
	SELECT CASE shipMethod
		CASE shipMethod = "FEDEX" 
		'msgbox trackNum
			trackNum = "<a href='https://www.fedex.com/fedextrack/index.html?tracknumbers=" & trackNum & "&cntry_code=us'> " &_
			trackNum & "</a>"
		'kmsgbox trackNum
		CASE UCASE(shipMethod) = "SAIA" 
			trackNum = "<a href='www.saia.com'> " & UCASE(trackNum) & "</a>"
		CASE UCASE(shipMethod) = "SPEEDEE"
			trackNum = "<a href='http://www.speedeedelivery.com/'> " & UCASE(trackNum) & "</a>"
		CASE UCASE(shipMethod) = "UPS" 
			trackNum = "<a href='http://wwwapps.ups.com/WebTracking/processInputRequest?HTMLVersion=5.0&loc=en_US&Requester=UPSHome&tracknum=" &_
			trackNum & "&track.x=36&track.y=13'>" & trackNum & "</a><br />" 
		CASE UCASE(shipMethod) = "YRC" 
			trackNum = "<a href='www.YRC.com'> " & UCASE(trackNum) & "</a>"
		END SELECT
'## End Functions Tracking		
END FUNCTION 

'## Get the email address, special emails, & phone #
FUNCTION getEmail(soNum, emailAdd, phone)
SET objWorkbook = objWbs.Open("\\milkyway\Sourcebook\_Order Processing\Special Shipment Tracking Notifications.xlsx")
SET emailSheet = objWorkbook.Sheets("Sheet1")
emailAdd = solSheet.Range("M" & i - 1).value
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
'## End Function get Email		
END FUNCTION
	
'## Add the shipping address
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
'## End Function ShipAddress
END FUNCTION

'## Add items to the shipped and pending lists
FUNCTION ProductList (row, item, pending, locCode, soNum, shNum)		
	tmp = " "
	item = " "
	'## Shipped List
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
		'## Pending Lists
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

'!! Create and save the email as a draft
FUNCTION SendEmail(packNum, refNum, shipTo, trackNum, shipMethod, emailAdd, phone, item, pending)
	SET objEmail = objOutlook.CreateItem(olMailItem)
	'## Add special email iff they exist
	IF NOT emailAdd = "" THEN
		objEmail.Recipients.add(emailAdd)
		noEmail = ""
	ELSE 
		noEmail = "<font color='RED'><b>*NO EMAIL*</b></font><br /><br />"
	END IF
	'## Add Call instructions iff they exist
	'IF NOT phone = "" THEN
	'	toCall = "<font color='RED'><b>Please call when ships: " & phone & "</b></font><br /><br />"
	'ELSE 
		toCall = " "
'	END IF
	'## Copy service@... 
	objEmail.cc = "service@natureexplore.org"
	'## Standard subject line
	objEmail.Subject = "Nature Explore Shipping Information - Order: " & refNum(0)
	'## Compose HTML Email Draft
	objEmail.HTMLBody = "<font face='Calibri' size='3'>" & noEmail & toCall & "Hello,<br /><br />" &_ 
	"Items from your recent Nature Explore order have shipped. Below is your tracking information: <br /><br />" &_ 
	"<b>Shipping To:</b><br />" &_ 
	"&emsp;" & shipTo  & "<br />" &_ 
	"<b>Order #: </b>" & refNum(0) & "<br /><br />" &_ 
	"<table border='1' text-align='center'>" &_ 
	"<tr><td><b>Item #</b></td><td><b>Description</b></td><td><b>Qty</b></td></b></tr>" &_ 
	item &_ 
	"</table><br />" &_ 
	"<span style='text-align:center'><b>Shipment Method:</b> " & shipMethod & "<br />" &_ 
	"<b>Tracking number:</b> " & trackNum & "<br />" &_
	"<b>Number of Packages:</b> <br /><br />" &_ 
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
	'## Save the draft and move to folder for later review
	objEmail.Save
	DIM UPSdraft
	SET objDraft = objNamespace.GetDefaultFolder(olFolderDrafts)
	SET UPSdraft = objDraft.Folders("UPS Notifications")
	objEmail.Move UPSdraft
'## End Function sendEmail
END FUNCTION

'## Close files and quit program
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