'################################## SCRIPT HEADER #########################################
'##      This script will be used to create emails for Quality Control purposes.
'##
'##  '       = Line of code not approved for live use
'##  '##     = Comment
'##  '####   = Debugger
'##  '!!!!! = Update
'##                              
'##                          [Date Edited: 2-12-2014]  
'##                         [Edited By: Michael Kracke]  
'##                               [Version 2.0]
'################################### END HEADER ###########################################
'------------------------------------------------------------------------------------------
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!                                       UPDATES FROM LAST EDIT
'!  Changes in Code
'!  - Fixed the URL errors in products
'!
'!  Changes in Formatting
'!  - Added comments and spacing for easy reading   
'!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'------------------------------------------------------------------------------------------
'#################################### Item List ###########################################
'##     | Item # |         Description             |
'##      -------- ---------------------------------
'##     |  3902  | Tree Cookie Flooring, set of 16 |
'##     |  3912  | Marimba, Short                  |
'##     |  3917  | Redcedar Log Steps              |
'##     |  3918  | Natural Balance Beam, Crooked   |
'##     |  3920  | Natural Balance Beam, Straight  |
'##     |  4273  | Akambira, Short                 |
'##     |  5591  | Akambira, Large                 |
'##     |  5555  | Pole Trellis, Individual        |
'##     |  5556  | Pole Trellis, set of 2          |
'##     |  5590  | Marimba, Tall                   |
'##     |  6889  | Winged Chimes                   |
'##     |  6900  | Open Outdoor Storage Unit       |
'##     |  6901  | Closed Outdoor Storage Unit (S) |
'##     |  6902  | Closed Outdoor Storage Unit (L) |
'##     |  6941  | Resonated Metallophone, IG      |
'##     |  6942  | Resonated Metallophone, FS-S    |  
'##     |  6943  | Resonated Metallophone, FS-T    |  
'##     |  6919  | Rustic Arbor                    |
'##      -------- ---------------------------------
'##########################################################################################
'------------------------------------------------------------------------------------------
'################################# START OF SCRIPT ########################################
'## Setup and configure
Option Explicit
DIM objApp, objWbs, objWorkbook, solSheet, shSheet
SET objApp = CreateObject("Excel.Application")
SET objWbs = objApp.WorkBooks
objApp.Visible = FALSE
SET objWorkbook = objWbs.Open("\\milkyway\Sourcebook\_Order Processing\Daily Dimensions Sourcebook Orders - Morton - NatExplr Company")
SET solSheet = objWorkbook.Sheets("Sales Order Line")
SET shSheet =  objWorkbook.Sheets("Sales Order Header")
DIM i, j, thisWeek, itemNum, itemNum2, prodList, tipList, URLbegin, soNum, oldSo
DIM addTip1, addTip2, addTip3, addTip4, tip1, tip2, tip3, tip4, row
DIM p3902, p3912, p3917, p3918, p3920, p4273, p5555, p5556, p5590, p5591, p6889
DIM p6919, p6941, p6942, p6943, BREAK
DIM itemCount, toEmail, plural1, plural2, plural3, emailAdd, orderDate, shipTo, poNum
i = 2
j = i 
itemCount = 0
thisWeek = DATE - 8
toEmail = FALSE 
addTip1 = FALSE 
addTip2 = FALSE 
addTip3 = FALSE 
addTip4 = FALSE 
BREAK = TRUE
plural1 = "an item that requires" 
plural2 = "its" 
plural3 = "product" 
shipTo = ""
tipList = ""
tip1 = 	"<br /> &emsp; - Have you identified your installation area and factored in Use Zone space?"
tip2 =	"<br /> &emsp; - Permanently installed wooden items are untreated. If choosing to weather treat, ensure the sealant complies with safety codes."
tip3 = 	"<br /> &emsp; - Have you reviewed the mallet safety guidelines?"
tip4 = 	"<br /> &emsp; - Are the additional installation materials and tools gathered?"
URLbegin = "<br /><a href='http://www.natureexplore.org/NaturalProducts/merchDetail.cfm?ID="	

'## Find items from past week			
DO WHILE NOT CDATE(solSheet.Range("B" & i).value) = thisWeek AND BREAK
	IF solSheet.Range("B" & i).value = "" THEN
		BREAK = FALSE
	END IF
	i = i + 1
LOOP
BREAK = TRUE
DO WHILE NOT CDATE(solSheet.Range("B" & i).value) = " " AND BREAK
	IF solSheet.Range("B" & i).value = "" THEN
		BREAK = FALSE
	END IF
	itemNum = solSheet.Range("R" & i).value
	soNum = solSheet.Range("A" & i).value
	oldSo = solSheet.Range("A" & i - 1).value
	emailAdd = solSheet.Range("M" & i - 1).value
	orderDate = CDATE(solSheet.Range("B" & i).value)	
	IF NOT soNum = oldSo THEN 
 		IF toEmail THEN
			DO WHILE NOT shSheet.Range("A" & j).value = oldSo		
				j = j + 1
			LOOP
    		row = shSheet.range("A" & j).row
 		    poNum = shSheet.Range("D" & j).value
 			IF NOT solSheet.Range("AH" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AH" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AI" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AI" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AJ" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AJ" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AK" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AK" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AL" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AL" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AM" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AM" & i - 1).value & "<br /> &emsp;"
 			END IF
 			IF NOT solSheet.Range("AN" & i - 1).value = "" THEN
 				shipTo = shipTo & solSheet.Range("AN" & i - 1).value  & "<br />"
 			END IF
 			Email tipList, prodList, oldSo, plural1, plural2, plural3, emailAdd, orderDate, shipTo, poNum
 			toEmail = FALSE
 		END IF
		
		'## Clear all variables
 		j = 1 
 		itemCount = 0
 		shipTo = "" 
 		orderDate = "" 
		poNum = ""
		tipList = "" 
 		prodList = " " 
 		plural1 = "an item that requires" 
 		plural2 = "its" 
 		plural3 = "product"
		prodList = ""
		itemcount = 1
		addTip1 = FALSE 
		addTip2 = FALSE 
		addTip3 = FALSE 
		addTip4 = FALSE 
 		p3902 = FALSE 
 		p3912 = FALSE 
 		p3917 = FALSE 
 		p3918 = FALSE 
 		p3920 = FALSE 
 		p4273 = FALSE 
 		p5555 = FALSE 
 		p5556 = FALSE 
 		p5590 = FALSE 
 		p5591 = FALSE 
 		p6889 = FALSE 
 		p6919 = FALSE 
 		p6941 = FALSE 
 		p6942 = FALSE 
 		p6943 = FALSE
 		toEmail = FALSE
 		BREAK = TRUE
	END IF
	
	'## Add products
 	SELECT CASE itemNum
 		CASE "3902"
 			prodList = prodList & URLbegin & "15'>3902 - Tree Cookie Flooring, set of 16</a>"
 			p3902 = TRUE
 			itemCount = itemCount + 1
 		CASE "3912"
 			prodList = prodList & URLbegin & "14'>3912 - Marimba, Short</a>"
 			p3912 = TRUE
 			itemCount = itemCount + 1
 		CASE "3917"
 			prodList = prodList & URLbegin & "5'>3917 - Redcedar Log Steps</a>"
 			p3917 = TRUE
 			itemCount = itemCount + 1
 		CASE "3918"
 			prodList = prodList & URLbegin & "6'>3918 - Natural Balance Beam, Crooked</a>"
 			p3918 = TRUE
 			itemCount = itemCount + 1
 		CASE "3920"
 			prodList = prodList & URLbegin & "6'>3920 - Natural Balance Beam, Straight</a>"
 			p3920 = TRUE
 			itemCount = itemCount + 1
		CASE "4273"
			prodList = prodList & URLbegin & "36'>4273 - Akambira, Short</a>"
			p4273 = TRUE
			itemCount = itemCount + 1
		CASE "5555"
			prodList = prodList & URLbegin & "90'>5555 - Pole Trellis, Individual</a>"
			p5555 = TRUE
			itemCount = itemCount + 1
		CASE "5556"
			prodList = prodList & URLbegin & "90'>5556 - Pole Trellis, set of 2</a>"
			p5556 = TRUE
			itemCount = itemCount + 1
		CASE "5590"
			prodList = prodList & URLbegin & "14'> 5590 - Marimba, Tall</a>"
			p5590 = TRUE
			itemCount = itemCount + 1
		CASE "5591"
			prodList = prodList & URLbegin & "36'>5591 - Akambira, Tall</a>"
			p5591 = TRUE
			itemCount = itemCount + 1
		CASE "6889"
			prodList = prodList & URLbegin & "121'>6889 - Winged Chimes</a>"
			p6889 = TRUE
			itemCount = itemCount + 1
		CASE "6919"
			prodList = prodList & URLbegin & "113'>6919 - Rustic Arbor</a>"
			p6919 = TRUE
			itemCount = itemCount + 1
		CASE "6941"
			prodList = prodList & URLbegin & "128'>6941 - Resonated Metallophone, In-Ground</a>"
			p6941 = TRUE
			itemCount = itemCount + 1
		CASE "6942"
			prodList = prodList & URLbegin & "128'>6942 - Resonated Metallophone, Free Standing, Short</a>"
			p6942 = TRUE
			itemCount = itemCount + 1
		CASE "6943"
			prodList = prodList & URLbegin & "128'>6943 - Resonated Metallophone, Free Standing, Tall</a>"
			p6943 = TRUE
			itemCount = itemCount + 1
	END SELECT
	
	'## Add questions
	IF (p3917 OR p3918 OR p3920) AND NOT addTip1 THEN
		tipList = tipList & tip1
		addTip1 = TRUE	
	END IF
	IF (p3902 OR p3912 OR p3917 OR p3918 OR p3920 OR p4273 OR p5555 OR p5556 OR p5590 OR p5591 OR p6919) AND NOT addTip2 THEN
		tipList = tipList & tip2
		addTip2 = TRUE
	END IF
	IF (p3912 OR p4273 OR p5590 OR p5591) AND NOT addTip3 THEN
		tipList = tipList & tip3
		addTip3 = TRUE
	END IF
	IF (p3902 OR p3912 OR p3917 OR p3918 OR p3920 OR p4273 OR p5555 OR p5556 OR p5590 OR  p5591 OR p6889 OR p6919 OR p6941 OR p6942 OR p6943) AND NOT addTip4 THEN
		tipList = tipList & tip4
		addTip4 = TRUE
	END IF  
	IF addTip1 OR addTip2 OR addTip3 OR addTip4 THEN
		toEmail = TRUE
	END IF
	IF itemCount > 1 THEN
		plural1 = "items that require" 
		plural2 = "their" 
		plural3 = "products"
	END IF
	i = i + 1
LOOP

'## Send Email
FUNCTION Email(tipList, prodList, soNum, plural1, plural2, plural3, emailAdd, orderDate, shipTo, poNum)
	DIM objOutlook, objNamespace, objEmail, olMailItem, poIn, poSeg, noEmail,ccEmail
	poIn = FALSE & poSeg = ""
	SET objOutlook = CreateObject("Outlook.Application")
	SET objEmail = objOutlook.CreateItem(olMailItem)
	IF NOT  poNum = "" THEN
		poIn = TRUE
		poSeg = ", PO #: " & poNum
	END IF
	IF NOT emailAdd = "" THEN
		objEmail.Recipients.add(emailAdd)
		ccEmail = "service@natureexplore.org; jeffl@natureexplore.org;"  
		noEmail = ""
	ELSE
		emailAdd = "service@natureexplore.org; jeffl@natureexplore.org;"  
		ccEmail = ""
		noEmail = "<font color='RED'><b>*NO EMAIL*</b></font><br /><br />"
	END IF
	CONST olFolderDrafts = 16
	CONST olTxt = 0
	SET objOutlook = CreateObject("Outlook.Application")
	SET objNamespace = objOutlook.GetNamespace("MAPI")
	SET objEmail = objOutlook.CreateItem(olMailItem)
	objEmail.Recipients.add(emailAdd)
	objEmail.cc = ccEmail       
	objEmail.Subject = "Nature Explore Order: Assembly/Installation Instructions"
	objEmail.HTMLBody = "<font face='Calibri' size='3'>Hello, <br /><br />" &_ 
	"Your Nature Explore order, " & soNum & poSeg & ", placed on " & orderdate  &_
	", contains " & plural1 & " additional assembly/installation. In preparation for " & plural2 & " arrival, " &_
	"please review the  instructions and key considerations listed below.<br /><br />" &_
	"<b>Shipping To:</b> <br />" &_
	"&emsp;"  & shipTo &_ 
	"<b>" & prodList & "</b><br />" & tipList & "<br /><br />" &_ 
	"Outdoor classroom items require routine maintenance including visual inspection, " &_
	"bolt tightening, sanding and in most cases weather sealing. Wooden items will " &_
	"check and crack over time, and develop a beautiful silvery-gray patina. In " &_
	"addition to proper assembly/installation, we suggest creating a maintenance schedule " &_
	"to ensure items are in top condition; this will help extend the useful life of your " &_
	plural3 & ".<br /><br />" &_
	"<a href='http://www.natureexplore.org/NaturalProducts/warranty.cfm' " &_ 
	"style='text-decoration:none;'>Click here</a> for our quality commitment information. <br />" &_
	"<br />" &_ 
	"Thank you for taking the time to review the assembly/installation instructions. " &_
	"We will email you shipment tracking information as soon as it is available. " &_
	"(As specified in our Resource Guide, drop-shipped items can take up to 4-weeks to ship.)" &_
	"<br />" &_
	"<br />" &_
	"<br />" &_ 
	"Thanks again,<br />" &_
	"<br /> <b>Natural Products Specialist Team</b><br />"  &_
	"Nature Explore<br />" &_
	"<small>www.natureexplore.org<br />" &_
	"P: 1-888-908-8733 x1</small></font><br />"
	DIM objFolder, QAdraft
	objEmail.Save
	SET objFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
	SET QAdraft = objFolder.Folders("QC Email Drafts")
	objEmail.Move QAdraft
END FUNCTION

'## Close files
MsgBox "Done!"
objWorkbook.Close 
objWbs.Close 
objApp.Quit 
SET shSheet = NOTHING
SET solSheet = NOTHING
SET	objWorkbook = NOTHING
SET objWbs = NOTHING
SET objApp = NOTHING
'#################################### END SCRIPT ###########################################
'################################### SCRIPT FOOTER #########################################
'## Nature Explore QA Email drafting script
'##         Created, Edited, Maintained By: Michael Kracke - Nature Explore 2013
'#################################### END FOOTER ###########################################
