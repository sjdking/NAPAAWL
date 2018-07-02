#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance Force
#INCLUDE ADOSQL.AHK

;App to extract data from Oasis for submission to NAPAAWL
;	1. Enter date range
;	2. Vaildate data, ask for corrections where necessary
;	3. Generate spreadsheet with relevant data for submission to NAPAAWL
;	NOTE: Requires NAPAAWLCSVIndex.txt and AppointmentReferences.xlsx for indexing

Global ADOSQL_LastError, ADOSQL_LastQuery

;Setup SQL connection strings for Oasis and TherapyClinic databases
OasisConnect := "Driver={SQL Server};Server=WSQE004SQL\OASIS9_2;Database=OASIS;Uid=OasisUserWSQE004SQL;Pwd="

FormatTime, extractedDate, A_Now, ddMMyyyy
FormatTime, extractedTime, A_Now, HHmmss

outputFile = C:\temp\AutoHotKeyScripts\Oasis\NAPAAWL\NAPAAWL Extracts\NAPAAWL_%extractedDate%_%extractedTime%.csv

numPatients := 0

Gui, Add, Text, x10 y10, Select start date:
Gui, Add, DateTime, x110 y10 vstartDate ChooseNone
Gui, Add, Text, x10 y40, Select end date:
Gui, Add, DateTime, x110 y40 vendDate ChooseNone
Gui, Add, Button, x10 y70 gbtnOK, OK
Gui, Add, Button, x50 y70 gbtnCancel, Cancel
Gui, Show,, Extract data for NAPAAWL DC

GoSub, AppointmentIndex

Return

btnOK:
{
	Gui, Submit, NoHide
	Gui, Destroy
	Gui, Add, Progress, w500 Range0-1000 vMyProgress
	Gui, Add, Text, vMyText wp Center
	Gui, Show
	
	queryStartDate := SubStr(startDate,1,8)
	queryEndDate := SubStr(endDate,1,8)
	outputFile = C:\temp\AutoHotKeyScripts\Oasis\NAPAAWL\NAPAAWL Extracts\NAPAAWL_%queryStartDate%_%A_Now%.csv
	csvIndexFile := "C:\temp\AutoHotKeyScripts\Oasis\NAPAAWL\NAPAAWLCSVIndex.txt"
	Loop, Read, %csvIndexFile%
	{
		Loop, Parse, A_LoopReadLine, `n
			FileAppend,%A_LoopField%`n,%outputFile%
	}

	GoSub, extractPAAPPLNS
	Return
}

;Subroutine to find all appointments in the specified date range that are associated with a patient
extractPAAPPLNS:
{
	queryPAAPPLNS := % "SELECT * FROM PAAPPLNS WHERE SKey BETWEEN '" . queryStartDate . "%' AND '" . queryEndDate . "%'"
	returnPAAPPLNS := ADOSQL(OasisConnect, queryPAAPPLNS)
	
	Loop
	{
		i := A_Index + 1
		GoSub, ClearFields
		SKey := % returnPAAPPLNS[i,1]
		patNumber := % returnPAAPPLNS[i,2]
		patNumber = %patNumber%
		apptColour := % returnPAAPPLNS[i,12]
		apptPicture := % returnPAAPPLNS[i,11]
		If (patNumber > 0)
			numPatients ++
		GuiControl,, MyProgress, %numPatients%
		GuiControl,, MyText, Building patient list %numPatients% Patient %patNumber% SKey %SKey%
		If (patNumber > 0)
		{
			queryPBPATMAS := % "SELECT * FROM PBPATMAS WHERE PatNumber = " . patNumber
			patReturn := ADOSQL(OasisConnect, queryPBPATMAS)
			
			patLastUpdated := % patReturn[2,100]
			
			patURN := % patReturn[2,3]
			patURN = %patURN%
			patLastName := % patReturn[2,4]
			patFirstName := % patReturn[2,6]
			patFirstName = %patFirstName%
			patLastName = %patLastName%
			patSecondName = % patReturn[2,152]
			patSecondName = %patSecondName%
			patDOB := % patReturn[2,11]
			patGender := % patReturn[2,12]
			patStreet1 := % patReturn[2,13]
			patStreet2 := % patReturn[2,14]
		;need to remove any commas from the address fields	
			patStreet1 := StrReplace(patStreet1,"`,")
			patStreet2 := StrReplace(patStreet2,"`,")
			
			patSuburb := % patReturn[2,15]
			patSuburb = %patSuburb%
			patSuburb := StrReplace(patSuburb,"`,")
			
			If (RegExMatch(patSuburb, " WA$| VIC$| NSW$| SA$| QLD$| TAS$| NT$| ACT$"))
			{
				FoundPos := RegExMatch(patSuburb, "\s(\w+$)", patState)
			;use %patState1%
			}
			Else
				patState1 := "WA"
			patSuburb := RegExReplace(patSuburb, " WA$| VIC$| NSW$| SA$| QLD$| TAS$| NT$| ACT$", "")
	
			If (patStreet2 != "")
			{
				mailAddress1 := % patReturn[2,135]
				mailAddress2 := % patReturn[2,136]
				mailSuburb := % patReturn[2,137]
			}
			Else
			{
				mailAddress1 := % patReturn[2,135]
				mailAddress2 := ""
				mailSuburb := % patReturn[2,136]
			}
			
			patPostcode := % patReturn[2,16]
			patHomePhone := % patReturn[2,20]
			patMobilePhone := % patReturn[2,23]
			patDateOfDeath := % patReturn[2,51]
			patMedicareNumber := % patReturn[2,29]
			patBirthCountry := % patReturn[2,43]
			patIndigStatus := % patReturn[2,44]
			patMaritalStatus := % patReturn[2,47]
			patResidentialStatus := % patReturn[2,49]
			patDVANumber := % patReturn[2,153]
			patDVAColour := % patReturn[2,154]
			
			If (patDVAColour = "BC")
				patPensionNumber := % patReturn[2,26]
			
			patRefAccNo := % patReturn[2,45]
			patRefDate := % patReturn[2,70]
			patRefPriority := % patReturn[2,46]
			
			mailPostcode := SubStr(mailSuburb,-3)
			mailSuburb := RegExReplace(mailSuburb, " WA \d+| VIC \d+| NSW \d+| SA \d+| QLD \d+| TAS \d+| NT \d+| ACT \d+", "")
			
			apptDate := SubStr(SKey,1,8)
			
			timeHH := SubStr(SKey,17,2)
			timeMMslot := SubStr(SKey,19,2)
			timeMM := 5*(timeMMslot-1)
			If (StrLen(timeMM) = 1)
				timeMM := % timeMM . "0"
			apptTime := % timeHH . ":" . timeMM
			
			queryPBSTICKY := % "SELECT * FROM PBSTICKY WHERE PatNumber = " . patNumber
			stickyReturn := ADOSQL(OasisConnect, queryPBSTICKY)
			stickyDetails := % stickyReturn[2,4]
			stickyDetails = %stickyDetails%
			foundDischarged := InStr(stickyDetails,"discharg",false)
			If (foundDischarged = 1)
			{
				dischargeDate := % stickyReturn[2,3]
				MsgBox, 4,,Note for %patURN% %patFirstName% %patLastName% reads`n"%stickyDetails%".`n`nHave they been discharged? Y/N
				IfMsgBox No
					dischargeDate = ""
			}
		;GoSub, extractApptDetails
		apptColumn := SubStr(SKey,9,8)
		apptType := apptArr[apptColour,apptColumn]
		
		FileAppend,105`,BLANK`,%patURN%`,BLANK`,%patLastName%`,%patFirstName%`,%patSecondName%`,%patDOB%`,%patDateOfDeath%`,BLANK`,%patGender%`,%patMedicareNumber%`,%patBirthCountry%`,BLANK`,%patIndigStatus%`,%patMaritalStatus%`,%patStreet1%`,%patStreet2%`,%patSuburb%`,%patPostcode%`,%patState1%`,%patResidentialStatus%`,%patHomePhone%`,%patMobilePhone%`,%patDVANumber%`,%patDVAColour%`,%patRefAccNo%`,REFERRAL STATUS CODE`,REFERRAL CATEGORY CODE`,%patRefDate%`,%patRefPriority%`,REFERRAL REASON CODE`,REFERRAL SOURCE CODE`,%dischargeDate%`,%SKey%`,%apptPicture%`,APPOINTMENT PRIORITY CODE`,APPOINTMENT REASON CODE`,%apptDate%`,%apptTime%`,%apptType%`,APPOINTMENT PAYMENT CLASSIFICATION CODE`,APPOINTMENT CLIENT TYPE CODE`,APPOINTMENT SESSION CODE`,BLANK`,BLANK`,CLINIC CATEGORY CODE`,CLINIC IDENTIFIER`,CLINIC TITLE`,DOCTOR LED CLINIC INDICATOR`,CLINIC NMDS TIER 1 CODE`,CLINIC NHCDC TIER 2 CODE`,APPOINTMENT OUTCOME CODE`,APPOINTMENT ATTENDANCE CODE`,APPOINTMENT PATIENT ARRIVAL TIME`,APPOINTMENT PATIENT SEEN TIME`,BLANK`,OTH`,APPOINTMENT DELIVERY MODE CODE`,APPOINTMENT DELIVERY SETTING CODE`,BLANK`,BLANK`,BLANK`,%extractedDate%`,%extractedTime%`,BLANK`,BLANK`,BLANK`,BLANK`,BLANK`,APPOINTMENT CANCELLATION CODE`,BLANK`,BLANK`,BLANK`,BLANK`,%mailAddress1%`,%mailAddress2%`,%mailSuburb%`,%mailPostcode%`,ON HOLD`,%patLastUpdated%`n,%outputFile%
		
		;Return
		}
		
	}
	Until SKey = ""
	Gui, Destroy
	
	MsgBox, %numPatients% valid appointments extracted
	Return
}

ClearFields:
{
	SKey := "" 				;used to get appointment data, time and appointment column
	patNumber := "" 		;used to extract patient information from PBPATMAS table
	apptColour := ""		;used to cross reference to appointment column to determine appointment type
	apptPicture := ""		;used to determine appointment status (Oasis Icon)
	patURN := ""			;patient URN
	patLastName := ""		;patient last name
	patFirstName := ""		;patient first name
	patSecondName := ""		;patient second name
	patDOB := ""			;patient date of birth
	patGender := ""			;patient gender
	patStreet1 := ""		;residential address field 1, will be processed to remove any commas
	patStreet2 := ""		;residential address field 2, will be processed to remove any commas
	patSuburb := ""			;residential address suburb
	patPostcode := ""		;residential address postcode
	patState1 := ""			;residential state
	patHomePhone := ""		;home phone number
	patMobilePhone := ""	;mobile phone number
	patDateOfDeath := ""	;patient date of death (blank if NA)
	patMedicareNumber := ""	;medicare number including line number
	patBirthCountry := ""	;patient country of birth
	patIndigStatus := ""	;patient indigenous status
	patMaritalStatus := "TO DO"	;patient marital status
	patResidentialStatus := "TO DO"	;patient residential status
	patDVANumber := ""		;patient DVA number
	patDVAColour := ""		;patient DVA card colour
	patRefAccNo := ""		;referral account number
	patRefDate := ""		;referral date
	patRefPriority := ""	;referral priority
	apptDate := ""			;appointment date DDMMYYYY
	apptTime := ""			;appointment time HH:MM
	dischargeDate := ""		;date of discharge
	apptPicture := ""		;Oasis icon for appointment
	apptType := ""			;appointment type
	mailAddress1 := ""		;mailing address field 1
	mailAddress2 := ""		;mailing address field 2
	mailSuburb := ""		;mailing address field 3
	extractedDate := ""		;system extracted date
	extractedTime := ""		;system extracted time
	patLastUpdated := ""	;last patient update
	Return
}

;Subroutine to generate array of appointment types
AppointmentIndex:
{
	inputFile := "C:\temp\AutoHotKeyScripts\Oasis\NAPAAWL\AppointmentReferences.xlsx"

	oWB := ComObjGet(inputFile)

	apptArr := {}

	Loop, 77
	{
		column := A_Index + 2
		apptColumnIndex := oWB.Sheets(1).Cells(1,column).Value
		Loop, 34
		{
			row := A_Index
			apptColourIndex := Floor(oWB.Sheets(1).Cells(row,1).Value)
			loopVal := oWB.Sheets(1).Cells(row, column).Value
			If (loopVal !="")
			{
				apptArr[apptColourIndex,apptColumnIndex] := loopVal
			}
		}
	}
	Return
}

btnCancel:
ExitApp

GuiClose:
ExitApp

^Esc::ExitApp