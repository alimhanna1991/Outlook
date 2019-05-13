On Error Resume Next
Const END_OF_STORY = 6
Set objSysInfo = CreateObject("ADSystemInfo")
' ///////////////////////////////////// 1.1.2 ////////////////////////////////////
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objuser.FirstName + " " + objUser.LastName
strTitle = objUser.Title
strCompany = objUser.Company
strOffice = objUser.physicalDeliveryOfficeName
strPhone = objUser.telephonenumber
strFax = objUser.faxnumber
strMobile = objUser.mobile 'get  user Mobile #
strPOBox = objUser.PostOfficeBox
strCity = objUser.l
strCountry = objUser.co
strMail = objuser.mail ' get  user Email form Active Dir
	strHyperlink = "www.yourwebsite.com" ' add website

Set objWord = CreateObject("Word.Application")
objWord.FontNames.Item = "Arial"
Set objDoc = objWord.Documents.Add()

Set objSelection = objWord.Selection
Set objRange = objDoc.Range()
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
	
' new /////////////////////
Const NUMBER_OF_ROWS =13 
Const NUMBER_OF_COLUMNS =3
objDoc.Tables.Add objRange, NUMBER_OF_ROWS, NUMBER_OF_COLUMNS

Set objTable = objDoc.Tables(1)

objSelection.TypeParagraph()
objSelection.TypeText(Chr(13))
objSelection.Style = "No Spacing"
objSelection.PageSetup.LeftMargin()=0
objTable.Cell(1,1).Range.Text= strName
objTable.Cell(2,1).Range.Text = strTitle
objTable.Cell(3,1).Range.Text= strCompany
objTable.Cell(4,1).Range.Text = "_________________________________________________"  
objTable.Cell(4,2).Range.Text ="" 
objTable.Cell(6,2).Range.InlineShapes.AddPicture("path/to/logo")
objTable.Cell(6,2).SetHeight= 70
objTable.Cell(6,2).SetWidth= 70
objTable.Cell(5,1).Range.Text = "P.O.Box: "  & strPOBox
objTable.Cell(6,1).Range.Text = strCity + ", " + strCountry
objTable.Cell(7,1).Range.Text ="Tel      :" & strPhone 
objTable.Cell(8,1).Range.Text="Fax     :" & strFax
objTable.Cell(9,1).Range.Text="Mob    :" & strMobile 
objTable.Cell(11,1).Range.Hyperlinks.Add objTable.Cell(11,1).Range,strHyperlink & " ", , , strHyperlink, "_blank"
'objTable.Cell(10,1).Range.Hyperlinks.Add objTable.Cell(10,1).Range, "Mail to: " & strMail, , , strMail
objLink = objSelection.Hyperlinks.Add(objTable.Cell(10,1).Range,"E-mail" & strMail, , , strMail)
Set Range =objDoc.Tables(1).Range
Range.Style = "No Spacing"
Set Range = objDoc.Tables(1).Cell(1, 1).Range 
Range.Font.Bold = True ' set  cell to bold

Set Range = objDoc.Tables(1).Cell(2, 1).Range
Range.Font.Color  = RGB(128,128,128) 
Range.Font.Name = "Arial"
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(3, 1).Range
Range.Font.Color  = RGB(128,128,128)
Range.Font.Name = "Arial"
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(4, 1).Range
Range.Font.Color  = RGB(28,134,94) 
Set Range = objDoc.Tables(1).Cell(5, 1).Range
Range.Font.Color  = RGB(128,128,128)
Range.Font.Name = "Arial" 
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(5, 2).Range
Range.Borders.DistanceFromLeft = 0
Set Range = objDoc.Tables(1).Cell(6, 1).Range
Range.Font.Color  = RGB(128,128,128) 
Range.Font.Name = "Arial"
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(7, 1).Range
Range.Font.Color  = RGB(128,128,128) 
Range.Font.Name = "Arial"
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(8, 1).Range
Range.Font.Color  = RGB(128,128,128)
Range.Font.Name = "Arial"
Range.Font.Size = "9" 
Set Range = objDoc.Tables(1).Cell(9, 1).Range
Range.Font.Color  = RGB(128,128,128) 
Range.Font.Name = "Arial"
Range.Font.Size = "9"
Set Range = objDoc.Tables(1).Cell(10, 1).Range
Range.Font.Name = "Arial"
Range.Font.Size = "8"
Range.Hyperlinks 
Set Range = objDoc.Tables(1).Cell(11, 1).Range
Range.Font.Name = "Arial" 
Range.Font.Size = "8"
Set Range = objDoc.Tables(1).Cell(13, 1).Range
Range.Font.Color  = RGB(128,128,128)
Range.Font.Name = "Arial" 
Range.Font.Size = "7"
Set objCell1 = objTable.Cell(6,2)
Set objCell2 = objTable.Cell(5,2)
Set objCell3 = objTable.Cell(5,2)
Set objCell4 = objTable.Cell(7,2)
Set objCell5 = objTable.Cell(8,2)
Set objCell6 = objTable.Cell(9,2)
Set objCell7 = objTable.Cell(10,2)
Set objCell9 = objTable.Cell(11,2)
Set objCell11 = objTable.Cell(13,2)
Set objCell12 = objTable.Cell(4,1)
Set objCell13 = objTable.Cell(4,2)
objCell2.Merge(objCell3) 'Merge cells
objCell1.Merge(objCell2) 'Merge cells
objCell4.Merge(objCell3)
objCell5.Merge(objCell3)
objCell6.Merge(objCell3)
objCell7.Merge(objCell3)
objCell9.Merge(objCell3)
objCell10.Merge(objCell3)
objCell11.Merge(objCell3)
objCell12.Merge(objCell13)
objCell14.Merge(objCell3)

objTable.Columns(1).Width =10 ' set width of column 1 as required
objTable.Rows(13).Width = 500 ' set width of column 2 as required
objSelection.EndKey END_OF_STORY
objSelection.Font.Color = RGB(128,128,128)
objSelection.Font.Size = "7"
objSelection.Font.Name = "Arial"
objSelection.TypeText "This email and any attachments are confidential and may also be privileged.  If you are not the addressee, do not disclose, copy, circulate or in any other way use or rely on the information contained in this email or any attachments.  If received in error, notify the sender immediately and delete this email and any attachments from your system. Alwataniya Microfinance is not responsible for the political, religious, racial or partisan opinion in any correspondence conducted by its domain users. Therefore, any such opinion expressed, whether explicitly or implicitly, in any said correspondence is not to be interpreted as that of Alwataniya Microfinance. Emails cannot be guaranteed to be secure or error free as the message and any attachments could be intercepted, corrupted, lost, delayed, incomplete or amended. Although Alwataniya Microfinance has taken steps to ensure that e-mails and attachments are free from any virus, we advise that, in keeping with best business practice, the recipient must ensure they are actually virus free. Alwataniya Microfinance does not accept liability for damage caused by this email or any attachments. Alwataniya Microfinance may monitor all incoming and outgoing e-mails in line with its business practice."

Set objSelection = objDoc.Range()
'\\\\\\\\\\\\\\\////////////////////

'\\\\\\\\\\\\\\\//////////////////
objSignatureEntries.Add "Your Signature Name as Ali Mhanna Sig", objSelection
objSignatureObject.NewMessageSignature = "Your Signature Name as Ali Mhanna Sig"
objSignatureObject.ReplyMessageSignature = "Your Signature Name as Ali Mhanna Sig"
objDoc.Saved = True
objWord.Quit
