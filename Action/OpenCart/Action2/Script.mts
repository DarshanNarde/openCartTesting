Dim GoogleBrowser, GooglePage
Set GoogleBrowser= Browser ("name:=Your Store")
Set GooglePage= Browser ("title:=Your Store")

'Checking store logo
With Browser("name:=Your Store")
	a = .Page("title:=Your Store").Link("xpath:=//DIV[@id='logo']/H1/A[normalize-space()='Your Store'] ").GetROProperty("visible")
	'Generating Report for logo visibility
	If a = "True" Then
		Reporter.ReportEvent micPass,"Confirmation","Logo is visible"
		msgbox "Logo is visible"

	End If

	'Checking that "Your Store" Button Redirect to home page
	.Page("title:= Your Store").WebElement("xpath:= //DIV/DIV/DIV/H4[normalize-space()='MacBook']").Click
	With .Page("title:=Your Store")
		.Link("xpath:= //DIV[@id='logo']/H1/A[normalize-space()='Your Store']").Click
		
		'Validating Currency Option
		msgbox "Default currency is Dollar"
		'Setting Currency to Euro
		.WebButton("class:=btn btn-link dropdown-toggle").Click
		.WebButton("xpath:=//LI/BUTTON[normalize-space()=""€ Euro""]").Click
		.Sync
		c2= .WebElement("class:=price","xpath:= //DIV[1]/DIV[1]/DIV[2]/P[2]").GetROProperty("innertext")
	End With
End With
'Validating currency is changed or not
If instr (c2,"€") Then
		Reporter.ReportEvent micPass, "Result","Currency is changed"
		msgbox "Currency is changed to Euro "
		Browser("name:=Your Store").Page("title:=Your Store").WebButton("class:=btn btn-link dropdown-toggle").Click
		Browser("name:=Your Store").Page("title:=Your Store").WebButton("xpath:= //LI/BUTTON[normalize-space()='$ US Dollar']").Click
		Browser("name:=Your Store").Page("title:=Your Store").Sync
	else
		Reporter.ReportEvent micFail, "Result","Currency remains same"
End If
'Validating Shopping cart is initially empty or not
Browser("opentitle:=Your Store").Page("title:=Your Store").WebButton("class:=btn btn-inverse btn-block btn-lg dropdown-toggle").Click
Wait(2)
a=Browser("name:=Your Store").Page("title:=Your Store").WebElement("xpath:=//DIV[@id='cart']/UL[1]").Exist
If a="True" Then
	Reporter.ReportEvent micPass,"shopping cart initially","After login shopping cart is empty"
else
	Reporter.ReportEvent micFail,"shopping cart initially","After login shopping cart is not empty"
End If
Browser("name:=Your Store").Page("title:=Your Store").Link("xpath:=//DIV[@id='logo']/H1/A[normalize-space()='Your Store']").Click
