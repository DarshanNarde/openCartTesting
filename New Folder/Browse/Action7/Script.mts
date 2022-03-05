

URL="https://demo.opencart.com/"
SystemUtil.Run"chrome.exe",URL
DataTable.ImportSheet "C:\Users\HP\Desktop\UFT\Sprint2\Datafiles\SprintSP.xls", "Sheet1", "Global"

'Check visibilty for home page	
With Browser("name:=Your Store").Page("title:=Your Store")
	if .WebElement("XPath:=//body").GetRoProperty("visible") then
		wait(3)
		Reporter.ReportEvent micPass,"Confirm","Home Page is visble"
		msgbox "Home Page is visble"
	End If


	'Check visibilty for Company telephone number
	.WebElement("outertext:=123456789",Company_Number).GetRoProperty("visible")
	wait(3)
	Reporter.ReportEvent micPass,"Confirm","No. is visble"
	msgbox "Company telephone number is 123456789"

	'Check My Account Link is Working
	If .WebElement("XPath:=//span[contains(text(),'My Account')]").Exist Then
		Wait(3)
		.WebElement("XPath:=//span[contains(text(),'My Account')]").Click
		Reporter.ReportEvent micPass,"Confirm","My Account Link is working"	
	End If


'	'Resgister New Account
'	Services.StartTransaction "Register_Account"
'	.WebElement("XPath:=//span[contains(text(),'My Account')]").Click
'	.Link("text:=Register",Register_Link).Click
End With
'With Browser("name:=Register Account").Page("title:=Register Account")
'	.WebEdit(First,Name).Set DataTable("Name")
'	.WebEdit("XPath:=//input[@id='input-lastname']").Set DataTable("Last_Name")
'	.WebEdit("XPath:=//input[@id='input-email']").Set DataTable("Email_id")
'	
'
'	.WebEdit("XPath:=//input[@id='input-telephone']").Set DataTable("No.")
'	.WebEdit("XPath:=//input[@id='input-password']").Set DataTable("Password")
'	.WebEdit("XPath:= //input[@id='input-confirm']").Set DataTable("Confirm_Password")
'	.WebCheckBox("name:=agree",Agree_Check_box).Click
'	.WebElement(Click_For_Register).Click
'	.WebButton(Register_Button).Click	
'	
'	
'End With
'Browser("name:=Your Account Has Been Created!").Page("title:=Your Account Has Been Created!").WebButton("name:=Continue",Continue_Button).Click
'msgbox "Your Account is Successfully Created!!"
'Browser("name:=My Account").Page("title:=My Account").Link("innertext:=Logout",Logout_Button).Click
'msgbox "Log Out Succesfully.!"
'Services.EndTransaction "Register_Account"


'Acount Login
Services.StartTransaction "Account_Login"
if Browser("name:=Your Store").Page("title:=Your Store").WebElement(Login_Link).Exist (5)Then 
Browser("name:=Your Store").Page("title:=Your Store").WebElement(Login_Link).Click
	With Browser("name:=Account Login").Page("title:=Account Login")
		.WebEdit("html id:=input-email","XPath:=//input[@id='input-email']").Set DataTable("Email_id", dtGlobalSheet)
		.WebEdit("XPath:=//input[@id='input-password']").SetSecure DataTable("Password", dtGlobalSheet)
		.WebButton("name:=Login",Login_Click).Click
	End With
	'msgbox "login Succesfully.!"
End If
Services.EndTransaction "Account_Login"



'Check My Wish List link is working
With Browser("My Account").Page("My Account")
	.Link("Wish List (0)").Click Wait(2)
	Reporter.ReportEvent micPass,"Confirm","Wish List link is working"


	'Check Shopping Cart link is working

	.Link("Shopping Cart").Click
	.WebButton("Continue").Click
End With
Reporter.ReportEvent micPass,"Confirm","Shopping Cart link is working"



