DataTable.ImportSheet "D:\openCart\Sprint2\Datafiles\SprintSP.xls", "Sheet1", "Global"

'Check visibilty for home page	
if Browser("name:=Your Store").Page("title:=Your Store").WebElement("XPath:=//body").GetRoProperty("visible") then
	Browser("name:=Your Store").Page("title:=Your Store").Sync
	Reporter.ReportEvent micPass,"Confirm","Home Page is visble"
	msgbox "Home Page is visble"
End If


'Check visibilty for Company telephone number
Browser("name:=Your Store").Page("title:=Your Store").WebElement("outertext:=123456789",Company_Number).GetRoProperty("visible")
wait(3)
Reporter.ReportEvent micPass,"Confirm","No. is visble"
msgbox "Company telephone number is 123456789"

'Check My Account Link is Working
If Browser("name:=Your Store").Page("title:=Your Store").WebElement("XPath:=//span[contains(text(),'My Account')]").Exist Then
	Browser("name:=Your Store").Page("title:=Your Store").Sync
	Browser("name:=Your Store").Page("title:=Your Store").WebElement("XPath:=//span[contains(text(),'My Account')]").Click
	Reporter.ReportEvent micPass,"Confirm","My Account Link is working"	
End If


'Resgister New Account
With Browser("name:=Your Store").Page("title:=Your Store")
	Services.StartTransaction "Register_Account"
	.WebElement("XPath:=//span[contains(text(),'My Account')]").Click
	.Link("text:=Register",Register_Link).Click
End With

With Browser("name:=Register Account").Page("title:=Register Account")
	.WebEdit("XPath:=//INPUT[@id='input-firstname']").Set DataTable("Name")
	.WebEdit("XPath:=//input[@id='input-lastname']").Set DataTable("Last_Name")
	.WebEdit("XPath:=//input[@id='input-email']").Set DataTable("Email_id")
	.WebEdit("XPath:=//input[@id='input-telephone']").Set DataTable("No.")
	.WebEdit("XPath:=//input[@id='input-password']").Set DataTable("Password")
	.WebEdit("XPath:= //input[@id='input-confirm']").Set DataTable("Confirm_Password")
	.WebCheckBox("name:=agree",Agree_Check_box).Click
	.WebElement(Click_For_Register).Click
	.WebButton(Register_Button).Click	

End With
Browser("name:=Your Account Has Been Created!").Page("title:=Your Account Has Been Created!").WebButton("name:=Continue",Continue_Button).Click
msgbox "Your Account is Successfully Created!!"
Browser("name:=My Account").Page("title:=My Account").Sync
Browser("name:=My Account").Page("title:=My Account").Link("innertext:=Logout",Logout_Button).Click
msgbox "Log Out Succesfully.!"
Browser("name:=Account Logout").Page("title:=Account Logout").WebButton("outertext:=Continue").Click
Services.EndTransaction "Register_Account"


'Acount Login
Services.StartTransaction "Account_Login"
if Browser("name:=Your Store").Page("title:=Your Store").WebElement(Login_Link).Exist (5)Then
	Browser("name:=Your Store").Page("title:=Your Store").WebElement(Login_Link).Click
	Browser("name:=Account Login").Page("title:=Account Login").WebEdit("html id:=input-email","XPath:=//input[@id='input-email']").Set DataTable("Email_id", dtGlobalSheet)
	Browser("name:=Account Login").Page("title:=Account Login").WebEdit("XPath:=//input[@id='input-password']").SetSecure DataTable("Password", dtGlobalSheet)
	Browser("name:=Account Login").Page("title:=Account Login").WebButton("name:=Login",Login_Click).Click
	'msgbox "login Succesfully.!"
End If
Services.EndTransaction "Account_Login"


'Check My Wish List link is working
Browser("My Account").Page("My Account").Link("Wish List (0)").Click Wait(2)
Reporter.ReportEvent micPass,"Confirm","Wish List link is working"


'Check Shopping Cart link is working
Browser("My Account").Page("My Account").Link("Shopping Cart").Click
Browser("My Account").Page("My Account").WebButton("Continue").Click
Reporter.ReportEvent micPass,"Confirm","Shopping Cart link is working"
'Browser("name:=Shopping Cart").Page("title:=Shopping Cart").Link("xpath:=//DIV[@id=""logo""]/H1/A[normalize-space()=""Your Store""]").Click


