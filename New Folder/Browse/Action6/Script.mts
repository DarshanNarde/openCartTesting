 'importing sheet
DataTable.ImportSheet "C:\Users\HP\Desktop\UFT\Sprint2\Datafiles\SprintAJ.xls", "Sheet1", "Global"
'xpath_webtable_sc_6c="xpath of webtable of shopping cart consisting of 6 column"
'xpath_webtable_sc_2c="xpath of webtable of shopping cart consisting of 2 column"
Dim gb,gp
Set gb=Browser("name:=Shopping Cart")
Set gp=Page("title:=Shopping Cart")

''Opening the application with url by chrome
'URL="HTTP://demo.opencart.com"
'SystemUtil.Run "Chrome.exe",URL

'starting transaction
Services.StartTransaction "shopping_cart"
'clicking on shopping cart
'Browser("name:=Your Store").Page("title:=Your Store").Link("xpath:=//LI/A[normalize-space()='Shopping Cart']").Click

'checking for the image of product is showing or not in shopping cart
With AutomationPage1
         .WebTable(xpath_webtable_sc_6c).Image("class:=img-thumbnail").WaitProperty"visible","True"
c=.WebTable(xpath_webtable_sc_6c).Image("class:=img-thumbnail").Exist'GetROProperty("visible")
If c="True" Then
'Creating Report based on image shown in shopping cart
Reporter.ReportEvent micPass,"checking image on cart","The image of the product is shown in the shopping cart"
else
Reporter.ReportEvent micFail,"checking image on cart","The image of the product is not shown in the shopping cart"
End If

'Fetching product name at runtime
d=.WebTable(xpath_webtable_sc_6c).WebElement(xpath_webelement_name).GetROProperty("outertext")

DataTable.Value("product_name")=d

'Fetching product model
z=.WebTable(xpath_webtable_sc_6c,"name:=Image").WebElement(xpath_webelement_model).GetROProperty("outertext")

DataTable.Value("product_model")=z
'Entering the quantity
.WebTable(xpath_webtable_sc_6c).WebEdit("class:=form-control").Set DataTable("Quantity",dtGlobalSheet)
.WebTable(xpath_webtable_sc_6c).WebButton("class:=btn btn-primary").Click

'checkpoint for modified quantity
u=.WebElement("outertext:=Success: You have modified your shopping cart! ×").CheckProperty("visible",1)
If u="True" Then
msgbox "Success: You have modified your shopping cart!"
End If
'fetching Quantity 
e=.WebTable(xpath_webtable_sc_6c).WebEdit("class:=form-control").GetROProperty("value")
'Fetching Unit Price and Total Price
m=.WebTable(xpath_webtable_sc_6c).WebElement(xpath_webelement_unitprice).GetROProperty("outertext")
n=.WebTable(xpath_webtable_sc_6c).WebElement(xpath_webelement_totalprice).GetROProperty("outertext")
'Adding Unit price in Datatable
DataTable.Value("unit_price")=m

'custom checkpoint to validate whether product of quanity & unit price is same as total price
     q=split(m,"$")
      k=q(1)
      g=split(n,"$")
      l=g(1)
      w=FormatNumber(e*k,2)
'Generating Report to validate total price&quanity*unit price
If w=l Then
Reporter.ReportEvent micPass,"total_price check","the product of quantity and unit price matches with total price"
else
Reporter.ReportEvent micFail,"total_price check","the product of quantity and unit price does not matches with total price"
End If
DataTable.Value("total_price")=n

'Estimate shipping taxes
.Link("xpath:=//DIV[2]/DIV[1]/H4[1]/A[1]").WebElement("xpath:=//DIV[2]/DIV[1]/H4[1]/A[1]/I[1]").Click
.WebList("name:=country_id","select type:=ComboBox Select").select("India")
wait(1)
.WebList("name:=zone_id","select type:=ComboBox Select").select("Bihar")
.WebEdit("name:=postcode").set("423402")
.WebButton("name:=Get Quotes").Click
.WebRadioGroup("name:=shipping_method").Click
.WebButton("name:=Apply Shipping").Click
ai=.WebElement("class:=alert alert-success alert-dismissible").GetROProperty("innertext")
msgbox ai


'Fetching shipping charges

.WebTable(xpath_webtable_sc_2c).WebElement(xpath_webelement_shippingcharge).WaitProperty"visible","True"
c=.WebTable(xpath_webtable_sc_2c).WebElement(xpath_webelement_shippingcharge).GetROProperty("outertext")
DataTable.Value("shipping_charge")=c


.WebElement("xpath:=//DIV[@id='content']/DIV[2]/DIV[1]").WaitProperty"Visible","True"
	wait(2)
	h=.WebTable("xpath:=//DIV[@id='content']/DIV[2]/DIV[1]/TABLE[1]").GetROProperty("text")
	
        h1=split(h,":")
	DataTable.Value("final_price")=h1(3)

       .WebButton("name:=Checkout").Click
End With

Services.EndTransaction "shopping_cart"

'export excel sheet
DataTable.ExportSheet "C:\Users\HP\Desktop\UFT\Sprint2\Datafiles\Sprint2.xls", "Global", "Sheet1"
