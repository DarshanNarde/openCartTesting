'Login_openCart
'importing sheet

	DataTable.ImportSheet "D:\openCart\Sprint2\Datafiles\SprintAJ.xls", "Sheet1", "Global"
	
	'starting transaction
	Services.StartTransaction "shopping_cart"
	'clicking on shopping cart

	'checking for the image of product is showing or not in shopping cart
With AutomationPage1
      '  .WebTable(xpath_webtable_sc_6c).Image("class:=img-thumbnail").WaitProperty"visible","True"
	c=.WebTable(xpath_webtable_sc_6c,"name:=Image").Image("class:=img-thumbnail").Exist'GetROProperty("visible")
	If  c = "True" Then
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
	e=.WebTable(xpath_webtable_sc_6c).WebEdit("class:=form-control").GetROProperty("value")


	m=.WebTable(xpath_webtable_sc_6c).WebElement(xpath_webelement_unitprice).GetROProperty("outertext")
	DataTable.Value("unit_price")=m
	n=.WebTable(xpath_webtable_sc_6c).WebElement(xpath_webelement_totalprice).GetROProperty("outertext")
	 DataTable.Value("subtotal")=n

	'custom checkpoint to validate whether product of quanity & unit price is same as total price
	 
	   
        ak = DataTable.Value("Validation")
'Generating Report to validate subtotal&quanity*unit price
	If ak="True"Then
		Reporter.ReportEvent micPass,"subtotal_price check","the product of quantity and unit price matches with subtotal price"
		else
			Reporter.ReportEvent micFail,"subtotal_price check","the product of quantity and unit price does not matches with subtotal price"
	End If
	DataTable.Value("subtotal")=n

	'Estimate shipping taxes
	.Link("xpath:=//DIV[2]/DIV[1]/H4[1]/A[1]").WebElement("xpath:=//DIV[2]/DIV[1]/H4[1]/A[1]/I[1]").Click
	.WebList("name:=country_id","select type:=ComboBox Select").select("India")
	.Sync
	 wait(1)
	.WebList("value:=--- Please Select ---","name:=zone_id").select("Goa")
	.Sync
	.WebEdit("name:=postcode").set("416104")
	.Sync
	.WebButton("name:=Get Quotes").Click
	.WebRadioGroup("name:=shipping_method").Click
	.WebButton("name:=Apply Shipping").Click
        .WebElement("innertext:=Success: Your shipping estimate has been applied! × ").checkproperty "visible","True"
          

	'Fetching shipping charges
	
	.WebTable(xpath_webtable_sc_2c).WebElement(xpath_webelement_shippingcharge).WaitProperty"visible","True"
	
	c=.WebTable(xpath_webtable_sc_2c).WebElement(xpath_webelement_shippingcharge).GetROProperty("outertext")
	DataTable.Value("shipping_charge")=c
	
	.WebElement("xpath:=//DIV[@id='content']/DIV[2]/DIV[1]").WaitProperty"Visible","True"
	wait(2)
	h=.WebTable("xpath:=//DIV[@id='content']/DIV[2]/DIV[1]/TABLE[1]").GetROProperty("text")
	
        h1=split(h,":")
	DataTable.Value("total_price")=h1(3)
	
	'custom checkpoint to validate whether addition of shipping charge &
	'subtotal is equal to total_price
	bk = DataTable.Value("validation")
	'Generating Report to validate whether addition of shipping charge &subtotal is equal to total_price
	If bk="True"Then
		Reporter.ReportEvent micPass,"total_price check","addition of shipping charge &subtotal price is equal to total price"
		else
			Reporter.ReportEvent micFail,"total_price check","addition of shipping charge & subtotal price is not equal to total price"
	End If
	
	'redirected to homepage
	.Link("name:=Continue Shopping").Click
End With


Services.EndTransaction "shopping_cart"


DataTable.ExportSheet "D:\openCart\Sprint2\Datafiles\SprintAJ.xls", "Global", "Sheet1"
