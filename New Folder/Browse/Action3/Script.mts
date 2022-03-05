
'taking current system date
curDate=Date
'import excel sheet 
DataTable.ImportSheet "C:\Users\HP\Desktop\UFT\Sprint2\Datafiles\SprintDN.xls", "Sheet1", "Global"


'transaction for add to cart module
Services.StartTransaction "addToCartModule"

'Enable reporting of all events "default setting"
Reporter.Filter=rfEnableAll

'checking the product is out of stock 
With CartBrowser
	.Sync
	wait(3)
	'cart status
	 ex=.WebButton(cartStatus,"Class Name:=WebButton").GetROProperty("outertext")
        msgbox "Existing cart details - "&ex
	.Sync
	x=.WebElement(Availability).Exist
	
	        'if-then-else condition == if the product is out of stock then it will nevigate to main page
		If x=True Then
			        URL="https://demo.opencart.com/index.php"
			        SystemUtil.Run "chrome",URL
			
		'else it will go for add to cart option
		Else
	
				'fetching product name and price 
				y=.WebElement(name).GetROProperty("innerhtml")
				z=.WebElement(price).GetROProperty("innerhtml")
	
				'adding the product name and price into excel sheet
				DataTable.Value("ProductName")=y
				DataTable.Value("Price")=z
	
				'Sync point 
				.Sync
				
				'clicking on adding cart button   
				.WebButton(buttonCart).Click
				
				
		End If
		
End With

'end of transaction
Services.EndTransaction "addToCartModule"

'transaction for success msg of product added into cart
Services.StartTransaction "SuccessMsg"

'checking product is added into cart or not with confirmation msg
With CartBrowser
     
	w=.WebElement(successMsg).GetROProperty("outertext")

	       'if success msg is showing it will generate report for pass result
		If instr(w,"Success") Then
			
			'report for pass condition 
			Reporter.ReportEvent micPass," Result"," product added into shopping cart"
	
			'msgbox for updated cart details
			'final cart status
			ms=.WebButton(finalCart).WebElement("html id:=cart-total").GetROProperty("outertext")
			msgbox "your product "&y&" had been added to your cart with a price value of "&z&" on "&curDate
			msgbox "Updated cart Details- "&ms
	
			'it will go for checkout 
			.WebButton(clickP).Click
			'wait before click on checkout link
			.WaitProperty "text", "Checkout",2000
			.Link(checkout).Click
			
			.Sync
	
		else
			'report for fails condition
			Reporter.ReportEvent micFail," Result","  product is not added into shopping cart"
		
		End If
		
End With

'end of transaction
Services.EndTransaction "SuccessMsg"

'export excel sheet
'DataTable.ExportSheet "C:\Users\Darshan\Desktop\UFT\test output\Sprint_2.xls", "Global", "Sheet1"

