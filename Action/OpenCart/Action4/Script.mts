
'Importing the External excel sheet for input value
DataTable.ImportSheet "D:\openCart\Sprint2\Datafiles\SprintAJ.xls", "Sheet1", "Global"


wait(3)
'Searchinhg the Product from Datatable

Browser("name:=Your Store").Page("title:=Your Store").WebEdit("class:=form-control input-lg").Set DataTable("Search", dtGlobalSheet)
'Mac
str1=DataTable("Search", dtGlobalSheet)
Browser("name:=Your Store").Page("title:=Your Store").WebButton("class:=btn btn-default btn-lg").Click

Services.StartTransaction "Searching_Product"
With Browser("name:=Search - .*").Page("title:=Search - .*")
	'Checkpoints for validating the Search process
	.WaitProperty "visible",True,3000
	str2=.WebElement("Class name:=WebElement","xpath:=//DIV[@id='content']/H1[1]").GetROProperty("innertext")
End With
'msgbox a = Search - Mac
'i =Mac

'Generating the reports for Search 
If Instr(str2,str1) Then
       ' Generating Pass Report for Search Complete
	Reporter.ReportEvent micPass," Result"," Search is Completed"
	else
	'Generating Fail Report for Search Fail
	Reporter.ReportEvent micFail," Result"," Search is Failed"
End If

'Searching Product based on category and its sub-category
With Browser("name:=Search - .*").Page("title:=Search - .*")
	
	.WebList(Category,"xpath:=//DIV[2]/SELECT[1]").Select"Desktops"
	.Sync

	'Radio marking to search in sub-category
	.WebCheckBox("xpath:=//DIV[3]/LABEL[1]/INPUT[1]").Set
	.Sync

	'Clicking Search Button
	.WebButton(SearchButton).Click
	.Sync

	'Checking result list from different grid view
	.WebButton(Gridview,"xpath:=//button[@id='list-view']","html id:=list-view").Click
	.Sync

	'Sorting the Result list 
	.WebList(Sortlist,"xpath:=//SELECT[@id='input-sort']").Select "Price (High > Low)"
	.Sync

	'Controlling the count of result list
	.WebList(Sortlist,"xpath:=//SELECT[@id='input-limit']").Select "25"
	.Sync

	'Adding product to wish list from result list
	.WebButton("xpath:=//DIV[1]/DIV[1]/DIV[2]/DIV[2]/BUTTON[2]").Click
End With
'Showing The msgbox of WishList
With Browser("name:=Search - .*").Page("title:=Search - .*")
	.Sync
       If .WebButton("xpath:=//DIV[1]/DIV[1]/DIV[2]/DIV[2]/BUTTON[2]").Exist="True" Then
		sc1=.Link("xpath:=//DIV[@id='product-search']/DIV[1]/A[1]").GetROProperty("text")
              sc2=.Link("xpath:=//DIV[@id='product-search']/DIV[1]/A[2]").GetROProperty("text")
              msg1 = "Success: You have added "&sc1&" to your "&sc2
              msgbox msg1
	End If
End With

With Browser("name:=Search - .*").Page("title:=Search - .*")
	
	'Adding product to comparison from result list
	.WebButton("xpath:=//DIV[1]/DIV[1]/DIV[2]/DIV[2]/BUTTON[3]").Click
End With
'Showing The msgbox of Comparison
With Browser("name:=Search - .*").Page("title:=Search - .*")
	.Sync
       If .WebButton("xpath:=//DIV[1]/DIV[1]/DIV[2]/DIV[2]/BUTTON[2]").Exist="True" Then
		sc1=.Link("xpath:=//DIV[@id='product-search']/DIV[1]/A[1]").GetROProperty("text")
              sc2=.Link("xpath:=//DIV[@id='product-search']/DIV[1]/A[2]").GetROProperty("text")
              msg2 = "Success: You have added "&sc1&" to your "&sc2
              msgbox msg2
       
	End If
End With
Services.EndTransaction "Searching_Product"

'Checkpoints for validating the Search product in result
With Browser("name:=Search - .*").Page("title:=Search - .*")
	b =.Image(ProductImage,"xpath:=//DIV[1]/DIV[1]/DIV[1]/A[1]/IMG[1]").Exist
	'msgbox b
	If b = "True" Then
		i2=.Link("xpath:=//DIV[1]/DIV[1]/DIV[2]/DIV[1]/H4[1]/A[1]").GetROProperty("text")
		.Image("Class name:=Image","xpath:=//DIV[1]/DIV[1]/DIV[1]/A[1]/IMG[1]").Click
		Else
			'Generating fail report if there is no product.
			Reporter.ReportEvent micFail," Result","There is no product that matches the search criteria. "
			'if there is no product than it will click on home button to go back
			.WebElement("html tag:=LI""xpath:=//DIV[@id='product-search']/UL[1]/LI[1]").Click
	End If
End With



