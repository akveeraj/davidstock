
Installing Smart Referrer for Access Database:		                                  	            20/02/2001 	


A.Contents of the zip file In the zip file you downloaded you will find the following files:

	1. SmartReferrer.mdb (Access Database)
	2. smart_Referrer.htm (.htm file to be included in pages requiring monitoring) 
	3. SmartReferrerAdmin.asp (Administration panel for pages to be monitored and referrer reports) 
	4. smart_referrer_fill.gif 
	5. smart_referrer_header.gif 


B. Installation steps 

	1. Important: Place the SmartReferrer.mdb and the smart_referrer.htm files either under the site root(preferable) or if you know the exact path to them you can place them anywhere convenient for you. Files 3,4 and 5 should be under the same folder e.g Admin folder.
	If your site is being hosted by another site or you haven't placed it under the site root, you will need to change the path to the access database in the smart_referrer.htm and SmartReferrerAdmin.asp pages. This can be done very simply by opening these pages and going to the top of the code. Here you will find the following lines:

	 '------------------If your site is hosted by another site then change your path in the DBQ value below i.e. in place of Server.MapPath("/SmartReferrer.mdb") type in Server.MapPath("/Your_site_path/SmartReferrer.mdb")-----------------

	strDBRef = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/SmartReferrer.mdb") & ";DefaultDir=" & Server.MapPath(".") & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5" 

	'------------------End of Database connection string -----------------

	Just give your database's ( SmartReferrer.mdb file's) path in the DBQ value i.e. in place of Server.MapPath("/SmartReferrer.mdb") type in Server.MapPath("/Your_site_path/SmartReferrer.mdb"). 

	2. Very Important NOTE: All Monitored pages must be .asp files.
	
	3. Include the smart_Referrer.htm file as a server side include(SSI) in all pages needing monitoring (recommended at the bottom). 
	(Or)
	just add this code :
	 <!--#include virtual="/your_folder_path/smart_Referrer.htm"  --> 
	Note : Make sure the path is right. 

	4. The Administration panel is for activating, editing, deleting and deactivating pages that need monitoring and accessing the 
reports. Add the complete URL of the pages to be monitored in the admin panel i.e., in SmartReferrerAdmin.asp. for eg. We need to add http://www.smartwebby.com/default.asp for our default page. 

	5. Also view the referrers to your page in the SmartReferrerAdmin.asp page. There are very comprehensive complete reports on the hits to your page. 

		a. Brief Report: Reports the total hits till date, This week's hits, Today's hits, Last recorded date's hits and Last recorded week's hits for the monitored page. 

		b. Today's Report: Gives information on all the referrers that have monitored your page today. 

		c. General Report: This report gives you the total hits till date, This week's hits, Today's hits, Last recorded date's hits and Last recorded week's hits for each of the referrers. 

		d. No Hits Report: This report gives all the referrers that have not accessed you page today. 

		e. Display of monitored and archived pages: Here you can see all the pages that are monitored and also the pages that are archived or deactivated. 



Good Luck and Best Wishes from the Smart Webby Team! If you have any problems or suggestions please contact us at info@smartwebby.com. 


   SmartWebby.com - 2001 All Rights Reserved. 

   This product has been created by SmartWebby.com for free distribution. 
   We will not be held responsible for any unwanted effects due to the usage of this product or any derivative. 
   No warrantees for usability are given or implied. 