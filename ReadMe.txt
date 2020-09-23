The purpose of this demonstration application is to illustrate advanced techniques in using forms within the application or in an process (not remote) DLL to encapsulate functionality without effecting usability.

The key benefit of using an in-process DLL is in application distribution. If a change is made to only one part of the application contained in a DLL, then only the DLL needs to be distributed. Another benefit could include cross application useability.

An example of where this might be useful is if we have an inventory application and we want to view the suppliers on the same screen with the stock item or view all stock items on the suppliers screen. A common viewport is used to display associated information. If we are looking at the supplier screen, other information can be displayed in the common view port and could include outstanding orders, reports, historical data, etc. Each screen of data would be stored in their own forms.

The same information is displayed in both of the inventory application examples but requires no double coding as each form encapsulates the code for the interface and functionality/business-logic. For this to work, an interface class would be created and implemented into each form creating a common method of communicating. 

A Form can be both a parent and child. The viewport on each parent form would be enabled or disabled depending if it was the parent or child form. 

A Special Note: The code to glue the forms together was borrowed from a Wizard Interface example application developed by Klaus H. Probst [kprobst@vbbox.com] http://www.vbbox.com/ - definitely worth a visit if you haven't already been!

Happy Coding...
