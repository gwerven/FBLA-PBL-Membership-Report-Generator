Name
———
FBLA-PBL Membership Report Generator v1.5


Description
—————
This program was created to maintain current member data and to output reports of member statistics for the national office of Future Business Leaders of America Phi-Beta-Lambda.


Features
————
Launching the application brings the user to a main screen consisting of five distinguished buttons. Momentarily hovering the cursor over a GUI component activates a tool tip text to appear, aiding the user in completing their intended task. This feature is much easier to access than a dedicated help menu.

1.) Edit Current Members - Opens the master file containing all current member data in Microsoft Excel for convenient editing.

2.) Add New Member - Transitions the user to another screen consisting of text boxes for inputting the information of a new member to be added to the master file. Ten text fields are provided for member number (Ex: 123456), first name (Ex: John), last name (Ex: Smith), grade (Ex: 10), school (Ex: Wood Crest High School), state (Ex: Arizona), email (Ex: jsmith@email.com), year joined (Ex: 2013), status (Ex: Active), and amount owed (Ex: 15.00). After inputting the desired information and clicking the “Add Member” button, a dialog will prompt the user to either add another member or not. If “Yes” is clicked, the text fields will clear. If “No” is clicked, the user will be returned to the home screen. Restart the program before viewing, printing, or exporting to ensure the report has updated information.

3.) Member Dues Report - Generates an .xls report upon clicking followed by a dialog informing the user that the report was successfully generated and then a screen consisting of four new buttons (View, Print, Export, and Back). The report will include the member number, first name, last name, year joined, grade, and amount owed for all members in the database that owe dues. The footer of the report will contain the total number of active and nonactive members as well as the number of members that owe dues and total amount owed.

	1.) View - Opens the generated .xls report in Microsoft Excel for convenient viewing.

	2.) Print - Prints the generated .xls report onto paper using the user’s default printing device.

	3.) Export - Exports the generated .xls report to a user specified destination (destination can be 	chosen via a dialog). Allows the user to give the file a customized name.

	4.) Back - Transitions the user to the previous main screen.

4.) Senior Email Report - Generates an .xls report upon clicking followed by a dialog informing the user that the report was successfully generated and then a screen consisting of four new buttons (View, Print, Export, and Back). The report will include the first name, last name, and email for all senior members in the database.

	1.) View - Opens the generated .xls report in Microsoft Excel for convenient viewing.

	2.) Print - Prints the generated .xls report onto paper using the user’s default printing device.

	3.) Export - Exports the generated .xls report to a user specified destination. (Destination can be chosen via a dialog). Allows the user to give the file a customized name.

	4.) Back - Transitions the user to the previous main screen.

	5.) FBLA-PBL Web Site - Opens a new tab in the user’s default web browser and navigates to the FBLA-PBL web page (http://www.fbla-pbl.org/).


Change Log
——————
v1.0 – Original release
v1.5 – Implementation of proper OOP principles in the GUI. Custom classes extending swing components were created for Buttons, Labels, Text Fields, Panels, and Frames in order to allow for the recycling of code.


Known Bugs
——————
None as of v1.5


System Requirements
——————————
Operating System: Windows XP or greater
Disk Space: 327 KB
Processor: Minimum Pentium 2 266 MHz
Memory: 128 MB


Install/Uninstall, Configuration, and Operating Instructions
————————————————————————
FBLA-PBL Membership Report Generator is currently distributed solely by disk copy. To run the application, simply insert the disk containing the application and its files into your computer’s optical disk drive and allow the application to run. No installation is required and therefore uninstallation is obsolete. Similarly, no configuration is required. FBLA-PBL Membership Report Generator runs as is out of the box.


Files List
————
addMemberButtonIcon.png
addNewMemberButtonIcon.png
backButtonIcon.png
cityscape.png
exportButtonIcon.png
FBLALogo.png
FBLAWebPageButtonIcon.png
FBLA-PBLLogo.png
masterFile.xls
memberDuesReportButtonIcon.png
printButtonIcon.png
readme.txt
seniorEmailReportButtonIcon.png
viewButtonIcon.png


Acknowledgements
—————————
Application created by Garrett Erven (Email: gwerven@gmail.com)
Apache POI, Microsoft Excel, and Google Chrome were used in the creation of this project.

*This application is not currently endorsed by Future Business Leaders of America Phi-Beta-Lambda.