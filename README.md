# ActiveDirectorySearch
This Active Directory Search tool was written in Visual Studio Community 2017 C# Windows Forms and exports the results from LDAP to csv format.

<h1 align="center">
  <img src="Images/main_form.png" alt="MyApp" />
</h1>

## Overview
This application uses LDAP (Lightweight Directory Access Protocol) to search Active Directory items from the treeview, the toolbar "Find" button or double clicking on the item in the listview

## Dependencies
|Software                        |Dependency                 |
|:-------------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|

| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
|LDAP |Lightweight Directory Access Protocol|
|CSV |Comma-Separated Values|


## Functionality
Listed below is the detailed functionality of this application and its components.  

####	Find Text (Dropdown)
* Lists all the values searched for in the current session

####	Find (Button)
* Queries Active Directory for the text in the "Find Text" textbox

#### Save (Button)
* Saves the current listbox view to a .csv file

####	User List (Button)
* This will change the layout of the listbox to show more information about each member of a group

####	Settings (Button)
* Opens the settings form. The domain name must be updated here.

####	About (Button)
* Opens the about form
