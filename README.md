<img align="left" src="Images/header.png">

This Active Directory Search tool was written in Microsoft Visual Studio Community 2017 C#/VB.NET Windows Forms and exports the results from LDAP to csv format. It uses LDAP (Lightweight Directory Access Protocol) to search Active Directory items from the treeview, the toolbar "Find" button or double clicking on the item in the listview.

[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.me/AnthonyDuguid/1.00)
[![Join the chat at https://gitter.im/ScriptHelp/Lobby](https://badges.gitter.im/ActiveDirectorySearch/Lobby.svg)](https://gitter.im/ActiveDirectorySearch/Lobby?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © Anthony Duguid")
[![Latest Release](https://img.shields.io/github/release/Visual-Studio-projects/Active-Directory-Search.svg?label=latest%20release)](https://github.com/Visual-Studio-projects/Active-Directory-Search/releases)
[![Github commits (since latest release)](https://img.shields.io/github/commits-since/Visual-Studio-projects/Active-Directory-Search/latest.svg)](https://github.com/Visual-Studio-projects/Active-Directory-Search)
<a href="https://github.com/Visual-Studio-projects/Active-Directory-Search/issues" target="_blank">
 <img alt="Issues" src="https://img.shields.io/github/issues/Visual-Studio-projects/Active-Directory-Search.svg" />
</a>

<h1 align="center">
  <img src="Images/main_form2.png" alt="MyApp" />
</h1>

<br>

## Table of Contents
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a>

<br>

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                        |Dependency                 |
|:-------------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|
|[www.IconArchive.com](http://www.iconarchive.com/show/silk-icons-by-famfamfam.html)|Icons|
|[Snagit](http://discover.techsmith.com/snagit-non-brand-desktop/?gclid=CNzQiOTO09UCFVoFKgod9EIB3g)|Read Me|
|Badges ([Library](https://shields.io/), [Custom](https://rozaxe.github.io/factory/), [Star/Fork](http://githubbadges.com))|Read Me|

<br>

<a id="user-content-glossary-of-terms" class="anchor" href="#glossary-of-terms" aria-hidden="true"> </a>
## Glossary of Terms
| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
|LDAP |Lightweight Directory Access Protocol|
|CSV |Comma-Separated Values|

<br>

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
Listed below is the detailed functionality of this application and its components.  

####	Find Text (Dropdown)
* Lists all the values searched for in the current session

####	Find (Button)
* Queries Active Directory for the text in the "Find Text" textbox

#### Save (Button)
* Saves the current listbox view to a .csv file to the user's desktop

####	User List (Button)
* This will change the layout of the listbox to show more information about each member of a group

####	Settings (Button)
* Opens the settings form. The domain name must be updated here.

####	About (Button)
* Opens the about form
