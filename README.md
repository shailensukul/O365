# O365

This repository contains projects and POCs for all things relating to Office 365 and SharePoint Online.
Please read further below for individual projects:

## Office 365 Documentor 
Project: [O365.Documentor] (https://github.com/shailensukul/Office365/tree/master/O365.Documentor)

Screenshot:

![Screenshot][logo]

[logo]: https://github.com/shailensukul/Office365/blob/master/O365.Documentor/O365Documentor.png "Office 365 Documentor" 

This console app does the following:
* Export site inventory of a site collection to a CSV file. You can feed this file into a Visio org chart to create a visual representation of your sites
* Export site inventory of a site collection to an XML file. Features per site are also exported, based on the FeatureNameFilter configuration key in the app.config file. You can use this file to recreate the Sites in another site collection. (Will release code for this shortly)
* Content Types Inventory - creates a CSV report of all the content types in the Site Collection which match a Group (group name is configured in the ContentTypeGroupName key in the app config file)
* Site Columns Inventory Report - creates a CSV report of all the site columns in the Site Collection which match a Group (group name is configured in the ColumnGroupName key in the app config file)
* PageLayouts Inventory Report - creates a CSV report of all the page layouts in the Site Collection.

## SharePoint Addin with ADAL Security
Project: Security.SharePoint.AddIn
TODO

[Markdown Cheatsheet](https://github.com/adam-p/markdown-here/wiki/Markdown-Here-Cheatsheet)
