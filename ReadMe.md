# Photo Import/Export tool for Active Directory and Office 365

Photos can be a pain. This tool is for people looking for a simple to use command line app to Export/Import photos from Active Directory, SharePoint Online or Exchange Online. 

It's simple. When it exports it drops the photos in /photos. It also keep a list of users and photos in UserList.json which has their UPN and the location of their photo. 

Example use cases:

* You want to export all those nice high resolutions photos from Office 365 (Exchange Online which is where Skype/Outlook stores photos) and import them to your Active Directory (sized appropriately of course)
* You want to export of all the photos from SharePoint Online for some other system you are working on.
* You want to bulk update the photos for your organization in Exchange Online from that folder that the HR team just gave you. 

It's a work in progress :)

## Versions

### v1.2.0 (2018-12-19)
- bug fix to set profile properties using API on admin site collection vs. mysite [\#2](https://github.com/Hyperfish/PhotoImporterExporter/pull/2)