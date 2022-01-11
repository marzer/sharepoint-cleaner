# sharepoint-cleaner
Bulk-cleans version history from a SharePoint server. Because apparently giving administrators a UI for this is challenging.

<br><br>

## Usage
sharepoint-cleaner is a command-line application requring .NET 4.8.
```
sharepoint-cleaner [site_uri] [username]
```

`site_uri` and `username` are optional; you will be prompted to input them if they are omitted.
Password is always provided at a prompt (i.e. it cannot be specified on the command-line).

Since the many server requests needed can take a lot of time (potentially hours), the application stores the session state locally in a file.
Re-running the application again with the same `site_uri` and `username` will attempt to resume the session from the
file to avoid duplicate work being done.

<br><br>

## Example run
```
---------------------------------------------------------
sharepoint-cleaner - github.com/marzer/sharepoint-cleaner
---------------------------------------------------------
Site URI: https://foobar.sharepoint.com/Assets
Username: foo@bar.com
Password: *************
Session: sharepoint-cleaner_4056058362.xml
---------------------------------------------------------
/Assets/_vti_pvt
/Assets/Access Requests
/Assets/SitePages
/Assets/SitePages/Forms
/Assets/SitePages/Forms/Repost Page
/Assets/SitePages/Forms/Site Page
/Assets/SitePages/Forms/Web Part Page
/Assets/SiteAssets
/Assets/SiteAssets/Assets Notebook
/Assets/SiteAssets/Forms
/Assets/SiteAssets/Forms/Document
/Assets/Lists
/Assets/Lists/PublishedFeed
/Assets/Lists/PublishedFeed/AB922B82-8406-4E49-B17B-9057BDF09503
/Assets/Lists/PublishedFeed/Attachments
/Assets/Lists/PublishedFeed/FEB96200-6E92-41DB-856B-E8702BCDF33A
/Assets/_catalogs
/Assets/_catalogs/masterpage
/Assets/_catalogs/masterpage/Forms
/Assets/_catalogs/masterpage/Forms/Document
/Assets/_catalogs/masterpage/Forms/MasterPage
/Assets/_catalogs/design
/Assets/_catalogs/design/Attachments
/Assets/images
/Assets/Shared Documents
/Assets/Shared Documents/Video
/Assets/Shared Documents/Video/RenderInPlace
/Assets/Shared Documents/Video/XRay
/Assets/Shared Documents/Video/XRay/Tampere olecranon 1
/Assets/Shared Documents/Video/XRay/Tampere olecranon 1/Session1_20211109
/Assets/Shared Documents/Video/XRay/Tampere olecranon 1/Vänni_-_-
/Assets/Shared Documents/Video/XRay/Tampere olecranon 1/Vänni_-_-/Session1_20211109
Session written to sharepoint-cleaner_4056058362.xml
---------------------------------------------------------
Processed 105 files and 29 folders, deleting 92 past versions.
```

<br><br>

## License and Attribution
This project is published under the terms of the [MIT license](https://github.com/marzer/sharepoint-cleaner/blob/main/LICENSE.txt).

