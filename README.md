# PowerShell-SPCmdlets
SharePoint 2010 and SharePoint 2013 custom PowerShell cmdlets.

# Referenced Assemblies
See [How to Build Office Developer Tools Projects with TFS Team Build 2012](https://msdn.microsoft.com/en-us/library/ff622991.aspx) for information related to using global assembly paths for project builds (Thanks to [@softwarecraft](https://github.com/softwarecraft) for the suggestion and link). If you build from a machine without SharePoint installed then you'll also need to make sure the correct version of the Microsoft.SharePoint.PowerShell assembly and any dependent assemblies are loaded into memory if you want the help files to get generated (I recommend just using a machine with SharePoint installed on it).

## SharePoint 2010 Referenced Assemblies
The SharePoint 2010 project requires access to the SharePoint 2010 (14.0.0.0) version of the following assemblies:
* Microsoft.BusinessData.dll
* Microsoft.Office.Policy.dll
* Microsoft.Office.Server.dll
* Microsoft.Office.Server.Search.Connector.dll
* Microsoft.Office.Server.Search.dll
* Microsoft.Office.Server.UserProfiles.dll
* Microsoft.Office.Word.Server.dll
* Microsoft.SharePoint.dll
* microsoft.sharepoint.portal.dll
* Microsoft.SharePoint.Powershell.dll
* Microsoft.SharePoint.Publishing.dll
* Microsoft.SharePoint.Search.dll
* Microsoft.SharePoint.Security.dll
* Microsoft.SharePoint.Taxonomy.dll
	
If building from a SharePoint 2013 machine then you should copy the SharePoint 2010 assemblies to a global assemblies reference folder and add the following registry key to point to the folder:
> [HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v3.5\AssemblyFoldersEx\SharePoint 2010]@="<AssemblyFolderLocation>"
	
## SharePoint 2013 Referenced Assemblies
The SharePoint 2013 project requires access to the SharePoint 2013 (15.0.0.0) version of the following assemblies:
* Microsoft.BusinessData.dll
* Microsoft.Office.Policy.dll
* Microsoft.Office.Server.dll
* Microsoft.Office.Server.Search.Connector.dll
* Microsoft.Office.Server.Search.dll
* Microsoft.Office.Server.UserProfiles.dll
* Microsoft.Office.Word.Server.dll
* Microsoft.SharePoint.dll
* microsoft.sharepoint.portal.dll
* Microsoft.SharePoint.Powershell.dll
* Microsoft.SharePoint.Publishing.dll
* Microsoft.SharePoint.Search.dll
* Microsoft.SharePoint.Security.dll
* Microsoft.SharePoint.Taxonomy.dll
* Microsoft.Sharepoint.WorkflowActions.dll
	
If building from a SharePoint 2010 machine then you should copy the SharePoint 2013 assemblies to a global assemblies reference folder and add the following registry key to point to the folder:
> [HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.5\AssemblyFoldersEx\SharePoint 2013]@="<AssemblyFolderLocation>"

