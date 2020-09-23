<div align="center">

## VB\.NET BrowseForFolder Dialog


</div>

### Description

UPDATED 11JUN02 - Sorry It took me so long to fix this code after the site had been hacked. Regardless, here is the VB.NET code. | This will show the Browse For Folder dialog.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[bleh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bleh.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB\.NET
**Category**       |[GUIs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/guis__10-30.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bleh-vb-net-browseforfolder-dialog__10-284/archive/master.zip)





### Source Code

<font face="Verdana" size="2">
<div align="center"><b>Extending The Browse For Folder Dialog Class For VB.NET</b></div>
<p>This code is simpily extending the Browse For Folder Dialog class that was on posted on PSC by Chris Andersen. After using his code, I was interested in knowing if there was any additional options available aside from just being able to set the title. I consulted the MSDN, and posted it under the C# section of PSC. Here is VB.NET version of that code. The above screen shot is showing the Browse For Folder dialog with StartLocation set to "MyDocuments" and Style set to "BrowseForEverything"</p>
<p><b>.StartLocation</b><br>
<hr width="100%" size="1" color="#000000">
The directory in which the browse dialog will start at. (duh) I haven't found a way to specifiy a certain path, like C:\PSC or something of that nature, but there are a number of members that we can use. They are:
<ul>
	<li>Desktop</li>
	<li>Favorites</li>
	<li>MyComputer</li>
	<li>MyDocuments</li>
	<li>MyPictures</li>
	<li>NetAndDialUpConnections</li>
	<li>NetworkNeighborhood</li>
	<li>Printers</li>
	<li>Recent</li>
	<li>SendTo</li>
	<li>StartMenu</li>
	<li>Templates</li>
</ul>
<p><b>.Style</b><br>
<hr width="100%" size="1" color="#000000">
They style of the browse dialog. We have a few different options to choose from.
<ul>
	<li>BrowseForComputer</li>
	<li>BrowseForEverything</li>
	<li>BrowseForPrinter</li>
	<li>RestrictToDomain</li>
	<li>RestrictToFilesystem</li>
	<li>RestrictToSubfolders</li>
	<li>ShowTextBox</li>
</ul>
</p>
<p><b>The Code</b><br>
<hr width="100%" size="1" color="#000000">
Here's an example of how to use the above methods in the Browse For Folder dialog.
<pre><font size="2">
Public Class BrowseForFolder
  Inherits System.Windows.Forms.Design.FolderNameEditor
  Dim bDialog As New FolderBrowser()
  Public Function BrowseDialog(ByVal sTitle As String)
   With bDialog
    .Style = FolderBrowserStyles.BrowseForEverything
    .StartLocation = FolderBrowserFolder.Desktop
    .Description = sTitle
    .ShowDialog()
    BrowseDialog = .DirectoryPath()
   End With
  End Function
 End Class
</font></pre>
</p>
<p><b>Useage:</b><br>
<hr width="100%" size="1" color="#000000">
Call the class like so:
<pre><font size="2">
Dim myDialog As New BrowseForFolder()
MsgBox(myDialog.BrowseDialog("My Dialog Title Here"))
</pre>
<b>Make sure you add a reference to the System.Windows.Forms.dll in your Solution Explorer!</b>
It's pretty simple. Thanks again to Chris Anderson for the or

