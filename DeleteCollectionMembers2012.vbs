'***************************************************
'Name:		DeleteCollectionMembers2012.vbs
'Date:		20-Jan-2012
'Orig. Author:	Jim Rothe
'Modifications:	Henry E. Wilson
'		Included removing the arguments and including a count 
'		of resources deleted that is reported in the Application
'		Event log.
'
'Purpose:	To automate the DeleteSpecial command on a hardcoded
'		Site/Collection.  Generally these collections contain
'		obsolete resources,  resources with no client 
'		or decommissioned resources.  This script can be used 
'		as Scheduled task
'
' Modified:  13-Dec-2016
'		Change Server, SiteSode and Collection ID to point to the 
'		Delete Special collection in SCCM 2012 Production
'
'***************************************************
Option Explicit
On Error Resume Next

Dim wshShell, objLocator, objServer, objSite
Dim strColl, ResCount, objColl
Dim objService, strServer, strComputer

Const EVENT_SUCCESS = 0
Const EVENT_FAIL = 1

'Environment information must be provided here
strServer = "XSPW10W200P"
objSite = "P00"
'This is the CollectionID for "Delete Special" collection under the Maintenance Collection
objColl = "P00029E5"

Set wshShell = wscript.CreateObject("wscript.shell")

Set objLocator = CreateObject("WbemScripting.SWbemLocator")
'Connects to the namespace of the site server
Set objService = objLocator.ConnectServer(strServer,"root\sms\site_" & objSite)

'Verifies connectivity with the namespace
If Err.Number<>0 Then
	'writes a failure event in the application event log
	wshShell.LogEvent EVENT_FAIL, "Could not locate site " & objSite
	wscript.quit
End If	

'Logs the start of the script in the application event log
wshShell.LogEvent EVENT_SUCCESS, "DeleteCollectionMembers2012.vbs has started" & VBCrLf & _
"Site code = " & objSite & VBCrLf & "Collection ID = " & objColl

'Gets the instance of the collection
Set strColl = objService.Get("SMS_Collection='" & objColl & "'")

'Verifies existence of the collection
If Err.Number<>0 Then
	'writes a failure event in the application event log
	wshShell.LogEvent EVENT_FAIL, "Could not locate 2012 collection " & objColl
	wscript.quit
 End If	

'Get the count of resources to be deleted
Set ResCount = objService.execquery("select * from SMS_CM_RES_COLL_" &objColl)

'Write the count to the Application Event Log
wshShell.LogEvent EVENT_SUCCESS, "There are " & ResCount.Count & " resources in the 2012 collection to be deleted.  Collection ID: " & objColl & _
" in site " & objSite

'Deletes all members of collection
strColl.DeleteAllMembers


'Write a completed entry in the Application Event Log
wshShell.LogEvent EVENT_SUCCESS, "The Delete Special of " & ResCount.Count & " resources in the 2012 collection have been deleted.  Collection ID: " & objColl & _
" in site " & objSite

