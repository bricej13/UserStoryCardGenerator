#########################################
# Brice's fancy TFS queryer
# Version 2.1
# Purpose: Pulls data from TFS and creates an HTML file that will print nicely (from Firefox or Chrome)
# Instructions: 
# 1. Install Team Foundation Power Tools (http://www.microsoft.com/en-us/download/details.aspx?id=35775)
# 2. Change settings
#   - $usid       is a single user story id. Set this variable if you only want the tasks from a single user story.
#   - $stretchstackrank  is the threshold above which user stories will be considered 'stretch'. They print with a slightly different format for easy differentiation.
#   - $query      is the path in TFS for the query you want to run. Note that it includes the project name, the folder structure, and the name of the query
#   - $collection is collection URI. This can be seen in the Team->Connect to Team Foundation Server window
# 3. Run the script!
# 5. Clean up
#   - Go through the new document and reduce the font size for really long transitions
#   
# 
#########################################


# Settings

## $usid: User Story Id. If this id is not 0, we will only get the cards for that user story.
$usid = 0# @(19029) # syntax for multiple stories: @(16329, 12345)

## $stretchstackrank: Stackrank where stretch stories begin. Above this stackrank the cards will print out lighter to distinguish them. Set to 0 to ignore
$stretchstackrank = 999

## $query: Query for USER STORIES only (make sure the query does not include tasks/bugs/children)
$query = "PassportPlus\Team Queries\Passport 2.5\Sprint 4\Sprint 4"

## $collection: URI for tfs server + project
$collection = "http://daltfcsrc02.freemanco.com:8080/tfs/ProjectSimplify"

# End user settings section



function get-tfs-connection {
    # Create connection to TFS server
    param(
        [string] $serverName = $(throw 'serverName is required')
    )

    # Use left-over connection if script exited prematurely
    if ($tfs) {
        $tfs.Dispose()
    }

    # load the required dll
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.TeamFoundation.Client")

    $propertiesToAdd = (
        ('VCS', 'Microsoft.TeamFoundation.VersionControl.Client', 'Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer'),
        ('WIT', 'Microsoft.TeamFoundation.WorkItemTracking.Client', 'Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore'),
        ('CSS', 'Microsoft.TeamFoundation', 'Microsoft.TeamFoundation.Server.ICommonStructureService'),
        ('GSS', 'Microsoft.TeamFoundation', 'Microsoft.TeamFoundation.Server.IGroupSecurityService')
    )

    # fetch the TFS instance, but add some useful properties to make life easier
    # Make sure to "promote" it to a psobject now to make later modification easier
    [psobject] $tfs = [Microsoft.TeamFoundation.Client.TeamFoundationServerFactory]::GetServer($collection)
    foreach ($entry in $propertiesToAdd) {
        $scriptBlock = '
            [System.Reflection.Assembly]::LoadWithPartialName("{0}") > $null
            $this.GetService([{1}])
        ' -f $entry[1],$entry[2]
        $tfs | add-member scriptproperty $entry[0] $ExecutionContext.InvokeCommand.NewScriptBlock($scriptBlock)
    }
    return $tfs
}

function get-userstory-ids($tfs, $workitemquery)  {
    # Returns a list of the ids of the user stories in a query
    try {
        $command = "TFPT.EXE query /format:id `"$query`" /collection:$Collection"
        $ids = Invoke-Expression "$command"
        return $ids
    }
    catch {
        throw [System.Management.Automation.CommandNotFoundException] "Team Foundation Power Tools not installed or not in system PATH (TFPT.exe). Please install them from http://www.microsoft.com/en-us/download/details.aspx?id=35775" 
    }

}

function get-styles()  {
    $style = @"
    <style type='text/css'>
    @page {margin-top: 20px; margin-left: 15px; margin-bottom: 10px; }
    body {font-family: sans-serif; -webkit-print-color-adjust:exact;}
	table { page-break-inside: avoid;font-size: 1.6em;  margin-bottom: 30px; border-style: solid; border-collapse: collapse; border-width: 1px; }
	td { padding: 5px 10px; border-color: white; }
	table.userstory { width: 600px; }
	table.task { width: 300px; float: left; margin-right: 30px;}
	tr.header { background-color: #666666; color: #FFF; }
    tr.stretchheader { background-color: #EEE; color: #666 }
	td.points { text-align: right; width: 33%; font-size: 1.2em; }
	td.stackrank {text-align: left; width: 33%; font-size: 1.2em; }
	td.taskusid {text-align: left; width: 33%; }
	td.usid { text-align: center; font-size: 1.5em; width: 34%; font-weight: bold; }
	td.usdescription { height: 200px; vertical-align: top; }
	td.taskdescription { height: 160px; vertical-align: top; position: relative;}
	/*div.initialbox {border-top: 1px solid gray; border-left: 1px solid gray; position: absolute; bottom: 0px; right: 0px; height: 50px; width: 50px;}*/
	div.pagebreak { page-break-before: always; }
    </style>
"@
    return $style
}

function get-us-html($us) {
    $headerclass = 'header'
    if ($us.StackRank -ge $stretchstackrank) {
        $headerclass = 'stretchheader'
    }

    $table = "<table border=3 class='userstory'><tr class='{5}'>	<td class='stackrank'>{0}</td><td class='usid'>{1}</td><td class='points'>{2}-{3}pts</td>
    </tr><tr><td colspan=3 class='usdescription'>{4}</td></tr></table>" -f $us.StackRank, $us.Id, $us.FiftyPercent, $us.NinetyPercent, $us.Title, $headerclass
    return $table
}

function get-task-html($t) {
    $table = "<table class='task'><tr class='header'><td class='taskusid'>US{0}</td></tr><tr>
    <td class='taskdescription'>{1}<div class='initialbox'> </div></td></tr></table>" -f $t.Id, $t.Title
    return $table
}

function export-to-html($userstories, $tasks) {
    $pagehtml = "<html><head>" 
    $pagehtml += get-styles
    $pagehtml += "</head><body>"

    foreach ($u in $userstories) {
       $pagehtml += get-us-html($u)
    }
    $pagehtml += "<div class='pagebreak'></div>"

    foreach ($t in $tasks) {
        if ($t) {
            $pagehtml += get-task-html($t)
        }
    }
    return $pagehtml += "</body></html>"
}

function sort-tasks($tasks) {
    $sorted = @(0)
    for ($i = 0; $i -le $tasks.Length * 2; $i++) {
        $sorted += 0
    }
    $previd = 0
    $pagesize = 8
    $startindex = 0

    foreach ($task in $tasks) {
        $index = 0
        if ($task.Id -ne $previd) {
            write-host "New User Story: " + $task.Id
            while ($sorted[$index]) {
                $index += 1
            }
            $startindex = $index
        }
        else {
            $index = $startindex
            while ($sorted[$index]) {
                $index = $index + $pagesize
            }
            
            
        }
        write-host "Setting index " + $index + " to " + $task.Id + ": " + $task.Description
        $sorted[$index] = $task
        $previd = $task.Id
    }
    return $sorted
}

$tfs = get-tfs-connection($collection)

# Get the user story ids
if ($usid) {
    $ids = $usid
}
else {
    $ids = get-userstory-ids($tfs, $query)
}


$userstories = @()
$workitems = @()

# Get data & tasks for each User Story
foreach ($id in $ids) {

    $us = $tfs.wit.GetWorkItem($id)

    "Gathering data for " + $us.Id
    # Uncomment this line for a full list of available fields
    # $us.Fields | Select ReferenceName
    $stackrank   = $us.Fields | where {$_.ReferenceName -eq "Microsoft.VSTS.Common.StackRank"} | select value
    $storypoints = $us.Fields | where {$_.ReferenceName -eq "Microsoft.VSTS.Scheduling.StoryPoints"} | select value
    $90percent = $us.Fields | where {$_.ReferenceName -eq "Freeman.Common.90StoryPoints"} | select value
    $50percent = $us.Fields | where {$_.ReferenceName -eq "Freeman.Common.50StoryPoints2"} | select value

    $newus = New-Object System.Object
    $newus | Add-Member -type NoteProperty -name Title -value $us.Title
    $newus | Add-Member -type NoteProperty -name Id -value $us.Id
    $newus | Add-Member -type NoteProperty -name StoryPoints -value $storypoints.Value
    $newus | Add-Member -type NoteProperty -name StackRank -value $stackrank.Value
    $newus | Add-Member -type NoteProperty -name FiftyPercent -value $50Percent.Value
    $newus | Add-Member -type NoteProperty -name NinetyPercent -value $90Percent.Value
    
    $userstories += $newus
    
    foreach ($link in $us.WorkItemLinks) {
        $wi = $tfs.wit.GetWorkItem($link.TargetId)
        
        $t = $wi | select Type
        
        if ($t.Type.Name.Equals("Task")) {
            $newwi = New-Object System.Object
            $newwi | Add-Member -type NoteProperty -name Id -value $newus.Id
            $newwi | Add-Member -type NoteProperty -name Title -value $wi.Title

            $workitems += $newwi
        }
    }
}

$HTMLSaveLocation = "$env:temp\SprintCards-$(Get-Date -format 'yyyy-MM-dd hh-mm-ss').html"

$userstories = $userstories | Sort-Object Id
# $workitems = sort-tasks $workitems


export-to-html $userstories $workitems | Out-File $HTMLSaveLocation

"Saved file to: ", $HTMLSaveLocation

invoke-item $HTMLSaveLocation



# Get rid of TFS connection
$tfs.Dispose();
Clear-Variable tfs


# Version History
## 2.2
# - Changed sorting on tasks for easier cutting. Now one task per story will be on each page. This should help sorting when using a paper cutter
## 2.1
# - A limit can be now set to differentiate stretch stories. They will print with a lighter-colored header for easy differentiation
#
## 2.0
# - Changed out to HTML directly
# - Fixed bug that caused issues if script was exited prematurely
# - Added functionality to print a single user story
