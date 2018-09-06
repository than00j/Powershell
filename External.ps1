Function Set-WINRMListener
{
<#
  .SYNOPSIS
    Set-WinRMListener will configure a remote computer's WinRM service to listen for WinRM requests

  .DESCRIPTION
    Set-WinRMListener uses the system.management.managementclass object to create 3 registry keys.  These registry entries will configure the WinRM service when it restarts next.  Use Restart-WinRM to restart the WinRM service after it has been successfully configured.

    Instead of :

    psexec \\[computer name] -u [admin account name] -p [admin account password] -h -d powershell.exe "enable-psremoting -force"
    
    OR

    $cmd1 = "cmd /c C:\Source\sysinternals\PsExec.exe \\$s -s winrm.cmd quickconfig -q"
    $cmd2 = "cmd /c C:\Source\sysinternals>PsExec.exe \\$s -s powershell.exe enable-psremoting -force"
    Invoke-Expression -Command $cmd1
    Invoke-Expression -Command $cmd2

  .EXAMPLE 
    Get-Content .\servers.txt | Set-WinRMListener 

    Sets the WinRM service on the computers in the servers.txt file

  .EXAMPLE 
    Set-WinRMListener -ComputerName (get-content .\servers.txt) -IPv4Range "10.10.1.0 255.255.255.0"

    Sets the WinRM service on the computers in the servers.txt file and restricts the IPv4 address to 10.10.1.0 /24

#>
[cmdletBinding()]
Param
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName,
    
    # Enter the IPv4 address range for the WinRM listener
    [Parameter(
    HelpMessage='Enter the IPv4 address range for the WinRM listener')]
    [String]$IPv4Range = '*',
    
    # Enter the IPv4 address range for the WinRM listener
    [Parameter(
    HelpMessage='Enter the IPv4 address range for the WinRM listener')]
    [String]$IPv6Range = '*'
  )
Begin
  {
    Write-Debug 'Opening Begin block'
    $HKLM = 2147483650
    $Key = 'SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
    $DWORDName = 'AllowAutoConfig' 
    $DWORDvalue = '0x1'
    $String1Name = 'IPv4Filter'
    $String2Name = 'IPv6Filter'
    Write-Debug 'Finished begin block variables are built'
  }
Process
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to create remote registry handle'
            $Reg = New-Object -TypeName System.Management.ManagementClass -ArgumentList \\$computer\Root\default:StdRegProv
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to create Remote Key'
            if (($reg.CreateKey($HKLM, $key)).returnvalue -ne 0) {Throw 'Failed to create key'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attemping to set DWORD value'
            if (($reg.SetDWORDValue($HKLM, $Key, $DWORDName, $DWORDvalue)).ReturnValue -ne 0) {Throw 'Failed to set DWORD'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set first REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $String1Name, $IPv4Range)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set second REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $String2Name, $IPv6Range)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }
  }
End {}
}

Function Restart-WinRM
{
<#
  .SYNOPSIS
    Restarts the WinRM service

  .DESCRIPTION
    Uses the win32_service class to get and restart the WinRM service.  This function was designed to be used after Set-WinRMListener to allow the new registry configuration to take hold.

  .EXAMPLE
    Restart-WinRM -computername TestVM

    Restarts the WinRM service on TestVM

  .EXAMPLE
    Get-Content .\servers.txt | Restart-WinRM 

    Restarts the WinRM service on all the computers in the servers.txt file

#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to stop WinRM'
            (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StopService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Stopped') {Throw 'Failed to Stop WinRM'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to start WinRM'
            (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StartService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Running') {Throw 'Failed to Start WinRM'}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
End {}
}

Function Set-WinRMStartup
{
<#
  .SYNOPSIS
    Changes the startup type of the WinRM service to automatic

  .DESCRIPTION
    Uses the Win32_service class to change the startup type on the WinRM service to Automatic

  .EXAMPLE
    Set-WinRMStartup -Computername TestVM

    Sets the WinRM service startup type to Automatic on test VM

#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            if (((Get-WmiObject win32_service -Filter "name='WinRM'" -ComputerName $computer).ChangeStartMode('Automatic')).ReturnValue -ne 0) {Throw "Failed to change WinRM Startup type on $Computer"}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "Failed to change startmode on the WinRM service for $computer"
            break
          }
        Write-Verbose "Completed operations on $computer"
        Write-Debug "Finished process block for $Computer"
      }
  }
}

Function Set-WinRMFirewallRule
{
<#
  .SYNOPSIS
    Set-WinRMFirewallRule will configure a remote computer's firewall

  .DESCRIPTION
    Set-WinRMFirewallRule uses the system.management.managementclass object to create 2 registry keys.  These registry entries will configure the Windows Firewall service when it nest restarts.  Use Restart-WindowsFirewall to restart the Windows Firewall service after it has been successfully configured.  Only use this function if you're using the windows firewall.

  .EXAMPLE 
    Get-Content .\servers.txt | Set-WinRMFirewallRule 

    Sets the rules for the windows firewall on the computers in the servers.txt file

  .EXAMPLE 
    Set-WinRMFirewallRule -ComputerName (get-content .\servers.txt)

    Sets the rules for the windows firewall on the computers in the servers.txt file

#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin 
  {
    Write-Debug 'Opening Begin block'
    $HKLM = 2147483650
    $Key = 'SOFTWARE\Policies\Microsoft\WindowsFirewall\FirewallRules'
    $Rule1Value = 'v2.20|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Public|LPort=5985|RA4=LocalSubnet|RA6=LocalSubnet|App=System|Name=@FirewallAPI.dll,-30253|Desc=@FirewallAPI.dll,-30256|EmbedCtxt=@FirewallAPI.dll,-30267|'
    $Rule1Name = 'WINRM-HTTP-In-TCP-PUBLIC'
    $Rule2Value = 'v2.20|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Domain|Profile=Private|LPort=5985|App=System|Name=@FirewallAPI.dll,-30253|Desc=@FirewallAPI.dll,-30256|EmbedCtxt=@FirewallAPI.dll,-30267|'
    $Rule2Name = 'WINRM-HTTP-In-TCP'
    Write-Debug 'Completed Begin block, Finshed creating variables'
  }
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to create remote registry handle'
            $Reg = New-Object -TypeName System.Management.ManagementClass -ArgumentList \\$computer\Root\default:StdRegProv
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to create Remote Key'
            if (($reg.CreateKey($HKLM, $key)).returnvalue -ne 0) {Throw 'Failed to create key'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set first REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $Rule1Name, $Rule1Value)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set second REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $Rule2Name, $Rule2Value)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
}

Function Restart-WindowsFirewall
{
<#
  .SYNOPSIS
    Restarts the windows firewall service

  .DESCRIPTION
    Uses the Win32_Service class to restart the windows firewall service.

  .EXAMPLE
    Restart-WindowsFirewall -computername TestVM

    Restarts the Windows Firewall service on TestVM

  .EXAMPLE
    Get-Content .\servers.txt | Restart-WindowsFirewall 

    Restarts the Windows Firewall service on all the computers in the servers.txt file

#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to stop MpsSvc'
            (Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).StopService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).state -notlike 'Stopped') {Throw 'Failed to Stop MpsSvc'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to start MpsSvc'
            (Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).StartService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).state -notlike 'Running') {Throw 'Failed to Start MpsSvc'}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
End {}
}

===============

#function ShowColorizedContent{
#requires -version 2.0

param(
    $filename = $(throw "Please specify a filename."),
    $highlightRanges = @(),
    [System.Management.Automation.SwitchParameter] $excludeLineNumbers)

# [Enum]::GetValues($host.UI.RawUI.ForegroundColor.GetType()) | % { Write-Host -Fore $_ "$_" }
$replacementColours = @{ 
    "Command"="Yellow";
    "CommandParameter"="Yellow";
    "Variable"="Green" ;
    "Operator"="DarkCyan";
    "Grouper"="DarkCyan";
    "StatementSeparator"="DarkCyan";
    "String"="Cyan";
    "Number"="Cyan";
    "CommandArgument"="Cyan";
    "Keyword"="Magenta";
    "Attribute"="DarkYellow";
    "Property"="DarkYellow";
    "Member"="DarkYellow";
    "Type"="DarkYellow";
    "Comment"="Red";
}
$highlightColor = "Green"
$highlightCharacter = ">"

## Read the text of the file, and parse it
$file = (Resolve-Path $filename).Path
$content = [IO.File]::ReadAllText($file)
$parsed = [System.Management.Automation.PsParser]::Tokenize($content, [ref] $null) | 
    Sort StartLine,StartColumn

function WriteFormattedLine($formatString, [int] $line)
{
    if($excludeLineNumbers) { return }
    
    $hColor = "Gray"
    $separator = "|"
    if($highlightRanges -contains $line) { $hColor = $highlightColor; $separator = $highlightCharacter }
    Write-Host -NoNewLine -Fore $hColor ($formatString -f $line,$separator)
}

Write-Host

WriteFormattedLine "{0:D3} {1} " 1

$column = 1
foreach($token in $parsed)
{
    $color = "Gray"

    ## Determine the highlighting colour
    $color = $replacementColours[[string]$token.Type]
    if(-not $color) { $color = "Gray" }

    ## Now output the token
    if(($token.Type -eq "NewLine") -or ($token.Type -eq "LineContinuation"))
    {
        $column = 1
        Write-Host

	WriteFormattedLine "{0:D3} {1} " ($token.StartLine + 1)
    }
    else
    {
        ## Do any indenting
        if($column -lt $token.StartColumn)
        {
            Write-Host -NoNewLine (" " * ($token.StartColumn - $column))
        }

        ## See where the token ends
        $tokenEnd = $token.Start + $token.Length - 1

        ## Handle the line numbering for multi-line strings
        if(($token.Type -eq "String") -and ($token.EndLine -gt $token.StartLine))
        {
            $lineCounter = $token.StartLine
            $stringLines = $(-join $content[$token.Start..$tokenEnd] -split "`r`n")
            foreach($stringLine in $stringLines)
            {
                if($lineCounter -gt $token.StartLine)
                {
                    WriteFormattedLine "`n{0:D3} {1}" $lineCounter
                }
                Write-Host -NoNewLine -Fore $color $stringLine
                $lineCounter++
            }
        }
        ## Write out a regular token
        else
        {
            Write-Host -NoNewLine -Fore $color (-join $content[$token.Start..$tokenEnd])
        }

        ## Update our position in the column
        $column = $token.EndColumn
    }
}

Write-Host "`n" 
#}

========================

function Show-Object{
#############################################################################
##
## Show-Object
##
## From Windows PowerShell Cookbook (O'Reilly)
## by Lee Holmes (http://www.leeholmes.com/guide)
##
##############################################################################

<#

.SYNOPSIS

Provides a graphical interface to let you explore and navigate an object.


.EXAMPLE

PS > $ps = { Get-Process -ID $pid }.Ast
PS > Show-Object $ps

#>

param(
    ## The object to examine
    [Parameter(ValueFromPipeline = $true)]
    $InputObject
)

Set-StrictMode -Version 3

Add-Type -Assembly System.Windows.Forms

## Figure out the variable name to use when displaying the
## object navigation syntax. To do this, we look through all
## of the variables for the one with the same object identifier.
$rootVariableName = dir variable:\* -Exclude InputObject,Args |
    Where-Object {
        $_.Value -and
        ($_.Value.GetType() -eq $InputObject.GetType()) -and
        ($_.Value.GetHashCode() -eq $InputObject.GetHashCode())
}

## If we got multiple, pick the first
$rootVariableName = $rootVariableName| % Name | Select -First 1

## If we didn't find one, use a default name
if(-not $rootVariableName)
{
    $rootVariableName = "InputObject"
}

## A function to add an object to the display tree
function PopulateNode($node, $object)
{
    ## If we've been asked to add a NULL object, just return
    if(-not $object) { return }

    ## If the object is a collection, then we need to add multiple
    ## children to the node
    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object))
    {
        ## Some very rare collections don't support indexing (i.e.: $foo[0]).
        ## In this situation, PowerShell returns the parent object back when you
        ## try to access the [0] property.
        $isOnlyEnumerable = $object.GetHashCode() -eq $object[0].GetHashCode()

        ## Go through all the items
        $count = 0
        foreach($childObjectValue in $object)
        {
            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $newChildNode = New-Object Windows.Forms.TreeNode
            $newChildNode.Text = "$($node.Name)[$count] = $childObjectValue : " +
                $childObjectValue.GetType()

            ## Use the node name to keep track of the actual property name
            ## and syntax to access that property.
            ## If we can't use the index operator to access children, add
            ## a special tag that we'll handle specially when displaying
            ## the node names.
            if($isOnlyEnumerable)
            {
                $newChildNode.Name = "@"
            }

            $newChildNode.Name += "[$count]"
            $null = $node.Nodes.Add($newChildNode)               

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $newChildNode $childObjectValue

            $count++
        }
    }
    else
    {
        ## If the item was not a collection, then go through its
        ## properties
        foreach($child in $object.PSObject.Properties)
        {
            ## Figure out the value of the property, along with
            ## its type.
            $childObject = $child.Value
            $childObjectType = $null
            if($childObject)
            {
                $childObjectType = $childObject.GetType()
            }

            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $childNode = New-Object Windows.Forms.TreeNode
            $childNode.Text = $child.Name + " = $childObject : $childObjectType"
            $childNode.Name = $child.Name
            $null = $node.Nodes.Add($childNode)

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $childNode $childObject
        }
    }
}

## A function to add a placeholder if required to a node.
## If there are any properties or children for this object, make a temporary
## node with the text "..." so that the node shows a '+' sign to be
## expanded.
function AddPlaceholderIfRequired($node, $object)
{
    if(-not $object) { return }

    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object) -or
        @($object.PSObject.Properties))
    {
        $null = $node.Nodes.Add( (New-Object Windows.Forms.TreeNode "...") )
    }
}

## A function invoked when a node is selected.
function OnAfterSelect
{
    param($Sender, $TreeViewEventArgs)

    ## Determine the selected node
    $nodeSelected = $Sender.SelectedNode

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $nodeSelected

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    $resultObject = Invoke-Expression $nodePath
    $outputPane.Text = $nodePath

    ## If we got some output, put the object's member
    ## information in the text box.
    if($resultObject)
    {
        $members = Get-Member -InputObject $resultObject | Out-String       
        $outputPane.Text += "`n" + $members
    }
}

## A function invoked when the user is about to expand a node
function OnBeforeExpand
{
    param($Sender, $TreeViewCancelEventArgs)

    ## Determine the selected node
    $selectedNode = $TreeViewCancelEventArgs.Node

    ## If it has a child node that is the placeholder, clear
    ## the placeholder node.
    if($selectedNode.FirstNode -and
        ($selectedNode.FirstNode.Text -eq "..."))
    {
        $selectedNode.Nodes.Clear()
    }
    else
    {
        return
    }

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $selectedNode 

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    Invoke-Expression "`$resultObject = $nodePath"

    ## And populate the node with the result object.
    PopulateNode $selectedNode $resultObject
}

## A function to handle keypresses on the form.
## In this case, we capture ^C to copy the path of
## the object property that we're currently viewing.
function OnKeyPress
{
    param($Sender, $KeyPressEventArgs)

    ## [Char] 3 = Control-C
    if($KeyPressEventArgs.KeyChar -eq 3)
    {
        $KeyPressEventArgs.Handled = $true

        ## Get the object path, and set it on the clipboard
        $node = $Sender.SelectedNode
        $nodePath = GetPathForNode $node
        [System.Windows.Forms.Clipboard]::SetText($nodePath)

        $form.Close()
    }
}

## A function to walk through the parents of a node,
## creating virtual PowerShell syntax to access this property.
function GetPathForNode
{
    param($Node)

    $nodeElements = @()

    ## Go through all the parents, adding them so that
    ## $nodeElements is in order.
    while($Node)
    {
        $nodeElements = ,$Node + $nodeElements
        $Node = $Node.Parent
    }

    ## Now go through the node elements
    $nodePath = ""
    foreach($Node in $nodeElements)
    {
        $nodeName = $Node.Name

        ## If it was a node that PowerShell is able to enumerate
        ## (but not index), wrap it in the array cast operator.
        if($nodeName.StartsWith('@'))
        {
            $nodeName = $nodeName.Substring(1)
            $nodePath = "@(" + $nodePath + ")"
        }
        elseif($nodeName.StartsWith('['))
        {
            ## If it's a child index, we don't need to
            ## add the dot for property access
        }
        elseif($nodePath)
        {
            ## Otherwise, we're accessing a property. Add a dot.
            $nodePath += "."
        }

        ## Append the node name to the path
        $nodePath += $nodeName
    }

    ## And return the result
    $nodePath
}

## Create the TreeView, which will hold our object navigation
## area.
$treeView = New-Object Windows.Forms.TreeView
$treeView.Dock = "Top"
$treeView.Height = 500
$treeView.PathSeparator = "."
$treeView.Add_AfterSelect( { OnAfterSelect @args } )
$treeView.Add_BeforeExpand( { OnBeforeExpand @args } )
$treeView.Add_KeyPress( { OnKeyPress @args } )

## Create the output pane, which will hold our object
## member information.
$outputPane = New-Object System.Windows.Forms.TextBox
$outputPane.Multiline = $true
$outputPane.ScrollBars = "Vertical"
$outputPane.Font = "Consolas"
$outputPane.Dock = "Top"
$outputPane.Height = 300

## Create the root node, which represents the object
## we are trying to show.
$root = New-Object Windows.Forms.TreeNode
$root.Text = "$InputObject : " + $InputObject.GetType()
$root.Name = '$' + $rootVariableName
$root.Expand()
$null = $treeView.Nodes.Add($root)

## And populate the initial information into the tree
## view.
PopulateNode $root $InputObject

## Finally, create the main form and show it.
$form = New-Object Windows.Forms.Form
$form.Text = "Browsing " + $root.Text
$form.Width = 1000
$form.Height = 800
$form.Controls.Add($outputPane)
$form.Controls.Add($treeView)
$null = $form.ShowDialog()
$form.Dispose()
}

=========================

Set-StrictMode -Off

# load Forms NameSpace  
[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")   
   
#region Make the form  
$frmMain = new-object Windows.Forms.form    
$frmMain.Size = new-object System.Drawing.Size @(800,600)    
$frmMain.text = "All In All Azhguraja"
#endregion Make the form  
#region Define Used Controls  
$MainMenu = new-object System.Windows.Forms.MenuStrip  
$statusStrip = new-object System.Windows.Forms.StatusStrip  
$FileMenu = new-object System.Windows.Forms.ToolStripMenuItem  
$ToolMenu = new-object System.Windows.Forms.ToolStripMenuItem('&Tools')  
$miQuery = new-object System.Windows.Forms.ToolStripMenuItem('&Query (run)')  
$miSelectQuery = new-object System.Windows.Forms.ToolStripMenuItem('&SelectQuery')  
$miSelectQuery.add_Click({$sq | out-propertyGrid;$wmiSearcher.Query = $sq})  
[void]$ToolMenu.DropDownItems.Add($miSelectQuery)  
$miRelatedObjectQuery = new-object System.Windows.Forms.ToolStripMenuItem('&RelatedObjectQuery')  
$miRelatedObjectQuery.add_Click({$roq | out-propertyGrid;$wmiSearcher.Query = $roq})  
[void]$ToolMenu.DropDownItems.Add($miRelatedObjectQuery)  
$miRelationshipQuery = new-object System.Windows.Forms.ToolStripMenuItem('&RelationshipQuery')  
$miRelationshipQuery.add_Click({$rq | out-propertyGrid ;$wmiSearcher.Query = $rq})  
[void]$ToolMenu.DropDownItems.Add($miRelationshipQuery)  
$oq = new-object System.Management.ObjectQuery  
$eq = new-object System.Management.EventQuery  
$sq = new-object System.Management.SelectQuery  
$roq = new-object System.Management.RelatedObjectQuery  
$rq = new-object System.Management.RelationshipQuery  
$wmiSearcher = [wmisearcher]''  
[void]$ToolMenu.DropDownItems.Add($miQuery)  
$miQuery.add_Click({  
    $wmiSearcher | out-propertyGrid  
    $moc = $wmiSearcher.get()  
    $DT =  new-object  System.Data.DataTable  
    $DT.TableName = $lblClass.text  
    $Col =  new-object System.Data.DataColumn  
    $Col.ColumnName = "WmiPath"  
    $DT.Columns.Add($Col)  
    $i = 0  
    $j = 0 ;$lblInstances.Text = $j; $lblInstances.Update()  
    $MOC |  
    ForEach-Object {  
        $j++ ;$lblInstances.Text = $j; $lblInstances.Update()  
        $MO = $_  
         
        # Make a DataRow  
        $DR = $DT.NewRow()  
        $Col =  new-object System.Data.DataColumn  
        $DR.Item("WmiPath") = $mo.__PATH  
        $MO.psbase.properties |  
        ForEach-Object {  
         
            $prop = $_  
             
            If ($i -eq 0)  {  
     
                # Only On First Row make The Headers  
                 
                $Col =  new-object System.Data.DataColumn  
                $Col.ColumnName = $prop.Name.ToString()  
   
                $prop.psbase.Qualifiers |  
                ForEach-Object {  
                    If ($_.Name.ToLower() -eq "key") {  
                        $Col.ColumnName = $Col.ColumnName + "*"  
                    }  
                }  
                $DT.Columns.Add($Col)   
            }  
             
            # fill dataRow   
             
            if ($prop.value -eq $null) {  
                $DR.Item($prop.Name) = "[empty]"  
            } ElseIf ($prop.IsArray) {  
                $DR.Item($prop.Name) =[string]::Join($prop.value ,";")  
            } Else {  
                $DR.Item($prop.Name) = $prop.value  
                #Item is Key try again with *  
                trap{$DR.Item("$($prop.Name)*") = $prop.Value.tostring();continue}  
            }  
        } #end ForEach  
        # Add the row to the DataTable  
         
        $DT.Rows.Add($DR)  
        $i += 1  
    }  
    $DGInstances.DataSource = $DT.psObject.baseobject    
    $status.Text = "Retrieved $j Instances"  
    $status.BackColor = 'YellowGreen'  
    $statusstrip.Update()  
})#$miQuery.add_Click  
 
$miQuit = new-object System.Windows.Forms.ToolStripMenuItem('&Quit')  
$miQuit.add_Click({$frmMain.close()})   
#$MainSplitContainer = new-object System.Windows.Forms.SplitContainer
$MainTabControl = New-object System.Windows.Forms.TabControl
$SCCMPage = New-Object System.Windows.Forms.TabPage
$SQLHealthPage = New-Object System.Windows.Forms.TabPage
$WMIExplorer = New-Object System.Windows.Forms.TabPage

#region WMIEXPLORER Controls
$SplitContainer1 = new-object System.Windows.Forms.SplitContainer  
$splitContainer2 = new-object System.Windows.Forms.SplitContainer  
$splitContainer3 = new-object System.Windows.Forms.SplitContainer  
$grpComputer = new-object System.Windows.Forms.GroupBox  
$grpNameSpaces = new-object System.Windows.Forms.GroupBox  
$grpClasses = new-object System.Windows.Forms.GroupBox  
$grpClass = new-object System.Windows.Forms.GroupBox  
$grpInstances = new-object System.Windows.Forms.GroupBox  
$grpStatus = new-object System.Windows.Forms.GroupBox  
$txtComputer = new-object System.Windows.Forms.TextBox  
$btnConnect = new-object System.Windows.Forms.Button  
$btnInstances = new-object System.Windows.Forms.Button  
$tvNameSpaces = new-object System.Windows.Forms.TreeView  
$lvClasses = new-object System.Windows.Forms.ListView  
$clbProperties = new-object System.Windows.Forms.CheckedListBox  
$clbProperties.CheckOnClick = $true  
$lbMethods = new-object System.Windows.Forms.ListBox  
$label1 = new-object System.Windows.Forms.Label  
$label2 = new-object System.Windows.Forms.Label  
$lblServer = new-object System.Windows.Forms.Label  
$lblPath = new-object System.Windows.Forms.Label  
$lblNameSpace = new-object System.Windows.Forms.Label  
$label6 = new-object System.Windows.Forms.Label  
$lblClass = new-object System.Windows.Forms.Label  
$label10 = new-object System.Windows.Forms.Label  
$lblClasses = new-object System.Windows.Forms.Label  
$label12 = new-object System.Windows.Forms.Label  
$lblProperties = new-object System.Windows.Forms.Label  
$label8 = new-object System.Windows.Forms.Label  
$lblMethods = new-object System.Windows.Forms.Label  
$label14 = new-object System.Windows.Forms.Label  
$lblInstances = new-object System.Windows.Forms.Label  
$label16 = new-object System.Windows.Forms.Label  
$dgInstances = new-object System.Windows.Forms.DataGridView  
$TabControl = new-object System.Windows.Forms.TabControl  
$tabPage1 = new-object System.Windows.Forms.TabPage  
$tabInstances = new-object System.Windows.Forms.TabPage  
$rtbHelp = new-object System.Windows.Forms.RichTextBox  
$tabMethods = new-object System.Windows.Forms.TabPage  
$rtbMethods = new-object System.Windows.Forms.RichTextBox  
#endregion WMIEXPLORER Controls

#endregion Define Used Controls         
#region Suspend the Layout  
#$MainSplitContainer.Panel1.SuspendLayout()
$MainTabControl.SuspendLayout()
$SCCMPage.SuspendLayout()
$SQLHealthPage.SuspendLayout()
$WMIExplorer.SuspendLayout()

#region WMIEXPLORER Controls
$splitContainer1.Panel1.SuspendLayout()  
$splitContainer1.Panel2.SuspendLayout()  
$splitContainer1.SuspendLayout()  
$splitContainer2.Panel1.SuspendLayout()  
$splitContainer2.Panel2.SuspendLayout()  
$splitContainer2.SuspendLayout()  
$grpComputer.SuspendLayout()  
$grpNameSpaces.SuspendLayout()  
$grpClasses.SuspendLayout()  
$splitContainer3.Panel1.SuspendLayout()  
$splitContainer3.Panel2.SuspendLayout()  
$splitContainer3.SuspendLayout()  
$grpClass.SuspendLayout()  
$grpStatus.SuspendLayout()  
$grpInstances.SuspendLayout()  
$TabControl.SuspendLayout()  
$tabPage1.SuspendLayout()  
$tabInstances.SuspendLayout()  
#endregion WMIEXPLORER Controls

$FrmMain.SuspendLayout()  
#endregion Suspend the Layout  
#region Configure Controls  
[void]$MainMenu.Items.Add($FileMenu)  
[void]$MainMenu.Items.Add($ToolMenu)  
$MainMenu.Location = new-object System.Drawing.Point(0, 0)  
$MainMenu.Name = "MainMenu"  
$MainMenu.Size = new-object System.Drawing.Size(1151, 24)  
$MainMenu.TabIndex = 0  
$MainMenu.Text = "Main Menu"  
#  
# statusStrip1  
#  
$statusStrip.Location = new-object System.Drawing.Point(0, 569)  
$statusStrip.Name = "statusStrip"  
$statusStrip.Size = new-object System.Drawing.Size(1151, 22);  
$statusStrip.TabIndex = 1  
$statusStrip.Text = "statusStrip"   


#  
# fileMenu  
#  
[void]$fileMenu.DropDownItems.Add($miQuit)  
$fileMenu.Name = "fileMenu"  
$fileMenu.Size = new-object System.Drawing.Size(35, 20)  
$fileMenu.Text = "&File"   

# TabControl  
#  ### Main tab control
$MainTabControl.Controls.Add($SCCMPage)
$MainTabControl.Controls.Add($SQLHealthPage)
$MainTabControl.Controls.Add($WMIExplorer)
$MainTabControl.Dock = [System.Windows.Forms.DockStyle]::Fill  
$MainTabControl.Location = new-object System.Drawing.Point(0, 0)  
$MainTabControl.Name = "TabControl"  
$MainTabControl.SelectedIndex = 0  
#$MainTabControl.Size = new-object System.Drawing.Size(771, 234)  
$MainTabControl.TabIndex = 0 
$MainTabControl.AutoSize = $true
#  
# SCCMPage  
#  
#$SCCMPage.Controls.Add($rtbHelp)  
$SCCMPage.Location = new-object System.Drawing.Point(4, 22)  
$SCCMPage.Name = "SCCMConsole"  
$SCCMPage.Padding = new-object System.Windows.Forms.Padding(3)  
$SCCMPage.Size = new-object System.Drawing.Size(763, 208)  
$SCCMPage.TabIndex = 0  
$SCCMPage.Text = "SCCM Console"  
$SCCMPage.UseVisualStyleBackColor = $true  
#  
# SQLHealthPage  
#  
#$SQLHealthPage.Controls.Add($grpInstances)  
$SQLHealthPage.Location = new-object System.Drawing.Point(4, 22)  
$SQLHealthPage.Name = "SQLHealthPage"  
$SQLHealthPage.Padding = new-object System.Windows.Forms.Padding(3)  
$SQLHealthPage.Size = new-object System.Drawing.Size(763, 208)  
$SQLHealthPage.TabIndex = 1  
$SQLHealthPage.Text = "SQL Health Check"  
$SQLHealthPage.UseVisualStyleBackColor = $true  
#  
# WMIExplorer  
#  
$WMIExplorer.Location = new-object System.Drawing.Point(4, 22)  
$WMIExplorer.Name = "WMIExplorer"  
$WMIExplorer.Padding = new-object System.Windows.Forms.Padding(3)  
$WMIExplorer.Size = new-object System.Drawing.Size(763, 208)  
$WMIExplorer.TabIndex = 2  
$WMIExplorer.Text = "WMIExplorer"  
$WMIExplorer.UseVisualStyleBackColor = $true  
$WMIExplorer.add_Click({$frmMain.Controls.Item("ToolMenu").Enabled = $true})
$WMIExplorer.Controls.Add($splitContainer1)

#region WMIEXPLORER Controls
$splitContainer1.Dock = [System.Windows.Forms.DockStyle]::Fill  
$splitContainer1.Location = new-object System.Drawing.Point(0, 24)  
$splitContainer1.Name = "splitContainer1"  
$splitContainer1.Panel1.Controls.Add($splitContainer2)  
$splitContainer1.Panel2.Controls.Add($splitContainer3)  
$splitContainer1.Size = new-object System.Drawing.Size(1151, 545)  
$splitContainer1.SplitterDistance = 372  
$splitContainer1.TabIndex = 2  
$splitContainer2.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$splitContainer2.Dock = [System.Windows.Forms.DockStyle]::Fill  
$splitContainer2.Location = new-object System.Drawing.Point(0, 0)  
$splitContainer2.Name = "splitContainer2"  
$splitContainer2.Orientation = [System.Windows.Forms.Orientation]::Horizontal  
$splitContainer2.Panel1.BackColor = [System.Drawing.SystemColors]::Control  
$splitContainer2.Panel1.Controls.Add($grpNameSpaces)  
$splitContainer2.Panel1.Controls.Add($btnConnect)  
$splitContainer2.Panel1.Controls.Add($grpComputer)  
$splitContainer2.Panel2.Controls.Add($grpClasses)  
$splitContainer2.Size = new-object System.Drawing.Size(372, 545)  
$splitContainer2.SplitterDistance = 302  
$splitContainer2.TabIndex = 0

$grpComputer.Anchor = "top, left, right"  
$grpComputer.Controls.Add($txtComputer)  
$grpComputer.Location = new-object System.Drawing.Point(12, 3)  
$grpComputer.Name = "grpComputer"  
$grpComputer.Size = new-object System.Drawing.Size(340, 57)  
$grpComputer.TabIndex = 0  
$grpComputer.TabStop = $false  
$grpComputer.Text = "Computer"  
$txtComputer.Anchor = "top, left, right"  
$txtComputer.Location = new-object System.Drawing.Point(7, 20)  
$txtComputer.Name = "txtComputer"  
$txtComputer.Size = new-object System.Drawing.Size(244, 20)  
$txtComputer.TabIndex = 0  
$txtComputer.Text = "."  
 
$btnConnect.Anchor = "top, right"  
$btnConnect.Location = new-object System.Drawing.Point(269, 23);  
$btnConnect.Name = "btnConnect"  
$btnConnect.Size = new-object System.Drawing.Size(75, 23)  
$btnConnect.TabIndex = 1  
$btnConnect.Text = "Connect"  
$btnConnect.UseVisualStyleBackColor = $true  
#  
# grpNameSpaces  
#  
$grpNameSpaces.Anchor = "Bottom, top, left, right"  
$grpNameSpaces.Controls.Add($tvNameSpaces)  
$grpNameSpaces.Location = new-object System.Drawing.Point(12, 67)  
$grpNameSpaces.Name = "grpNameSpaces"  
$grpNameSpaces.Size = new-object System.Drawing.Size(340, 217)  
$grpNameSpaces.TabIndex = 2  
$grpNameSpaces.TabStop = $false  
$grpNameSpaces.Text = "NameSpaces"  
#  
# grpClasses  
#  
$grpClasses.Anchor = "Bottom, top, left, right"  
$grpClasses.Controls.Add($lvClasses)  
$grpClasses.Location = new-object System.Drawing.Point(12, 14)  
$grpClasses.Name = "grpClasses"  
$grpClasses.Size = new-object System.Drawing.Size(340, 206)  
$grpClasses.TabIndex = 0  
$grpClasses.TabStop = $False  
$grpClasses.Text = "Classes"  
#  
# tvNameSpaces  
#  
$tvNameSpaces.Anchor = "Bottom, top, left, right"  
$tvNameSpaces.Location = new-object System.Drawing.Point(7, 19)  
$tvNameSpaces.Name = "tvNameSpaces"  
$tvNameSpaces.Size = new-object System.Drawing.Size(325, 184)  
$tvNameSpaces.TabIndex = 0  
#  
# tvClasses  
#  
$lvClasses.Anchor = "Bottom, top, left, right"  
$lvClasses.Location = new-object System.Drawing.Point(7, 19)  
$lvClasses.Name = "tvClasses"  
$lvClasses.Size = new-object System.Drawing.Size(325, 172)  
$lvClasses.TabIndex = 0  
$lvClasses.UseCompatibleStateImageBehavior = $False  
$lvClasses.ShowItemToolTips = $true  
$lvClasses.View = 'Details'  
$colName = $lvClasses.Columns.add('Name')  
$colname.Width = 160  
$colPath = $lvClasses.Columns.add('Description')  
$colname.Width = 260  
$colPath = $lvClasses.Columns.add('Path')  
$colname.Width = 260  
#  
# splitContainer3  
#  
$splitContainer3.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$splitContainer3.Dock = [System.Windows.Forms.DockStyle]::Fill  
$splitContainer3.Location = new-object System.Drawing.Point(0, 0)  
$splitContainer3.Name = "splitContainer3"  
$splitContainer3.Orientation = [System.Windows.Forms.Orientation]::Horizontal  
#  
# splitContainer3.Panel1  
#  
$splitContainer3.Panel1.Controls.Add($grpStatus)  
$splitContainer3.Panel1.Controls.Add($grpClass)  
#  
# splitContainer3.Panel2  
#  
$splitContainer3.Panel2.Controls.Add($TabControl)  
$splitContainer3.Size = new-object System.Drawing.Size(775, 545)  
$splitContainer3.SplitterDistance = 303  
$splitContainer3.TabIndex = 0  
#  
# grpClass  
#  
$grpClass.Anchor = "Bottom, top, left, right"  
$grpClass.Controls.Add($lblInstances)  
$grpClass.Controls.Add($label16)  
$grpClass.Controls.Add($lblMethods)  
$grpClass.Controls.Add($label14)  
$grpClass.Controls.Add($lblProperties)  
$grpClass.Controls.Add($label8)  
$grpClass.Controls.Add($lblClass)  
$grpClass.Controls.Add($label10)  
$grpClass.Controls.Add($lbMethods)  
$grpClass.Controls.Add($clbProperties)  
$grpClass.Controls.Add($btnInstances)  
$grpClass.Location = new-object System.Drawing.Point(17, 86)  
$grpClass.Name = "grpClass"  
$grpClass.Size = new-object System.Drawing.Size(744, 198)  
$grpClass.TabIndex = 0  
$grpClass.TabStop = $False  
$grpClass.Text = "Class"  
#  
# btnInstances  
#  
$btnInstances.Anchor = "Bottom, Left"  
$btnInstances.Location = new-object System.Drawing.Point(6, 169);  
$btnInstances.Name = "btnInstances";  
$btnInstances.Size = new-object System.Drawing.Size(96, 23);  
$btnInstances.TabIndex = 0;  
$btnInstances.Text = "Get Instances";  
$btnInstances.UseVisualStyleBackColor = $true  
#  
# grpStatus  
#  
$grpStatus.Anchor = "Top,Left,Right"  
$grpStatus.Controls.Add($lblClasses)  
$grpStatus.Controls.Add($label12)  
$grpStatus.Controls.Add($lblNameSpace)  
$grpStatus.Controls.Add($label6)  
$grpStatus.Controls.Add($lblPath)  
$grpStatus.Controls.Add($lblServer)  
$grpStatus.Controls.Add($label2)  
$grpStatus.Controls.Add($label1)  
$grpStatus.Location = new-object System.Drawing.Point(17, 3)  
$grpStatus.Name = "grpStatus"  
$grpStatus.Size = new-object System.Drawing.Size(744, 77)  
$grpStatus.TabIndex = 1  
$grpStatus.TabStop = $False  
$grpStatus.Text = "Status"  
#  
# label1  
#  
$label1.AutoSize = $true  
$label1.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label1.Location = new-object System.Drawing.Point(7, 20)  
$label1.Name = "label1"  
$label1.Size = new-object System.Drawing.Size(62, 16)  
$label1.TabIndex = 0  
$label1.Text = "Server :"  
#  
# label2  
#  
$label2.AutoSize = $true  
$label2.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label2.Location = new-object System.Drawing.Point(7, 41)  
$label2.Name = "label2"  
$label2.Size = new-object System.Drawing.Size(51, 16)  
$label2.TabIndex = 1  
$label2.Text = "Path  :"  
#  
# lblServer  
#  
$lblServer.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblServer.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblServer.Location = new-object System.Drawing.Point(75, 20)  
$lblServer.Name = "lblServer"  
$lblServer.Size = new-object System.Drawing.Size(144, 20)  
$lblServer.TabIndex = 2  
#  
# lblPath  
#  
$lblPath.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblPath.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblPath.Location = new-object System.Drawing.Point(75, 40)  
$lblPath.Name = "lblPath"  
$lblPath.Size = new-object System.Drawing.Size(567, 20)  
$lblPath.TabIndex = 3  
#  
# lblNameSpace  
#  
$lblNameSpace.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblNameSpace.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblNameSpace.Location = new-object System.Drawing.Point(337, 20)  
$lblNameSpace.Name = "lblNameSpace"  
$lblNameSpace.Size = new-object System.Drawing.Size(144, 20)  
$lblNameSpace.TabIndex = 5  
#  
# label6  
#  
$label6.AutoSize = $true  
$label6.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label6.Location = new-object System.Drawing.Point(229, 20)  
$label6.Name = "label6"  
$label6.Size = new-object System.Drawing.Size(102, 16)  
$label6.TabIndex = 4  
$label6.Text = "NameSpace :"  
#  
# lblClass  
#  
$lblClass.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblClass.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblClass.Location = new-object System.Drawing.Point(110, 26)  
$lblClass.Name = "lblClass"  
$lblClass.Size = new-object System.Drawing.Size(159, 20)  
$lblClass.TabIndex = 11  
#  
# label10  
#  
$label10.AutoSize = $true  
$label10.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label10.Location = new-object System.Drawing.Point(6, 26)  
$label10.Name = "label10"  
$label10.Size = new-object System.Drawing.Size(55, 16)  
$label10.TabIndex = 10  
$label10.Text = "Class :"  
#  
# lblClasses  
#  
$lblClasses.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblClasses.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblClasses.Location = new-object System.Drawing.Point(595, 21)  
$lblClasses.Name = "lblClasses"  
$lblClasses.Size = new-object System.Drawing.Size(47, 20)  
$lblClasses.TabIndex = 9  
#  
# label12  
#  
$label12.AutoSize = $true  
$label12.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label12.Location = new-object System.Drawing.Point(487, 21)  
$label12.Name = "label12"  
$label12.Size = new-object System.Drawing.Size(76, 16)  
$label12.TabIndex = 8  
$label12.Text = "Classes  :"  
#  
# clbProperties  
#  
$clbProperties.Anchor = "Bottom, top,left"  
$clbProperties.FormattingEnabled = $true  
$clbProperties.Location = new-object System.Drawing.Point(510, 27)  
$clbProperties.Name = "clbProperties"  
$clbProperties.Size = new-object System.Drawing.Size(220, 160)  
$clbProperties.TabIndex = 1  
#  
# lbMethods  
#  
$lbMethods.Anchor = "Bottom, top, Left"  
$lbMethods.FormattingEnabled = $true  
$lbMethods.Location = new-object System.Drawing.Point(280, 27)  
$lbMethods.Name = "lbMethods"  
$lbMethods.Size = new-object System.Drawing.Size(220, 160)  
$lbMethods.TabIndex = 2  
#  
# lblProperties  
#  
$lblProperties.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblProperties.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblProperties.Location = new-object System.Drawing.Point(110, 46)  
$lblProperties.Name = "lblProperties"  
$lblProperties.Size = new-object System.Drawing.Size(119, 20)  
$lblProperties.TabIndex = 13  
#  
# label8  
#  
$label8.AutoSize = $true  
$label8.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label8.Location = new-object System.Drawing.Point(6, 46)  
$label8.Name = "label8"  
$label8.Size = new-object System.Drawing.Size(88, 16)  
$label8.TabIndex = 12  
$label8.Text = "Properties :"  
#  
# lblMethods  
#  
$lblMethods.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblMethods.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblMethods.Location = new-object System.Drawing.Point(110, 66)  
$lblMethods.Name = "lblMethods"  
$lblMethods.Size = new-object System.Drawing.Size(119, 20)  
$lblMethods.TabIndex = 15  
#  
# label14  
#  
$label14.AutoSize = $true  
$label14.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label14.Location = new-object System.Drawing.Point(6, 66)  
$label14.Name = "label14"  
$label14.Size = new-object System.Drawing.Size(79, 16)  
$label14.TabIndex = 14  
$label14.Text = "Methods  :"  
#  
# lblInstances  
#  
$lblInstances.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D  
$lblInstances.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$lblInstances.Location = new-object System.Drawing.Point(110, 86)  
$lblInstances.Name = "lblInstances"  
$lblInstances.Size = new-object System.Drawing.Size(119, 20)  
$lblInstances.TabIndex = 17  
#  
# label16  
#  
$label16.AutoSize = $true  
$label16.Font = new-object System.Drawing.Font("Microsoft Sans Serif",9.75 ,[System.Drawing.FontStyle]::Bold)  
$label16.Location = new-object System.Drawing.Point(6, 86)  
$label16.Name = "label16"  
$label16.Size = new-object System.Drawing.Size(82, 16)  
$label16.TabIndex = 16  
$label16.Text = "Instances :"  
#  
# grpInstances  
#  
$grpInstances.Anchor = "Bottom, top, left, right"  
$grpInstances.Controls.Add($dgInstances)  
$grpInstances.Location = new-object System.Drawing.Point(17, 17)  
$grpInstances.Name = "grpInstances"  
$grpInstances.Size = new-object System.Drawing.Size(744, 202)  
$grpInstances.TabIndex = 0  
$grpInstances.TabStop = $False  
$grpInstances.Text = "Instances"  
#  
# dgInstances  
#  
$dgInstances.Anchor = "Bottom, top, left, right"  
$dgInstances.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize  
$dgInstances.Location = new-object System.Drawing.Point(10, 19)  
$dgInstances.Name = "dgInstances"  
$dgInstances.Size = new-object System.Drawing.Size(728, 167)  
$dgInstances.TabIndex = 0  
$dginstances.ReadOnly = $true  
# TabControl  
#  
$TabControl.Controls.Add($tabPage1)  
$TabControl.Controls.Add($tabInstances)  
$TabControl.Controls.Add($tabMethods)  
$TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill  
$TabControl.Location = new-object System.Drawing.Point(0, 0)  
$TabControl.Name = "TabControl"  
$TabControl.SelectedIndex = 0  
$TabControl.Size = new-object System.Drawing.Size(771, 234)  
$TabControl.TabIndex = 0  
#  
# tabPage1  
#  
$tabPage1.Controls.Add($rtbHelp)  
$tabPage1.Location = new-object System.Drawing.Point(4, 22)  
$tabPage1.Name = "tabPage1"  
$tabPage1.Padding = new-object System.Windows.Forms.Padding(3)  
$tabPage1.Size = new-object System.Drawing.Size(763, 208)  
$tabPage1.TabIndex = 0  
$tabPage1.Text = "Help"  
$tabPage1.UseVisualStyleBackColor = $true  
#  
# tabInstances  
#  
$tabInstances.Controls.Add($grpInstances)  
$tabInstances.Location = new-object System.Drawing.Point(4, 22)  
$tabInstances.Name = "tabInstances"  
$tabInstances.Padding = new-object System.Windows.Forms.Padding(3)  
$tabInstances.Size = new-object System.Drawing.Size(763, 208)  
$tabInstances.TabIndex = 1  
$tabInstances.Text = "Instances"  
$tabInstances.UseVisualStyleBackColor = $true  
#  
# richTextBox1  
#  
$rtbHelp.Dock = [System.Windows.Forms.DockStyle]::Fill  
$rtbHelp.Location = new-object System.Drawing.Point(3, 3)  
$rtbHelp.Name = "richTextBox1"  
$rtbHelp.Size = new-object System.Drawing.Size(757, 202)  
$rtbHelp.TabIndex = 0  
$rtbHelp.Text = ""  
#  
# tabMethods  
#  
$tabMethods.Location = new-object System.Drawing.Point(4, 22)  
$tabMethods.Name = "tabMethods"  
$tabMethods.Padding = new-object System.Windows.Forms.Padding(3)  
$tabMethods.Size = new-object System.Drawing.Size(763, 208)  
$tabMethods.TabIndex = 2  
$tabMethods.Text = "Methods"  
$tabMethods.UseVisualStyleBackColor = $true  
 
$rtbMethods.Dock = [System.Windows.Forms.DockStyle]::Fill  
$rtbMethods.Font = new-object System.Drawing.Font("Lucida Console",8 )  
$rtbMethods.DetectUrls = $false  
$tabMethods.controls.add($rtbMethods) 
#endregion WMIEXPLORER Controls

#endregion Configure Controls  
# Configure  Main Form  
#region frmMain  
 
#  
$frmMain.AutoScaleDimensions = new-object System.Drawing.SizeF(6, 13)  
$frmMain.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font  
$frmMain.ClientSize = new-object System.Drawing.Size(1151, 591)  
#$frmMain.Controls.Add($MainsplitContainer) 
$frmMain.Controls.Add($MainTabControl)
$frmMain.Controls.Add($statusStrip)  
$frmMain.Controls.Add($MainMenu)  
$frmMain.MainMenuStrip = $mainMenu  
$FrmMain.Name = "frmMain"  
$FrmMain.Text = "All In All Azhguraja"

$MainSplitContainer.ResumeLayout($false)
$MainTabControl.ResumeLayout($false)
$SCCMPage.ResumeLayout($false)
$SQLHealthPage.ResumeLayout($false)
$WMIExplorer.ResumeLayout($false)

$splitContainer1.Panel1.ResumeLayout($false)  
$splitContainer1.Panel2.ResumeLayout($false)  
$splitContainer1.ResumeLayout($false)  
$splitContainer2.Panel1.ResumeLayout($false)  
$splitContainer2.Panel2.ResumeLayout($false)  
$splitContainer2.ResumeLayout($false)  
$grpComputer.ResumeLayout($false)  
$grpComputer.PerformLayout()  
$grpNameSpaces.ResumeLayout($false)  
$grpClasses.ResumeLayout($false)  
$splitContainer3.Panel1.ResumeLayout($false)  
$splitContainer3.Panel2.ResumeLayout($false)  
$splitContainer3.ResumeLayout($false)  
$grpClass.ResumeLayout($false)  
$grpClass.PerformLayout()  
$grpStatus.ResumeLayout($false)  
$grpStatus.PerformLayout()  
$grpInstances.ResumeLayout($false)  
$TabControl.ResumeLayout($false)  
$tabPage1.ResumeLayout($false)  
$tabInstances.ResumeLayout($false)  


$mainMenu.ResumeLayout($false)  
$mainMenu.PerformLayout()  
$MainMenu.ResumeLayout($false)  
$MainMenu.PerformLayout()  

$frmMain.ResumeLayout($false)  
$FrmMain.PerformLayout()  
$status = new-object System.Windows.Forms.ToolStripStatusLabel  
$status.BorderStyle = 'SunkenInner'  
$status.BorderSides = 'All'  
$status.Text = "Not Connected"  
[void]$statusStrip.Items.add($status)  
$slMessage = new-object System.Windows.Forms.ToolStripStatusLabel  
$slMessage.BorderStyle = 'SunkenInner'  
$slMessage.BorderSides = 'All'  
$slMessage.Text = ""  
[void]$statusStrip.Items.add($slMessage)  
#endregion frmMain  

#region Helper Functions  
Function out-PropertyGrid {  
  Param ($Object,[switch]$noBase,[Switch]$array)  
  $PsObject = $null  
  if ($object) {  
      $PsObject = $object  
  }Else{  
     if ($Array.IsPresent) {  
         $PsObject = @()  
         $input |ForEach-Object {$PsObject += $_}  
     }Else{  
         $input |ForEach-Object {$PsObject = $_}  
     }  
  }  
  if ($PsObject){  
      $form = new-object Windows.Forms.Form   
      $form.Size = new-object Drawing.Size @(600,600)   
      $PG = new-object Windows.Forms.PropertyGrid   
      $PG.Dock = 'Fill'   
      $form.text = "$psObject"   
      if ($noBase.IsPresent) {"no";  
          $PG.selectedobject = $psObject   
      }Else{  
          $PG.selectedobject = $psObject.PsObject.baseobject   
      }   
      $form.Controls.Add($PG)   
      $Form.Add_Shown({$form.Activate()})    
      $form.showdialog()  
  }  
} #Function out-PropertyGrid  
Function Update-Status {  
  $script:computer = $Script:NameSpaces.__SERVER  
  $txtComputer.Text = $script:computer  
  $lblPath.Text = $Script:NameSpaces.__PATH                                 
  $lblProperties.Text = $Script:NameSpaces.__PROPERTY_COUNT                                 
  $lblClass.Text = $Script:NameSpaces.__RELPATH                                     
  $lblServer.Text = $script:Computer  
  $lblnamespace.Text = $Script:NameSpaces.__NAMESPACE  
} # Function Update-Status  
Function Set-StatusBar ([Drawing.Color]$Color,$Text) {  
  $status.BackColor = $color  
  $status.Text = $text  
  $statusstrip.Update()    
}  
#endregion Helper Functions  
#################### Main ###############################  
#region Global Variables  
$FontBold = new-object System.Drawing.Font("Microsoft Sans Serif",8,[Drawing.FontStyle]'Bold' )  
$fontNormal = new-object System.Drawing.Font("Microsoft Sans Serif",8,[Drawing.FontStyle]'Regular')  
$fontCode = new-object System.Drawing.Font("Lucida Console",8 )  
# Create Script Variables for WMI Connection  
$Script:ConnectionOptions = new-object System.Management.ConnectionOptions  
$script:WmiConnection = new-object system.management.ManagementScope  
$script:WmiClass = [wmiClass]''  
# NamespaceCaching , Make HashTable to store Treeview Items  
$script:nsc = @{}  
# Make DataSet for secondary Cache  
$Script:dsCache = new-object data.dataset  
if (-not ${Global:WmiExplorer.dtClasses}){  
    ${Global:WmiExplorer.dtClasses} = new-object data.datatable  
    [VOID](${Global:WmiExplorer.dtClasses}.Columns.add('Path',[string]))  
    [VOID](${Global:WmiExplorer.dtClasses}.Columns.add('Namespace',[string]))  
    [VOID](${Global:WmiExplorer.dtClasses}.Columns.add('name',[string]))  
    [VOID](${Global:WmiExplorer.dtClasses}.Columns.add('Description',[string]))  
    ${Global:WmiExplorer.dtClasses}.tablename = 'Classes'  
}  
#endregion  
#region Control Handlers  
# Add Delegate Scripts to finetune the WMI Connection objects to the events of the controls  
$slMessage.DoubleClickEnabled = $true  
$slMessage.add_DoubleClick({$error[0] | out-PropertyGrid})  
$lblNameSpace.add_DoubleClick({$script:WmiConnection | out-PropertyGrid})  
$lblserver.add_DoubleClick({$Script:ConnectionOptions | out-PropertyGrid})  
$lblClass.add_DoubleClick({$script:WmiClass | out-PropertyGrid})  
 
$btnConnect.add_click({ConnectToComputer})  
$TVNameSpaces.add_DoubleClick({GetClassesFromNameSpace})  
$lvClasses.Add_DoubleClick({GetWmiClass})  
$btnInstances.add_Click({GetWmiInstances})  
$dgInstances.add_DoubleClick({OutputWmiInstance})  
$lbMethods.Add_DoubleClick({GetWmiMethod})  
$clbProperties.add_Click({  
  trap{Continue}  
  $DGInstances.Columns.Item(($this.SelectedItem)).visible = -not $clbProperties.GetItemChecked($this.SelectedIndex)  
})  
$TVNameSpaces.add_AfterSelect({  
    if ($this.SelectedNode.name -ne $Computer){  
        $lblPath.Text = "$($script:WmiConnection.path.path.replace('\root',''))\$($this.SelectedNode.Text)"   
    }  
   
    $lblProperties.Text = $Script:NameSpaces.__PROPERTY_COUNT                                 
    $lblServer.Text = $Script:NameSpaces.__SERVER  
    $lblnamespace.Text = $this.SelectedNode.Text  
    if ($this.SelectedNode.tag -eq "NotEnumerated") {  
        (new-object system.management.managementClass(  
                "$($script:WmiConnection.path.path.replace('\root',''))\$($this.SelectedNode.Text):__NAMESPACE")  
        ).PSbase.getInstances() | Sort-Object $_.name |  
        ForEach-Object {  
          $TN = new-object System.Windows.Forms.TreeNode  
          $TN.Name = $_.name  
          $TN.Text = ("{0}\{1}" -f $_.__NameSpace,$_.name)  
          $TN.tag = "NotEnumerated"  
          $this.SelectedNode.Nodes.Add($TN)  
        }  
         
        # Set tag to show this node is already enumerated  
        $this.SelectedNode.tag = "Enumerated"  
    }  
    $mp = ("{0}\{1}" -f $script:WmiConnection.path.path.replace('\root','') , $this.SelectedNode.text)  
    $lvClasses.Items.Clear()  
    if($Script:nsc.Item("$mp")){ # in Namespace cache  
        $lvClasses.BeginUpdate()  
        $lvClasses.Items.AddRange(($nsc.Item( "$mp")))  
        $status.Text = "$mp : $($lvClasses.Items.count) Classes"  
        $lvClasses.EndUpdate()  
        $lblClasses.Text = $lvClasses.Items.count  
    } else {  
        if(${Global:WmiExplorer.dtClasses}.Select("Namespace='$mp'")){ # In DataTable Cache  
            $status.BackColor = 'beige'  
            $status.Text = "$mp : Classes in Cache, DoubleClick NameSpace to retrieve Classes"  
        } else {  
            $status.BackColor = 'LightSalmon'  
            $status.Text = "$mp : Classes not recieved yet, DoubleClick NameSpace to retrieve Classes"  
        }  
    }  
}) # $TVNameSpaces.add_AfterSelect  
#endregion  
#region Processing Functions  
#region ConnectToComputer  
# Connect to Computer  
Function ConnectToComputer {  
     
    $computer = $txtComputer.Text  
    Set-StatusBar 'beige' "Connecting to : $computer"  
     
    # Try to Connect to Computer  
    &{  
        trap {  
            Set-StatusBar 'Red' "Connecting to : $computer Failed"  
            $slMessage.Text = "$_.message"  
            Continue  
        }  
        &{  
            # Connect to WMI root  
             
            $script:WmiConnection.path = "\\$computer\root"  
            $script:WmiConnection.options = $Script:ConnectionOptions  
            $script:WmiConnection.Connect()  
             
            # Get Avaiable NameSpaces  
     
            $opt = new-object system.management.ObjectGetOptions  
            $opt.UseAmendedQualifiers = $true  
            $Script:NameSpaces = new-object System.Management.ManagementClass(  
                $script:WmiConnection,[Management.ManagementPath]'__Namespace',$opt  
            )  
            Update-Status  
            # Create a TreeNode for the WMI Root found  
            $computer = $txtComputer.Text  
            $TNRoot = new-object System.Windows.Forms.TreeNode("Root")  
            $TNRoot.Name = $Computer  
            $TNRoot.Text = $lblPath.Text  
            $TNRoot.tag = "Enumerated"  
             
            # Create NameSpaces List  
             
            $Script:NameSpaces.PSbase.getInstances() | Sort-Object $_.name |  
            ForEach-Object {  
                $TN = new-object System.Windows.Forms.TreeNode  
                $TN.Name = $_.name  
                $TN.Text = ("{0}\{1}" -f $_.__NameSpace,$_.name)  
                $TN.tag = "NotEnumerated"  
                [void]$TNRoot.Nodes.Add($TN)  
            }  
            # Add to Treeview  
            $tvNameSpaces.Nodes.clear()  
            [void]$TVNamespaces.Nodes.Add($TNRoot)  
             
            # update StatusBar  
            Set-StatusBar 'YellowGreen' "Connected to : $computer"  
        }  
    }  
} # ConnectToComputer  
#endregion  
#region GetClasseFromNameSpace  
# Get Classes on DoubleClick on Namespace in TreeView  
Function GetClassesFromNameSpace {  
  if ($this.SelectedNode.name -ne $script:computer){  
    # Connect to WMI Namespace  
         
    $mp = ("{0}\{1}" -f $script:WmiConnection.path.path.replace('\root','') , $this.SelectedNode.text)  
      # Update Status  
         
      $lvClasses.BeginUpdate()  
      $lvClasses.Items.Clear()  
      $i = 0 ;$lblClasses.Text = $i; $lblclasses.Update()  
    if($Script:nsc.Item("$mp")){ #in Namespace Cache, so just attach to ListView again  
         
        $lvClasses.Items.AddRange(($nsc.Item( "$mp")))  
        # $lvClasses.Items.AddRange(([System.Windows.Forms.ListViewItem[]]($nsc.Item( "$mp") |  
            # where {$_.name -like 'win32_*'})))  
        $status.Text = "$mp : $($lvClasses.Items.count) Classes"  
        $i = $lvClasses.Items.count  
    } else { #Not In NameSpace Cache  
      if(${Global:WmiExplorer.dtClasses}.Select("Namespace = '$mp'")){ # In DataTable cache, so get from there  
        $status.Text = "loading cache from $($this.SelectedNode.name)"  
        $statusStrip.Update()  
        ${Global:WmiExplorer.dtClasses}.Select("Namespace = '$mp'") |  
        foreach {  
            $i++  
            $LI = New-Object system.Windows.Forms.ListViewItem  
            $li.Name = $_.name  
            $li.Text = $_.name  
            $li.SubItems.add($_.description)  
            $li.SubItems.add($_.path)  
            $li.ToolTipText = ($_.description)  
            $lvClasses.Items.add($li)  
            $status.Text = "$mp : $($lvClasses.Items.count) Classes"  
            $lblClasses.Text = $lvClasses.Items.count  
        }  
      } else { # Not in any Cache , Load WMI Classes  
        Set-StatusBar 'Khaki' "Getting Classes from $($this.SelectedNode.name)"  
        $mc = new-object System.Management.ManagementClass($mp,$opt)  
        $eo = New-Object system.management.EnumerationOptions  
        $eo.EnumerateDeep = $true  
        $eo.UseAmendedQualifiers = $true  
        $Mc.psbase.GetSubclasses($eo) |  
        ForEach-Object  {  
            $i++ ; if ($i%10 -eq 0){$lblClasses.Text = $i;$lblclasses.Update() }  
            Trap{$script:Description = "[Empty]";continue}  
            $script:description = $_.psbase.Qualifiers.item("description").value  
            ${Global:WmiExplorer.dtClasses}.Rows.Add($_.__path,$mp,$_.name,$description)  
            $LI = New-Object system.Windows.Forms.ListViewItem  
            $li.Name = $_.name  
            $li.Text = $_.name  
            $li.SubItems.add($description)  
            $li.SubItems.add($_.__path)  
            $li.ToolTipText = $description  
            $lvClasses.Items.add($li)  
        }  
        $status.Text = "Ready, Retrieved $i Classes from $mp"  
      } #if(${Global:WmiExplorer.dtClasses}.Select("Namespace = '$mp'"))  
      $lvClasses.Sorting = 'Ascending'  
      $lvClasses.Sort()  
      $script:nsc.Add($mp,(([System.Windows.Forms.ListViewItem[]]($lvClasses.Items)).clone()))  
       
    }  
    $lvClasses.EndUpdate()  
    $this.selectedNode.BackColor = 'AliceBlue'  
    $lblClasses.Text = $i;$lblclasses.Update()  
    $status.BackColor = 'YellowGreen'  
    $statusStrip.Update()  
  } #if($Script:nsc.Item("$mp"))  
     
} # GetClassesFromNameSpace  
#endregion  
#region GetWmiClass  
Function GetWmiClass {  
    # Update Status  
     
    $status.Text = "Retrieving Class"  
    $status.BackColor = 'Khaki'  
    $statusstrip.Update()  
    $lblClass.Text =  $this.SelectedItems |ForEach-Object {$_.name}  
    $lblPath.text = $this.SelectedItems |ForEach-Object {"$($_.SubItems[2].text)"}  
     
    # Add HelpText  
     
    $rtbHelp.Text = ""  
    $rtbHelp.selectionFont  = $fontBold  
    $rtbHelp.appendtext("$($lblClass.Text)`n`n")  
    $rtbHelp.selectionFont  = $fontNormal  
    $rtbHelp.appendtext(($this.SelectedItems |ForEach-Object {"$($_.SubItems[1].text)"}))  
    $rtbHelp.appendtext("`n")  
    $path = $lblPath.text  
     
    $opt = new-object system.management.ObjectGetOptions  
    $opt.UseAmendedQualifiers = $true  
     
    $script:WmiClass = new-object system.management.ManagementClass($path,$opt)  
    # Add Property Help  
     
    $rtbHelp.selectionFont  = $fontBold  
    $rtbHelp.appendtext("`n$($lblClass.Text) Properties :`n`n")  
    $rtbHelp.selectionFont  = $fontNormal  
     
    $i = 0 ;$lblProperties.Text = $i; $lblProperties.Update()  
    $clbproperties.Items.Clear()  
    $clbProperties.Items.add('WmiPath',$False)  
             
    $script:WmiClass.psbase.properties |  
    ForEach-Object {  
        $i++ ;$lblProperties.Text = $i; $lblProperties.Update()  
        $clbProperties.Items.add($_.name,$true)  
        $rtbHelp.selectionFont  = $fontBold  
        $rtbHelp.appendtext("$($_.Name) :`n" )  
        &{  
            Trap {$rtbHelp.appendtext("[Empty]");Continue}  
            $rtbHelp.appendtext($_.psbase.Qualifiers["description"].value)  
        }  
        $rtbHelp.appendtext("`n`n")  
    } # ForEach-Object  
     
    # Create Method Help  
    $rtbHelp.selectionFont  = $fontBold  
    $rtbHelp.appendtext( "$($lblClass.Text) Methods :`n`n" )  
    $i = 0 ;$lblMethods.Text = $i; $lblMethods.Update()  
    $lbmethods.Items.Clear()  
     
    $script:WmiClass.psbase.Methods |  
    ForEach-Object {  
        $i++ ;$lblMethods.Text = $i; $lblMethods.Update()  
        $lbMethods.Items.add($_.name)  
        $rtbHelp.selectionFont  = $fontBold  
        $rtbHelp.appendtext("$($_.Name) :`n")  
        &{  
            Trap {$rtbHelp.Text += "[Empty]"}  
            $rtbHelp.appendtext($_.Qualifiers["description"].value)  
        }  
        $rtbHelp.appendtext("`n`n" )  
    } #ForEach-Object  
      
    $tabControl.SelectedTab = $tabpage1  
    $status.Text = "Retrieved Class"  
    $status.BackColor = 'YellowGreen'  
    $statusstrip.Update()  
} # GetWmiClass  
#endregion  
#region GetWmiInstances  
Function GetWmiInstances {  
    $status.Text = "Getting Instances for $($lblClass.text)"  
    $status.BackColor = 'Red'  
    $statusstrip.Update()  
    $tabControl.SelectedTab = $tabInstances  
    $MC = new-object system.management.ManagementClass $lblPath.text  
    $MOC = $MC.PSbase.getInstances()  
     
    #trap{"Class Not found";break}  
     
    $DT =  new-object  System.Data.DataTable  
    $DT.TableName = $lblClass.text  
    $Col =  new-object System.Data.DataColumn  
    $Col.ColumnName = "WmiPath"  
    $DT.Columns.Add($Col)  
    $i = 0  
    $j = 0 ;$lblInstances.Text = $j; $lblInstances.Update()  
    $MOC | ForEach-Object {  
        $j++ ;$lblInstances.Text = $j; $lblInstances.Update()  
        $MO = $_  
         
        # Make a DataRow  
        $DR = $DT.NewRow()  
        $Col =  new-object System.Data.DataColumn  
         
        $DR.Item("WmiPath") = $mo.__PATH  
        $MO.psbase.properties |  
        ForEach-Object {  
            $prop = $_  
            If ($i -eq 0)  {  
     
                # Only On First Row make The Headers  
                 
                $Col =  new-object System.Data.DataColumn  
                $Col.ColumnName = $prop.Name.ToString()  
                $prop.psbase.Qualifiers | ForEach-Object {  
                    If ($_.Name.ToLower() -eq "key") {  
                        $Col.ColumnName = $Col.ColumnName + "*"  
                    }  
                }  
                $DT.Columns.Add($Col)   
            }  
             
            # fill dataRow   
             
            if ($prop.value -eq $null) {  
                $DR.Item($prop.Name) = "[empty]"  
            }  
            ElseIf ($prop.IsArray) {  
                                $ofs = ";"  
                $DR.Item($prop.Name) ="$($prop.value)"  
                                $ofs = $null  
            }  
            Else {  
                $DR.Item($prop.Name) = $prop.value  
                #Item is Key try again with *  
                trap{$DR.Item("$($prop.Name)*") = $prop.Value.tostring();continue}  
            }  
        }  
        # Add the row to the DataTable  
        $DT.Rows.Add($DR)  
        $i += 1  
    }  
    $DGInstances.DataSource = $DT.psObject.baseobject  
        $DGInstances.Columns.Item('WmiPath').visible =  $clbProperties.GetItemChecked(0)   
    $status.Text = "Retrieved $j Instances"  
    $status.BackColor = 'YellowGreen'  
    $statusstrip.Update()  
} # GetWmiInstances  
#endregion  
#region OutputWmiInstance  
Function OutputWmiInstance {  
    if ( $this.SelectedRows.count -eq 1 ) {  
        if (-not $Script:InstanceTab) {$Script:InstanceTab = new-object System.Windows.Forms.TabPage  
            $Script:InstanceTab.Name = 'Instance'  
            $Script:rtbInstance = new-object System.Windows.Forms.RichTextBox  
            $Script:rtbInstance.Dock = [System.Windows.Forms.DockStyle]::Fill  
            $Script:rtbInstance.Font = $fontCode  
            $Script:rtbInstance.DetectUrls = $false  
            $Script:InstanceTab.controls.add($Script:rtbInstance)  
            $TabControl.TabPages.add($Script:InstanceTab)  
        }  
        $Script:InstanceTab.Text = "Instance = $($this.SelectedRows | ForEach-Object {$_.DataboundItem.wmiPath.split(':')[1]})" 
        $Script:rtbInstance.Text = $this.SelectedRows |ForEach-Object {$_.DataboundItem |Format-List  * | out-String -width 1000 } 
        $tabControl.SelectedTab = $Script:InstanceTab  
    }  
}  # OutputWmiInstance  
#endregion  
#region GetWmiMethod  
Function GetWmiMethod {  
    $WMIMethod = $this.SelectedItem  
    $WmiClassName = $script:WmiClass.__Class  
    $tabControl.SelectedTab = $tabMethods  
    #$rtbmethods.ForeColor = 'Green'  
    $rtbMethods.Font  = new-object System.Drawing.Font("Microsoft Sans Serif",8)  
    $rtbMethods.text = ""  
    $rtbMethods.selectionFont  = $fontBold  
     
    $rtbMethods.AppendText(("{1} Method : {0} `n" -f $this.SelectedItem , $script:WmiClass.__Class))  
    $rtbMethods.AppendText("`n")  
    $rtbMethods.selectionFont  = $fontBold  
    $rtbMethods.AppendText("OverloadDefinitions:`n")  
    $rtbMethods.AppendText("$($script:WmiClass.$WMIMethod.OverloadDefinitions)`n`n")  
    $Qualifiers=@()  
    $script:WmiClass.psbase.Methods[($this.SelectedItem)].Qualifiers | ForEach-Object {$qualifiers += $_.name}  
    #$rtbMethods.AppendText( "$qualifiers`n" )  
    $static = $Qualifiers -Contains "Static"   
    $rtbMethods.selectionFont  = $fontBold  
    $rtbMethods.AppendText( "Static : $static`n" )  
    If ($static) {   
         $rtbMethods.AppendText( "A Static Method does not an Instance to act upon`n`n" )  
         $rtbMethods.AppendText("`n")  
     
         $rtbMethods.SelectionColor = 'Green'  
         $rtbMethods.SelectionFont = $fontCode  
         $rtbMethods.AppendText("# Sample Of Connecting to a WMI Class`n`n")  
         $rtbMethods.SelectionColor = 'Black'  
         $rtbMethods.SelectionFont = $fontCode  
         $SB = new-Object text.stringbuilder  
         $SB = $SB.Append('$Computer = "') ; $SB = $SB.AppendLine(".`"")  
         $SB = $SB.Append('$Class = "') ; $SB = $SB.AppendLine("$WmiClassName`"")    
         $SB = $SB.Append('$Method = "') ; $SB = $SB.AppendLine("$WmiMethod`"`n")  
         $SB = $SB.AppendLine('$MC = [WmiClass]"\\$Computer\' + "$($script:WmiClass.__NAMESPACE)" + ':$Class"')    
         #$SB = $SB.Append('$MP.Server = "') ; $SB = $SB.AppendLine("$($MP.Server)`"")    
         #$SB = $SB.Append('$MP.NamespacePath = "') ; $SB = $SB.AppendLine("$($script:WmiClass.__NAMESPACE)`"")    
         #$SB = $SB.AppendLine('$MP.ClassName = $Class')  
         $SB = $SB.AppendLine("`n")     
         #$SB = $SB.AppendLine('$MC = new-object system.management.ManagementClass($MP)')    
         $rtbMethods.AppendText(($sb.tostring()))  
         $rtbMethods.SelectionColor = 'Green'  
         $rtbMethods.SelectionFont = $fontCode  
         $rtbMethods.AppendText("# Getting information about the methods`n`n")  
         $rtbMethods.SelectionColor = 'Black'  
         $rtbMethods.SelectionFont = $fontCode  
         $rtbMethods.AppendText(  
             '$mc' + "`n" +  
             '$mc | Get-Member -membertype Method' + "`n" +  
             "`$mc.$WmiMethod"  
         )  
    } Else {  
         $rtbMethods.AppendText( "This is a non Static Method and needs an Instance to act upon`n`n" )  
         $rtbMethods.AppendText( "The Example given will use the Key Properties to Connect to a WMI Instance : `n`n" )  
         $rtbMethods.SelectionColor = 'Green'  
         $rtbMethods.SelectionFont = $fontCode  
         $rtbMethods.AppendText("# Example Of Connecting to an Instance`n`n")  
     
         $rtbMethods.SelectionColor = 'Black'  
         $rtbMethods.SelectionFont = $fontCode  
         $SB = new-Object text.stringbuilder  
         $SB = $SB.AppendLine('$Computer = "."')  
         $SB = $SB.Append('$Class = "') ; $SB = $SB.AppendLine("$WmiClassName.`"")    
         $SB = $SB.Append('$Method = "') ; $SB = $SB.AppendLine("$WMIMethod`"")  
         $SB = $SB.AppendLine("`n# $WmiClassName. Key Properties :")    
         $Filter = ""    
         $script:WmiClass.psbase.Properties | ForEach-Object {    
           $Q = @()  
           $_.psbase.Qualifiers | ForEach-Object {$Q += $_.name}   
           $key = $Q -Contains "key"   
           If ($key) {    
             $CIMType = $_.psbase.Qualifiers["Cimtype"].Value    
             $SB = $SB.AppendLine("`$$($_.Name) = [$CIMType]")    
             $Filter += "$($_.name) = `'`$$($_.name)`'"     
           }    
         }    
         $SB = $SB.Append("`n" + '$filter=');$SB = $SB.AppendLine("`"$filter`"")    
         $SB = $SB.AppendLine('$MC = get-WMIObject $class -computer $Computer -Namespace "' +  
             "$($script:WmiClass.__NAMESPACE)" + '" -filter $filter' + "`n")  
         $SB = $SB.AppendLine('# $MC = [Wmi]"\\$Computer\Root\CimV2:$Class.$filter"')   
         $rtbMethods.AppendText(($sb.tostring()))  
    }   
    $SB = $SB.AppendLine('$InParams = $mc.psbase.GetMethodParameters($Method)')  
    $SB = $SB.AppendLine("`n")  
    # output Method Parameter Help  
    $rtbMethods.selectionFont  = $fontBold  
    $rtbMethods.AppendText("`n`n$WmiClassName. $WMIMethod Method :`n`n")   
    $q = $script:WmiClass.PSBase.Methods[$WMIMethod].Qualifiers | foreach {$_.name}  
    if ($q -contains "Description") {  
         $rtbMethods.AppendText(($script:WmiClass.psbase.Methods[$WMIMethod].psbase.Qualifiers["Description"].Value))  
    }   
   
    $rtbMethods.selectionFont  = $fontBold  
    $rtbMethods.AppendText("`n`n$WMIMethod Parameters :`n")   
  # get the Parameters   
    
  $inParam = $script:WmiClass.psbase.GetMethodParameters($WmiMethod)  
  $HasParams = $False   
  if ($true) {   
    trap{$rtbMethods.AppendText('[None]') ;continue}    
    $inParam.PSBase.Properties | foreach {   
      $Q = $_.Qualifiers | foreach {$_.name}  
      # if Optional Qualifier is not present then Parameter is Mandatory   
      $Optional = $q -contains "Optional"  
      $CIMType = $_.Qualifiers["Cimtype"].Value   
      $rtbMethods.AppendText("`nName = $($_.Name) `nType = $CIMType `nOptional = $Optional")  
      # write Parameters to Example script   
      if ($Optional -eq $TRUE) {$SB = $SB.Append('# ')}   
      $SB = $SB.Append('$InParams.');$SB = $SB.Append("$($_.Name) = ");$SB = $SB.AppendLine("[$CIMType]")   
      if ($q -contains "Description") {$rtbMethods.AppendText($_.Qualifiers["Description"].Value)}  
      $HasParams = $true    
    }   
  }  
  # Create the Rest of the Script  
  $rtbMethods.selectionFont  = $fontBold  
  $rtbMethods.AppendText("`n`nTemplate Script :`n")   
  # Call diferent Overload as Method has No Parameters   
  If ($HasParams -eq $True) {   
      $SB = $SB.AppendLine("`n`"Calling $WmiClassName. : $WMIMethod with Parameters :`"")   
      $SB = $SB.AppendLine('$inparams.PSBase.properties | select name,Value | format-Table')   
      $SB = $SB.AppendLine("`n" + '$R = $mc.PSBase.InvokeMethod($Method, $inParams, $Null)')   
  }Else{   
      $SB = $SB.AppendLine("`n`"Calling $WmiClassName. : $WMIMethod `"")   
      $SB = $SB.AppendLine("`n" + '$R = $mc.PSBase.InvokeMethod($Method,$Null)')   
  }   
  $SB = $SB.AppendLine('"Result :"')   
  $SB = $SB.AppendLine('$R | Format-list' + "`n`n")  
  # Write Header of the Sample Script :   
   
  $rtbMethods.SelectionColor = 'Green'  
  $rtbMethods.SelectionFont = $fontCode  
  $rtbMethods.AppendText(@"  
# $WmiClassName. $WMIMethod-Method Template Script"   
# Created by PowerShell WmiExplorer  
# /\/\o\/\/ 2006  
# www.ThePowerShellGuy.com  
#  
# Fill InParams values before Executing   
# InParams that are Remarked (#) are Optional  
"@  
  )  
  $rtbMethods.SelectionColor = 'Black'  
  #$rtbMethods.SelectionFont = $fontCode  
  $rtbMethods.AppendText("`n`n" + $SB)  
  $rtbMethods.SelectionFont = new-object System.Drawing.Font("Lucida Console",6 )  
  $rtbMethods.AppendText("`n`n Generated by the PowerShell WMI Explorer  /\/\o\/\/ 2006" )  
         
} # GetWmiMethod  
#endregion  
#endregion  

# Show the Form  
$FrmMain.Add_Shown({$FrmMain.Activate()})  
   
trap {Write-Host $_ ;$status.Text = "unexpected error";$slMessage.Text = "$_.message";continue}  
& {  
    [void]$FrmMain.showdialog()  
}  
# Resolve-Error $Error[0] | out-string 
 
 =================================
 
 cls
Set-StrictMode -Version 2
#region host preparations
if ($Host.Name -eq 'ConsoleHost')
{
    Add-Type -AssemblyName System.Windows.Forms;
    Add-Type -AssemblyName System.Drawing;
}
#endregion host preparations
#region the resulting wizard
    #region adjustable settings
        #region controls settings
        # Form size and caption
        [string]$script:constWizardInitialCaption = `
            'This is a sample wizard';
        [int]$script:constWizardWidth = `
            [System.Windows.Forms.SystemInformation]::VirtualScreen.Width / 2;
        [int]$script:constWizardHeight = `
            [System.Windows.Forms.SystemInformation]::VirtualScreen.Height / 2;
        # Buttons Next, Back
        [int]$script:constButtonHeight = 23;
        [int]$script:constButtonWidth = 75;
        [int]$script:constButtonAreaHeight = `
            $script:constButtonHeight * 2;
        [string]$script:constButtonNextName = '&Next';
        [string]$script:constButtonBackName = '&Back';
        # Step display name and description
        [int]$script:constLabelStepDisplayNameLeft = 5;
        [int]$script:constLabelStepDisplayNameTop = 0;
        [float]$script:constLabelStepDisplayNameFontSize = 16;
        [int]$script:constLabelStepDescriptionLeft = 5;
        [int]$script:constLabelStepDescriptionTop = 30;
        # Form properties
        [bool]$script:constWizardRigthToLeft = $false;
        # Initial step number
        [int]$script:currentStep = 0;
        #endregion controls settings
    #endregion adjustable settings
    #region mandatory settings
        #region Initialization of the SplitContainer controls
        # The outer split container
        [System.Windows.Forms.SplitContainer]$script:splitOuter = `
            New-Object System.Windows.Forms.SplitContainer;
            $script:splitOuter.Dock = [System.Windows.Forms.DockStyle]::Fill;
            $script:splitOuter.IsSplitterFixed = $true;
            $script:splitOuter.Orientation = `
                [System.Windows.Forms.Orientation]::Horizontal;
            $script:splitOuter.SplitterWidth = 5;
        # The inner split container
        [System.Windows.Forms.SplitContainer]$script:splitInner = `
            New-Object System.Windows.Forms.SplitContainer;
            $script:splitInner.Dock = [System.Windows.Forms.DockStyle]::Fill;
            $script:splitInner.IsSplitterFixed = $true;
            $script:splitInner.Orientation = `
                [System.Windows.Forms.Orientation]::Horizontal;
            $script:splitInner.SplitterWidth = 5;
            $script:splitOuter.Panel1.Controls.Add($script:splitInner);
        # The labels for the curent step name and description
        [System.Windows.Forms.Label]$script:lblStepDisplayName = `
            New-Object System.Windows.Forms.Label;
            $script:lblStepDisplayName.Left = `
                $script:constLabelStepDisplayNameLeft;
            $script:lblStepDisplayName.Top = `
                $script:constLabelStepDisplayNameTop;
            [System.Drawing.Font]$private:font = `
                $script:lblStepDisplayName.Font;
            $private:font = `
                New-Object System.Drawing.Font($private:font.Name, `
                        $script:constLabelStepDisplayNameFontsize, `
                        $private:font.Style, $private:font.Unit, `
                        $private:font.GdiCharSet, $private:font.GdiVerticalFont );
            $script:lblStepDisplayName.Font = $private:font;
        [System.Windows.Forms.Label]$script:lblStepDescription = `
            New-Object System.Windows.Forms.Label;
            $script:lblStepDescription.Left = `
                $script:constLabelStepDescriptionLeft;
            $script:lblStepDescription.Top = `
                $script:constLabelStepDescriptionTop;
        $script:splitInner.Panel1.Controls.AddRange(($script:lblStepDisplayName, `
            $script:lblStepDescription));
        #endregion Initialization of the SplitContainer controls
        #region the Next and Back buttons
        [System.Windows.Forms.Button]$script:btnNext = `
            New-Object System.Windows.Forms.Button;
        $script:btnNext.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
            [System.Windows.Forms.AnchorStyles]::Right -bor `
            [System.Windows.Forms.AnchorStyles]::Top;
        $script:btnNext.Text = $script:constButtonNextName;
        [System.Windows.Forms.Button]$script:btnBack = `
            New-Object System.Windows.Forms.Button;
        $script:btnBack.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
            [System.Windows.Forms.AnchorStyles]::Right -bor `
            [System.Windows.Forms.AnchorStyles]::Top;
        $script:btnBack.Text = $script:constButtonBackName;
        $script:btnBack.Enabled = $false;
        $script:splitOuter.Panel2.Controls.AddRange(($script:btnBack, $script:btnNext));
        #endregion the Next and Back buttons
        #region Initialization of the main form
        $script:frmWizard = $null;
        [System.Windows.Forms.Form]$script:frmWizard = `
            New-Object System.Windows.Forms.Form;
        $script:frmWizard.Controls.Add($script:splitOuter);
 
        if ($script:constWizardRigthToLeft)
        {
            $script:frmWizard.RightToLeft = `
                [System.Windows.Forms.RightToLeft]::Yes;
            $script:frmWizard.RightToLeftLayout = $true;
        }
        else
        {
            $script:frmWizard.RightToLeft = `
                [System.Windows.Forms.RightToLeft]::No;
            $script:frmWizard.RightToLeftLayout = $false;
        }
        $script:frmWizard.Text = $script:constWizardInitialCaption;
        #endregion Initialization of the main form
    #endregion mandatory settings
    #region the Wizard steps
    [System.Collections.ArrayList]$script:wzdSteps = `
        New-Object System.Collections.ArrayList;
        # Here we create an 'enumeration' (PSObject)
        # and begin filling ArrayList $script:wzdSteps with Panel controls
        [System.EventHandler]$script:hndlRunControlsAdd = `
            {try{$script:splitInner.Panel2.Controls.Add($script:wzdSteps[$script:currentStep]);}catch{Write-Debug $Error[0]; Write-Debug $global:Error[0];}};
            #region function New-WizardStep
        function New-WizardStep
        {
            param(
                  [string]$StepName,
                  [string]$StepDisplayName,
                  [string]$StepDescription = ''
                  )
            # Storing parameters in step arrays
            Add-Member -InputObject $script:steps -MemberType NoteProperty `
                -Name $StepName -Value $script:wzdSteps.Count;
            $null = $script:stepDisplayNames.Add($StepDisplayName);
            $null = $script:stepDescriptions.Add($StepDescription);
            # Create and add the new step's panel to the array
            [System.Windows.Forms.Panel]$private:panel = `
                New-Object System.Windows.Forms.Panel;
            $null = $script:wzdSteps.Add($private:panel);
            $script:currentStep = $script:wzdSteps.Count - 1;
 
            $script:splitInner.Panel2.Controls.Add($script:wzdSteps[$script:currentStep]);
 
            $script:wzdSteps[$script:currentStep].Dock = `
                [System.Windows.Forms.DockStyle]::Fill;
            # To restore initial state for this code running before the user accesses the wizard.
            $script:currentStep = 0;
        }
            #endregion function New-WizardStep
            #region function Add-ControlToStep
        function Add-ControlToStep
        {
            param(
                  [string]$StepNumber,
                  [string]$ControlType,
                  [string]$ControlName,
                  [int]$ControlTop,
                  [int]$ControlLeft,
                  [int]$ControlHeight,
                  [int]$ControlWidth,
                  [string]$ControlData
                 )
            $private:ctrl = $null;
            try{
                $private:ctrl = New-Object $ControlType;
            }catch{Write-Error "Unable to create a control of $($ControlType) type";}
            try{
                $private:ctrl.Name = $ControlName;
            }catch{Write-Error "Unable to set the Name property with value $($ControlName) to a control of the $($ControlType) type";}
            try{
                $private:ctrl.Top = $ControlTop;
            }catch{Write-Error "Unable to set the Top property with value $($ControlTop) to a control of the $($ControlType) type";}
            try{
                $private:ctrl.Left = $ControlLeft;
            }catch{Write-Error "Unable to set the Left property with value $($ControlLeft) to a control of the $($ControlType) type";}
            try{
                $private:ctrl.Height = $ControlHeight;
            }catch{Write-Error "Unable to set the Height property with value $($ControlHeight) to a control of the $($ControlType) type";}
            try{
                $private:ctrl.Width = $ControlWidth;
            }catch{Write-Error "Unable to set the Width property with value $($ControlWidth) to a control of the $($ControlType) type";}
            try{
                $private:ctrl.Text = $ControlData;
            }catch{Write-Error "Unable to set the Text property with value $($ControlData) to a control of the $($ControlType) type";}
            try{
                $wzdSteps[$StepNumber].Controls.Add($private:ctrl);
            }catch{Write-Error "Unable to add a control of $($ControlType) type to the step $($StepNumber)";}
        }
            #endregion function Add-ControlToStep
        # Step data arrays
        [psobject]$script:steps = New-Object psobject;
        [System.Collections.ArrayList]$script:stepDisplayNames = `
            New-Object System.Collections.ArrayList;
        [System.Collections.ArrayList]$script:stepDescriptions = `
            New-Object System.Collections.ArrayList;
    #endregion the Wizard steps
    #region events of the wizard controls
        #region resizing
            #region function Initialize-WizardControls
    function Initialize-WizardControls
    <#
        .SYNOPSIS
            The Initialize-WizardControls function sets the wizard common controls to predefined positions.
 
        .DESCRIPTION
            The Initialize-WizardControls function does the following:
            - sets Top, Left, Width and Height properties of the wizard buttons
            - positions of the step labels
            - sets the SplitterDistance property of both Splitcontainer controls
 
        .EXAMPLE
            PS C:\> Initialize-WizardControls
            This example shows how to step the wizard forward.
 
        .INPUTS
            No input
 
        .OUTPUTS
            No output
    #>
    {
        # Set sizes of buttons
        $script:btnNext.Height = $script:constButtonHeight;
        $script:btnNext.Width = $script:constButtonWidth;
        $script:btnBack.Height = $script:constButtonHeight;
        $script:btnBack.Width = $script:constButtonWidth;
        # SplitterDistance of the outer split container
        # in other words, the area where Next and Back buttons are placed
        $script:splitOuter.SplitterDistance = `
            $script:splitOuter.Height - `
            $script:constButtonAreaHeight;
        #if ($script:splitOuter.SplitterDistance -lt 0)
        #{$script:splitOuter.SplitterDistance = 10;}
        #$script:splitOuter.SplitterDistance = `
        #   $script:splitOuter.Height - `
        #   $script:constButtonAreaHeight;
 
        # Placements of the buttons
        if ($script:constWizardRigthToLeft)
        {
            $script:btnNext.Left = 10;
            $script:btnBack.Left = $script:constButtonWidth + 20;
        }
        else
        {
            $script:btnNext.Left = $script:splitOuter.Width - `
                $script:constButtonWidth - 10;
            $script:btnBack.Left = $script:splitOuter.Width - `
                $script:constButtonWidth - `
                $script:constButtonWidth - 20;
        }
        $script:btnNext.Top = `
            ($script:constButtonAreaHeight - $script:constButtonHeight) / 2;
        $script:btnBack.Top = `
                ($script:constButtonAreaHeight - $script:constButtonHeight) / 2;
 
        # SplitterDistance of the inner split container
        # this is the place where step name is placed
        $script:splitInner.SplitterDistance = `
            $script:constButtonAreaHeight * 1.5;
            #$script:splitOuter.Panel2.Height * 1.5;
        #if ($script:splitInner.SplitterDistance -lt 0)
        #{$script:splitInner.SplitterDistance = 10;}
 
        # Step Display Name and Description labels
        $script:lblStepDisplayName.Width = `
            $script:splitInner.Panel1.Width - `
            $script:constLabelStepDisplayNameLeft * 2;
        $script:lblStepDescription.Width = `
            $script:splitInner.Panel1.Width - `
            $script:constLabelStepDescriptionLeft * 2;
 
        # Refresh after we have changed placements of the controls
        [System.Windows.Forms.Application]::DoEvents();
    }
            #endregion function Initialize-WizardControls
    [System.EventHandler]$script:hndlFormResize = {Initialize-WizardControls;}
    [System.EventHandler]$script:hndlFormLoad = {Initialize-WizardControls;}
    #[System.Windows.Forms.MouseEventHandler]$script:hndlSplitMouseMove = `
    #   {Initialize-WizardControls;}
    # Initial arrange on Load form.
    $script:frmWizard.add_Load($script:hndlFormLoad);
    #$script:frmWizard.add_Resize($script:hndlFormResize);
    $script:splitOuter.add_Resize($script:hndlFormResize);
        #endregion resizing
        #region steps
            #region function Invoke-WizardStep
    function Invoke-WizardStep
    <#
        .SYNOPSIS
            The Invoke-WizardStep function sets active panel on the wizard form.
 
        .DESCRIPTION
            The Invoke-WizardStep function does the following:
            - changes internal variable $script:currentStep
            - sets/resets .Enabled property of btnNext and btnBack
            - changes .Dock and .Left properties of every panel
 
        .PARAMETER  Forward
            The optional parameter Forward is used to point out the direction the wizard goes.
 
        .EXAMPLE
            PS C:\> Invoke-WizardStep -Forward $true
            This example shows how to step the wizard forward.
 
        .INPUTS
            System.Boolean
 
        .OUTPUTS
            No output
    #>
    {
        [CmdletBinding()]
        param(
              [Parameter(Mandatory=$false)]
              [bool]$Forward = $true
              )
        Begin{}
        Process{
        if ($Forward)
        {
            $script:btnBack.Enabled = $true;
            if ($script:currentStep -lt ($script:wzdSteps.Count - 1))
            {$script:currentStep++;}
            if ($script:currentStep -lt ($script:wzdSteps.Count - 1))
            {$script:btnNext.Enabled = $true;}
            else
            {$script:btnNext.Enabled = $false;}
        }
        else
        {
            $script:btnNext.Enabled = $true;
            if ($script:currentStep -gt 0)
            {$script:currentStep--;}
            if ($script:currentStep -gt 0)
            {$script:btnBack.Enabled = $true;}
            else
            {$script:btnBack.Enabled = $false;}
        }
        for($private:i = 0; $private:i -lt $script:wzdSteps.Count;
            $private:i++)
        {
            if ($private:i -ne $script:currentStep)
            {
                $script:wzdSteps[$private:i].Dock = `
                    [System.Windows.Forms.DockStyle]::None;
                $script:wzdSteps[$private:i].Left = 10000;
            }
            else
            {
                $script:wzdSteps[$private:i].Dock = `
                    [System.Windows.Forms.DockStyle]::Fill;
                $script:wzdSteps[$private:i].Left = 0;
            }
        }
        $script:lblStepDisplayName.Text = `
            $script:stepDisplayNames[$script:currentStep];
        $script:lblStepDescription.Text = `
            $script:stepDescriptions[$script:currentStep];
        }
        End{}
    }
            #endregion function Invoke-WizardStep
            #region function Initialize-WizardStep
    function Initialize-WizardStep
    # This is the selfsufficient function doing all the necessary
    # calculations for controls on each panel.
    # Also from the code can be seen how to address the panel you are interesting in
    # using the 'enumeration' created earlier
    # for example, $script:wzdSteps[$script:steps.Welcome]
    {
        $script:lblStepDisplayName.Text = `
            $script:stepDisplayNames[$script:currentStep];
        $script:lblStepDescription.Text = `
            $script:stepDescriptions[$script:currentStep];
    }
            #endregion function Initialize-WizardStep
    [System.EventHandler]$hndlStepForward = {
        # serve controls' data
        Initialize-WizardStep;
        # switch panels
        Invoke-WizardStep $true;
    }
    [System.EventHandler]$hndlStepBack = {
        # switch panels
        Invoke-WizardStep $false;
    }
    $script:btnNext.add_Click($script:hndlStepForward);
    $script:btnBack.add_Click($script:hndlStepBack);
        #endregion steps
    #endregion events of the wizard controls
    #region wizard initialization
        #region function Initialize-Wizard
    function Initialize-Wizard
    # This is one more selfsufficient function written to make
    # the latest preparations for the form run
    {
        #region control settings
        $script:frmWizard.Width = $script:constWizardWidth;
        $script:frmWizard.Height = $script:constWizardHeight;
        #endregion control settings
        Initialize-WizardStep;
    }
        #endregion function Initialize-Wizard
    #endregion wizard initialization
#endregion the resulting wizard
 
        # Step 1: Welcome
        New-WizardStep 'Welcome' 'This is the first step' 'Welcome to the PowerShell Wizard, the our dear customer!';
        # Add a label
        # Please note that we can use the enumeration $steps which is being created runtime
        # on a call of the New-WizardStep function
        Add-ControlToStep $steps.Welcome System.Windows.Forms.Label 'lblWelcome' 20 10 50 300 'This Wizard carries you through the steps you need to collect the files from a given path';
 
        # Step 2
        New-WizardStep 'Input' 'Step Two' 'Here you type some in controls, plz';
        # Add a label
        Add-ControlToStep $steps.Input System.Windows.Forms.Label 'lblInput' 20 10 20 300 'Please type the path to a catalog';
        # Add a text box
        Add-ControlToStep $steps.Input System.Windows.Forms.TextBox 'txtInput' 40 10 20 300
 
        # Unfortunately, there is no way right now to test the user input with ease
        # So we can disable the Next button manually (to be improved soon)
        # this code works wrong:
        #if ($wzdSteps[$steps.Input].Controls['txtInput'].Text.Length -eq 0)
        #{
        #   $btnNext.Enabled = $false;
        #}
        #else
        #{
        #   $btnNext.Enabled = $true;
        #}
 
        # Step 3
        New-WizardStep 'Progress' 'The third one' 'Wait, please. Sip a coffee' ;
        # Add a progress bar
        Add-ControlToStep $steps.Progress 'System.Windows.Forms.ProgressBar' 'pbDir' 200 50 100 400
 
        # Step 4
        New-WizardStep 'Output' 'Fourth' 'Now awake and read the output';
        # Add a list box
        Add-ControlToStep $steps.Output System.Windows.Forms.ListBox lbxFiles 50 50 300 400
 
        # Step 5: Finish
        New-WizardStep 'Finish' 'Finish!' 'Bye!';
 
        Initialize-Wizard;
        # Set the second step as active
        Invoke-WizardStep -Forward $true;
 
        $script:frmWizard.ShowDialog() | Out-Null;
