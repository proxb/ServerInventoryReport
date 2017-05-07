<#
    .SYNOPSIS
        Presents a GUI to users that automatically connects to a SQL database to display
        a server inventory report.

    .DESCRIPTION
        Presents a GUI to users that automatically connects to a SQL database to display
        a server inventory report.            

    .NOTES
        Name: ServerInventory.ps1
        Author: Boe Prox
        Version History:
            1.2.3.0 //Boe Prox - 07 May 2017
                - Fixed bug where Excel report only worked with up to 28 columns
            1.2.2.0 //Boe Prox - 5 May 2016
                - Added more keyboard shortcuts: Display Help (F1), Apply Filter (F5), Clear Filter (F8), Create Excel Report (CTRL+R)
                - Added help window for keyboard shortcuts and filter definitions
                - Updated icons for Apply and Clear buttons
            1.2.1.0 //Boe Prox - 4 May 2016
                - Added menu item for Edit>Go To Computer
                - Added keyboard shortcuts for Exit (CTRL+E),Go To Computer (CTRL+G) menu items and selecting All in treeview (CTRL+A)
            1.2.0.0 //Boe Prox - 3 May 2016
                - Added UI enhancements for progress and Status bar
                - Hid excel window from displaying when report is generated
            1.1.2.0 //Boe Prox - 2 May 2016
                - Fixed bug where LIKE statements would fail against integers
                - Added IS NULL and IS NOT NULL filter options
            1.1.1.0 //Boe Prox - 29 April 2016
                - Added more menus for additional filters
            1.1.0.0 //Boe Prox - 27 April 2016
                - Fixed bug with Filter text box not showing Red on filter error
                - Restrict users from resizing Excel Report window
                - Updated comment based help
            1.0.0.0 //Boe Prox - 22 Mar 2016
                - Initial version

    .EXAMPLE
        .\ServerInventory.ps1

        Description
        -----------
        Displays the ServerInventory UI
#>

$RunSpace=[RunspaceFactory]::CreateRunspace()
$RunSpace.ApartmentState = "STA"
$RunSpace.Open()
$PowerShell = [PowerShell]::Create()
$PowerShell.Runspace = $RunSpace
[void]$PowerShell.AddScript({ 
    #region Load Assemblies
    Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration, Microsoft.VisualBasic
    #endregion Load Assemblies

    #region User Defined Variables
    $SQLServer = 'vsql'
    ## Update Tables to include to new tables added to SQL so UI controls will be auto generated at runtime
    $Script:Tables = 'tbGeneral','tbOperatingSystem', 'tbNetwork','tbMemory','tbProcessor','tbUsers','tbGroups',
    'tbDrives','tbAdminShare','tbUserShare','tbServerRoles','tbSoftware','tbScheduledTasks','tbUpdates','tbServices'
    #endregion User Defined Variables

    #region Variables
    $Script:ExcelReports = New-Object System.Collections.ArrayList
    $ExcludeProperties = 'DataView','RowVersion','Row','IsNew','IsEdit','Error'
    $_Number = 0
    $Previous = 0
    $Script:Letters = [hashtable]::Synchronized(@{})
    $Range = 97..122
    For ($n=0;$n -le 5; $n++) {
        For ($i=1;$i -le $Range.count; $i++) {
            $_Number = $Previous + $i
            If ($n -ne 0) {
                $Script:Letters[$_Number] = "$([char]($Range[$n-1]))$([char]($Range[$i-1]))"
            }
            Else {
                $Script:Letters[$_Number] = [char]($Range[$i-1])
            }            
        }
        $Previous = $_number
    }
    $Script:UIHash = [hashtable]::Synchronized(@{Host=$Host})
    $Script:TempFilters = [hashtable]::Synchronized(@{})
    $Script:Filters = [hashtable]::Synchronized(@{})
    $Script:JobCleanup = [hashtable]::Synchronized(@{
        Host = $host
    })
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
    $ExcelIcon = 'AAABAAEAICAQAgAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgA
    IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAM
    DAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP//AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIiIiIiIiIiIiIiIiIiIgAC//////////
    ////////IAAv//////iP//+Id4//diAAL/////////iPiP/4//YgAC////////iIj//4d
    4/3IAAvj/////+I///4h4j/+CAAL4////////+I+I//j/8gAC+P//////+IiP//h3j/cA
    A/iDMzMyKPhzM4+IiP/4AAP4hzMzIij3Mzf4j/+I/wAD+IhzMyIogzM///+Hd/+AA/iIg
    yInj3Mzj/+IiI//8AP4iPciJvd3d/iIiP////AD+Ij/YiJ3d3/4j////4YAA/iIj4Iid3
    eP////+HJyAAP4iI/3J3d3////hyZ38gAD+IiP/3d3f//4d3d3d/IAA/iIiP93d3j4d3d
    3d3fyAAP4iIj3d3dm93d3d3d38gAD+IiIh3d2Zn93d3d3d/IAA/iIiHN3hmZo+IiIh3fy
    AAP/iIczN/dmZviIiIiI8gADj4iDMz/4dmZ/iIiIiPIAA3+IczOPiIZmaPiIiIjyAAM4+
    IiIiIiIiIiIiIiI8gAAN/+IiIiIiIiIiIiIiPIAAAN4//iIiIiIiIiIiIjyAAAAM3j///
    //////////8gAAAAMzMzMzMzMzMyIiIiIAD/////AAAAAQAAAAEAAAABAAAAAQAAAAEAA
    AABAAAAAQAAAAEAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAEAAAABAA
    AAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABgAAAAcAAAAHgAAAB8AAAAQ=='
    $ExcelBytes = [convert]::FromBase64String($ExcelIcon)
    $ExcelBitMap = New-Object System.Windows.Media.Imaging.BitmapImage
    $ExcelBitMap.BeginInit()
    $ExcelBitMap.StreamSource = [System.IO.MemoryStream]$ExcelBytes
    $ExcelBitMap.EndInit()

    $ClearIcon = 'AAABAAEAICAQAgAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAA
    gAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAM
    DAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP//AAD///8A//////iId3d3d4iP////////+HcA
    AAAAAAAAAHeP////+HAAAAAAAAAAAAAAAI////cAAAAAMzEAAAAAAAB/////cAAAA7szMA
    AAAAAHj/////+HcDu7M7MAAAd3j////////4M7u7uzOIj///////////g7u7u7uzj/////
    ///////4s7O7u7u3////////////+Du7u7u7s4////////////+Lu7u7u7t///////////
    //87u7u7u7P/////////////g7u7uLuz//////////////O7u7i7t//////////////3u7
    u7s5f/////////////+Du7t5k3//////////////+DuzmTs///////////////+HOXO7P/
    ///////////////4c7u3//////////////////gzN///////////////////hzP///////
    ////////////gzj///////////////////c3///////////////////4c4////////////
    //////+Hd///////////////////93N///////////////////h3OP////////////////
    //dzf//////////////////4dzj//////////////////3d4///////////////////4j/
    //////////////////////////wAP//AAAH/AAAAfwAAAH8AAAD/wAAD//AA///gAP//4A
    B//+AAf//wAH//8AA///gAP//4AD///AA///wAf//+AD///gA////Af///8H////A////8
    P////B////4f///+D////g////8H////A////4P///+D////x//////w=='
    $ClearBytes = [convert]::FromBase64String($ClearIcon)
    $ClearBitMap = New-Object System.Windows.Media.Imaging.BitmapImage
    $ClearBitMap.BeginInit()
    $ClearBitMap.StreamSource = [System.IO.MemoryStream]$ClearBytes
    $ClearBitMap.EndInit()

    $FilterIcon = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0I
    Ars4c6QAAAARnQU1BAACxjwv8YQUAAAPHSURBVEhLrVVbaBxVGP52ZneSzYWIWpo0qS1WS
    mtrqkI0RUSEIhakLxabIuJ77YsI0r6oFV/si+DtSfomWvGCoD4oPlSR2pdG3baGXiW72ew
    2iTHZ2cucOeeM3z87bROT0Hnwg7Mzey7f9//n+8+ZDBI8/nHxWS/rPNTbk3vUBXp6utzRZ
    AhRBOQ7XHR7Tvy/riyagUHm5mrAb5hfDR81PzyjtB3/+fmNX0h/POWpz6ei3u4c8m4GGXJ
    IJzkRWRm9BRESLCUW3FjTRgZNY7FQC/H9/sH21NETV6e3bejsn2FUv1zykaWQ68hQe5n85
    piWs4QmpJhMyUlQfDd8Jy/W5S229GUx5aNy+sV7B7Iy2Wo9kjW2+M2+oXixxD8xU49Dttb
    A8nmq2kLFNxyx0DbC/o0eJuYCnC41YKmwpcPA4dwnBjtw7IILE5gRYboZ0vB7l8ee3pH/5
    O0nNyQ9DCvSVGdYWvaK5IZiDNNIH9u3F/9Bua7gsC/UGk5nF87M53C+Gh0cP3Tfp8JyK2d
    ixzt/vr/34TteGljfB6UN+nNOHO2mfASPURq+WxJHzKhG0XdPzaPTAxZbgIKHu+/Konq9/
    sFvh7ceTiiXCwi2HT/3xyt7Bx/42/FQaVnuL4mZjOWPmCzksvkXiz5myw2+ZmA4tumeLhQ
    nF348+/L9exKqGO26W4KJV3cOf/hDubS9O0KWZBGjdqRxy6SJSo4CcyUfGWZhghD96z0Ur
    86X/ksuWCEg0ErvPvLZFeyjYY2Q+8sIpSk2ceMvkmtmZ5RGZz6Lmcoix83u9urlWFWgcGR
    XyW2p4aNfXcaeAQ91+hHSSM0DZhj1TMVnZjo+DzpsQQV2WNYky5dhVQFB4bWRwlzZP/Dde
    BUP3pmDYuFrxl9XBrPTfizk0fywERwoHN1VSJatAI/P2pj96aPzC5uf2755XcdO43loqgh
    T5RpqswpuzkLVmifPvfHIsWT6qlgzgxu4dvyxsbNXFtDrZdCgwWHTwqoA+R4XF94aHUumr
    YnbCggMjWbwsPLkYdP0xARi9+2RSsCSVCpI8ZBp7r0NNc8G784USJcBSUOSB2I0bw/DH9P
    6HwXiDDT3n0Jaome5yp2UBqkE5B4W8oAiJm4sWLnwUiCVQESyJomVmByGvHtoMlsapPOAB
    gfMIqCQkQqSDJhRGqQT4IekSULFWg2ZiTX0gRdgGqQzOa7/DMuUdw/Jo7nrsIv84qVAOg9
    4oYjJ4bUqspOTjb6hgTeLJ55Z8S1ZDakmDb3wZaTqzd/dnPP69MmDXyfdKQD8C3ZgJm6fk
    xTAAAAAAElFTkSuQmCC'
    $FilterBytes = [convert]::FromBase64String($FilterIcon)
    $FilterBitMap = New-Object System.Windows.Media.Imaging.BitmapImage
    $FilterBitMap.BeginInit()
    $FilterBitMap.StreamSource = [System.IO.MemoryStream]$FilterBytes
    $FilterBitMap.EndInit()
    #endregion Variabled

    #region Helper Functions
    Function Invoke-ExcelReport {
        $Script:ExcelReports.Clear()
        $Return = $ExcelReport.Invoke()    
        If ($Return -AND $Script:ExcelReports.Count -gt 0) {
            $UIHash.ProgressBar.Maximum = $ExcelReports.Count
            ## Begin runspace initialization
            $newRunspace =[runspacefactory]::CreateRunspace()
            $newRunspace.ApartmentState = "STA"
            $newRunspace.ThreadOptions = "ReuseThread"          
            $newRunspace.Open()
            $newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)          
            $newRunspace.SessionStateProxy.SetVariable("ExcelReports",$ExcelReports)     
            $newRunspace.SessionStateProxy.SetVariable("ExcludeProperties",$ExcludeProperties) 
            $newRunspace.SessionStateProxy.SetVariable("ReportPath",$ReportPath) 
            $newRunspace.SessionStateProxy.SetVariable("Letters",$Letters) 
            $PowerShell = [PowerShell]::Create().AddScript({  
                Add-Type –assemblyName System.Windows.Forms
                Function Show-ReportLocation {
                    Param($ReportLocation)
                    $title = "Report Completed"
                    $message = "The report has been saved to: $ReportLocation"
                    $button = [System.Windows.Forms.MessageBoxButtons]::OK
                    $icon = [Windows.Forms.MessageBoxIcon]::Information
                    [windows.forms.messagebox]::Show($message,$title,$button,$icon)
                }
                Function ConvertTo-MultidimensionalArray {
                    [cmdletbinding()]
                    Param (
                        [parameter(ValueFromPipeline)]
                        [object]$InputObject = (Get-Process),
                        [parameter()]
                        [ValidateSet('AliasProperty','CodeProperty','ParameterizedProperty','NoteProperty','Property','ScriptProperty','All')]
                        [string[]]$MemberType = 'Property',
                        [string[]]$PropertyOrder    
                    )
                    Begin {
                        If (-NOT $PSBoundParameters.ContainsKey('Data')) {
                            Write-Verbose 'Pipeline'
                            $isPipeline = $True
                        } Else {
                            Write-Verbose 'Not Pipeline'
                            $isPipeline = $False
                        }
                        $List = New-Object System.Collections.ArrayList
                        $PSBoundParameters.GetEnumerator() | ForEach {
                            Write-Verbose "$($_)"
                        }
                    }
                    Process {
                        If ($isPipeline) {
                            $null = $List.Add($InputObject)
                        }       
                    }
                    End {
                        If ($isPipeline) {
                            $InputObject = $List
                        } 
                        $rowCount = $InputObject.count
                        If ($PSBoundParameters.ContainsKey('PropertyOrder')){
                            $columns = $PropertyOrder
                        }
                        Else {
                            $columns = $InputObject | Get-Member -MemberType $MemberType | Select -Expand Name
                        }
        
                        $columnCount = $columns.count

                        ##Create data holder
                        $MultiArray = New-Object -TypeName 'string[,]' -ArgumentList ($rowCount+1),$columnCount

                        ##Add information to object
                        #Columns first
                        $col=0
                        $columns | ForEach {
                            $MultiArray[0,$col++] = $_
                        }
                        $col=0
                        $row=1
                        For ($i=0;$i -lt $rowCount;$i++) {
                            $columns | ForEach {
                                $MultiArray[$row,$col++] = $InputObject[$i].$_
                            }
                            $row++
                            $col=0
                        }
                        ,$MultiArray
                    }
                }        
                  
                $uiHash.Window.Dispatcher.Invoke("Background",[action]{ 
                    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    $UIHash.ExportToExcel.IsEnabled = $False
                    $UIHash.Filter_btn.IsEnabled = $False
                    $UIHash.ClearFilter_btn.IsEnabled = $False                   
                })
                #$uiHash.host.UI.WriteVerboseLine("Generating report for $($Script:ExcelReports -join '; ')")

                #Create excel COM object
                $excel = New-Object -ComObject excel.application
                $workbook = $excel.Workbooks.Add()

                #Make Visible
                $excel.Visible = $False
                $excel.DisplayAlerts = $False
                $ToCreate = $Script:ExcelReports.Count - 3
                #$uiHash.host.UI.WriteVerboseLine("ToCreate: $ToCreate")
                If ($ToCreate -gt 0) {
                    1..$ToCreate | ForEach {
                        [void]$workbook.Worksheets.Add()
                    }
                } 
                ElseIf ($ToCreate -lt 0) {
                    1..([math]::Abs($ToCreate)) | ForEach {
                        Try {
                            $Workbook.worksheets.item(2).Delete()
                        }
                        Catch {}
                    }
                }
                $i = 1
                ForEach ($Table in $Script:ExcelReports) {
                    #$uiHash.host.UI.WriteVerboseLine("Processing $Table")
                    $uiHash.Window.Dispatcher.Invoke("Background",[action]{ 
                        $uiHash.status_txt.text = "Generating Report: $(($Table).SubString(2))"  
                        $UIHash.ProgressBar.Value++               
                    })
                    $uiHash.Window.Dispatcher.Invoke("Background",[action]{ 
                        $Global:DataGrid = $UIHash."$($Table)_Datagrid"   
                    })
                    $Properties = $DataGrid.ItemsSource[0].psobject.properties|Where{
                        $ExcludeProperties -notcontains $_.Name
                    } | Select-Object -ExpandProperty Name              
                    #$uiHash.host.UI.WriteVerboseLine("Properties: $($Properties -join '; ')")
                    $RowCount = $DataGrid.ItemsSource.Count+1
                    $ColumnCount = ($DataGrid.ItemsSource | Get-Member -MemberType Property).Count
                    #$uiHash.host.UI.WriteVerboseLine("Column Count: $ColumnCount | Row Count: $RowCount" )
                    $uiHash.Window.Dispatcher.Invoke("Background",[action]{ 
                        $Global:__Data = $DataGrid.ItemsSource
                    })                                
                    $Data = $__Data|ConvertTo-MultidimensionalArray -PropertyOrder $Properties 
                    #$uiHash.host.UI.WriteVerboseLine("Data: $($Data|Out-String)")
                    $serverInfoSheet = $workbook.Worksheets.Item($i)
                    [void]$serverInfoSheet.Activate()
                    $serverInfoSheet.Name = $Table.Substring(2)
                    #$uiHash.host.UI.WriteVerboseLine(("Range: {0}, {1}" -f "A1","$($Letters[$ColumnCount])$($RowCount)"))
                    $Range = $serverInfoSheet.Range("A1","$($Letters[$ColumnCount])$($RowCount)")
                    $Range.Value2 = $Data
                    $UsedRange = $serverInfoSheet.UsedRange
                    $UsedRange.Value2 = $UsedRange.Value2
                    [void]$workbook.ActiveSheet.ListObjects.add( 1,$workbook.ActiveSheet.UsedRange,0,1)                                                     														
                    [void]$usedRange.EntireColumn.AutoFit() 
                    [void]$usedRange.EntireRow.AutoFit()
                    $i++
                }

                #Save the report
                Write-Verbose "Saving to $($Script:ReportPath)"
                $workbook.SaveAs(($Script:ReportPath) -f $pwd)

                #Quit the application
                $excel.Quit()

                #Release COM Object
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()

                Show-ReportLocation -ReportLocation $Script:ReportPath

                $uiHash.Window.Dispatcher.Invoke("Background",[action]{ 
                    [System.Windows.Input.Mouse]::OverrideCursor = $Null
                    $UIHash.ExportToExcel.IsEnabled = $True
                    $UIHash.Filter_btn.IsEnabled = $True
                    $UIHash.ClearFilter_btn.IsEnabled = $True  
                    $uiHash.status_txt.text = $Null 
                    $UIHash.ProgressBar.Value = 0                 
                })       
            })
            $PowerShell.Runspace = $newRunspace
            [void]$Jobs.Add((
                [pscustomobject]@{
                    PowerShell = $PowerShell
                    Runspace = $PowerShell.BeginInvoke()
                }
            ))
        }    
    }
    Function Set-Filter {
        If (-NOT [string]::IsNullOrEmpty($UIHash.Filter_txtbx.Text)) {            
            If ($Script:TreeItem.Source.Header -eq 'All') {
                $Filter = $UIHash.Filter_txtbx.Text
            }
            Else {
                $Filter = "computername = '$($Script:TreeItem.Source.Header)' AND ($($UIHash.Filter_txtbx.Text))"
            }  
            Try {
                Write-Verbose "Applying filter <$Filter> on $($Script:TreeItem.Source.Header)"
                $__Filter = [regex]::Replace($Filter,"(\w+) (LIKE)",'CONVERT($1,System.String) $2')
                Write-Verbose "Converted Filter: $($__Filter)"
                $UIHash."$($Script:TabName)_Datagrid".ItemsSource.RowFilter = $__Filter
                Write-Verbose "Saving filter to hashtable" 
                $Script:Filters[$Script:TabName] = $UIHash.Filter_txtbx.Text               
            }
            Catch {
                Write-Verbose $_ 
                $UIHash.Filter_txtbx.ToolTip = $_.Exception.InnerException.Message
                $UIHash.Filter_txtbx.Background = [System.Windows.Media.Brushes]::Red
            }            
            $UIHash.Count_lbl.Content = "Count: {0}" -f $UIHash."$($Script:TabName)_Datagrid".ItemsSource.Count   
        }
    }
    Function Clear-Filter {
        $UIHash.Filter_txtbx.Text = $Null
        $Script:Filters[$Script:TabName] = $Null 
        If ($Script:TreeItem.Source.Header -eq 'All') {
            $Filter = $Null            
        }
        Else {
            $Filter = "computername = '$($Script:TreeItem.Source.Header)'"
        }        
        Write-Verbose "Applying filter <$Filter> on $($Script:TreeItem.Source.Header)"     
        $UIHash."$($Script:TabName)_Datagrid".ItemsSource.RowFilter = $Filter   
        $UIHash.Count_lbl.Content = "Count: {0}" -f $UIHash."$($Script:TabName)_Datagrid".ItemsSource.Count  
    }
    Function Show-AboutHelp {
	    $rs=[RunspaceFactory]::CreateRunspace()
	    $rs.ApartmentState = "STA"
	    $rs.ThreadOptions = "ReuseThread"
	    $rs.Open()
	    $ps = [PowerShell]::Create()
	    $ps.Runspace = $rs
        $ps.Runspace.SessionStateProxy.SetVariable("pwd",$pwd)
	    [void]$ps.AddScript({ 

        [xml]$xaml = @"
        <Window
            xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
            xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
            x:Name='AboutWindow' Title='Help' Width = '400' Height = '670' WindowStartupLocation = 'CenterScreen' ShowInTaskbar = 'False'
            ResizeMode="NoResize">         
            <StackPanel>
                    <TextBlock FontWeight = 'Bold' FontSize = '20' TextDecorations="Underline">Keyboard Shortcuts</TextBlock>
                    <Label />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'CTRL+G - Go to specified computer' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'CTRL+R - Generate Excel Report' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'CTRL+E - Exit application' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'CTRL+A - Go to All Computers' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'F1   -   Displays help information' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'F5   -   Applies filter' />
                    <TextBlock FontSize = '14' Padding = '0' Text = 'F8   -   Clears filter' />
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '20' TextDecorations="Underline">Filter Definitions</TextBlock>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "= (EQUAL)" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find exact matches." />
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "LIKE" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find approximate matches using wildcards (* or %)." />
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "NOT" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find opposite of filter. Placed before the Column name in filter" TextWrapping = 'Wrap'/>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "&lt; (LESS THAN)" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find numbers smaller than given value" TextWrapping = 'Wrap'/>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "&gt; (GREATER THAN)" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find numbers larger than given value" TextWrapping = 'Wrap'/>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "&lt;= (LESS THAN OR EQUAL)" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find numbers equal to or smaller than given value" TextWrapping = 'Wrap'/>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "&gt;= (GREATER THAN OR EQUAL)" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find numbers equal to or larger than given value" TextWrapping = 'Wrap'/>
                    <Label />
                    <TextBlock FontWeight = 'Bold' FontSize = '14' Padding = '0' Text = "IS NULL" />
                    <TextBlock FontSize = '14' Padding = '0' Text = "Used to find columns with no data" TextWrapping = 'Wrap'/>

                    <Label />
                    <Button x:Name = 'CloseButton' Width = '100'> Close </Button>
            </StackPanel>
        </Window>
"@
        #Load XAML
        $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        $AboutWindow=[Windows.Markup.XamlReader]::Load( $reader )


        #Connect to Controls
        $CloseButton = $AboutWindow.FindName("CloseButton")
        $AuthorLink = $AboutWindow.FindName("AuthorLink")


        $CloseButton.Add_Click({
            $AboutWindow.Close()
        })

        #Show Window
        [void]$AboutWindow.showDialog()
        }).BeginInvoke()
    }
    Filter TreeFilter {
        Param($Computername)
        If ($_.Header -eq $Computername) {
            $_
        }
    }
    Function Select-Computer {
        $Computername = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a computername to go to.", "Go To Server")
        If (-Not [System.String]::IsNullOrEmpty($Computername)) {
            #Perform action
            $Computername = $Computername.Trim()
            $Tree = $UIHash.treeView.Items.Items | TreeFilter -Computername $Computername
            $Tree.Focus()
        } 
    }
    Function Invoke-SQLCmd {    
        [cmdletbinding(
            DefaultParameterSetName = 'NoCred',
            SupportsShouldProcess = $True,
            ConfirmImpact = 'Low'
        )]
        Param (
            [parameter()]
            [string]$Computername = 'S46',
        
            [parameter()]
            [string]$Database = 'Master',    
        
            [parameter()]
            [string]$TSQL,

            [parameter()]
            [int]$ConnectionTimeout = 30,

            [parameter()]
            [int]$QueryTimeout = 120,

            [parameter()]
            [System.Collections.ICollection]$SQLParameter,

            [parameter(ParameterSetName='Cred')]
            [Alias('RunAs')]        
            [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,

            [parameter()]
            [ValidateSet('Query','NonQuery')]
            [string]$CommandType = 'Query'
        )
        If ($PSBoundParameters.ContainsKey('Debug')) {
            $DebugPreference = 'Continue'
        }
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Debug $_
        }
        #region Make Connection
        Write-Verbose "Building connection string"
        $Connection=new-object System.Data.SqlClient.SQLConnection 
        Switch ($PSCmdlet.ParameterSetName) {
            'Cred' {
                $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,
                                                                                            $Database,$Credential.Username,
                                                                                            $Credential.GetNetworkCredential().password,$ConnectionTimeout   
                Remove-Variable Credential
            }
            'NoCred' {
                $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout                 
            }
        }   
        $Connection.ConnectionString=$ConnectionString
        Write-Verbose "Opening connection to $($Computername)"
        $Connection.Open()
        #endregion Make Connection

        #region Initiate Query
        Write-Verbose "Initiating query"
        $Command=new-object system.Data.SqlClient.SqlCommand($Tsql,$Connection)
        If ($PSBoundParameters.ContainsKey('SQLParameter')) {
            $SqlParameter.GetEnumerator() | ForEach {
                Write-Debug "Adding SQL Parameter: $($_.Key) with Value: $($_.Value)"
                If ($_.Value -ne $null) { 
                    [void]$Command.Parameters.AddWithValue($_.Key, $_.Value) 
                }
                Else { 
                    [void]$Command.Parameters.AddWithValue($_.Key, [DBNull]::Value) 
                }
            }
        }
        $Command.CommandTimeout=$QueryTimeout
        If ($PSCmdlet.ShouldProcess("Computername: $($Computername) - Database: $($Database)",'Run TSQL operation')) {
            Switch ($CommandType) {
                'Query' {
                    Write-Verbose "Performing Query operation"
                    $DataSet=New-Object system.Data.DataSet
                    $DataAdapter=New-Object system.Data.SqlClient.SqlDataAdapter($Command)
                    [void]$DataAdapter.fill($DataSet)
                    $DataSet.Tables
                }
                'NonQuery' {
                    Write-Verbose "Performing Non-Query operation"
                    [void]$Command.ExecuteNonQuery()
                }
            }
        }
        #endregion Initiate Query    

        #region Close connection
        Write-Verbose "Closing connection"
        $Connection.Close()        
        #endregion Close connection
    }
    Function New-TreeItem {
        Param($MainTree,$Computername)
        $TreeItem = New-Object System.Windows.Controls.TreeViewItem
        $TreeItem.Header = $Computername
        [void]$UIHash.All_trvw.Items.Add($TreeItem)
    }
    Function New-Tab {
        Param($Name, $Header)
        $Tab = New-Object System.Windows.Controls.TabItem
        $Tab.Name = $Name
        $Tab.Header = $Header
        $Tab
    }
    Function New-DataGrid {
        Param($Name, $AlternatingRowColor='LightBlue', $AlternatingRowCount=2)
        $DataGrid = New-Object System.Windows.Controls.DataGrid
        $DataGrid.Name = $Name
        $DataGrid.AlternatingRowBackground = $AlternatingRowColor
        $DataGrid.AlternationCount = $AlternatingRowCount
        $DataGrid.IsReadOnly = $True
        $DataGrid.CanUserAddRows = $False
        $DataGrid.CanUserDeleteRows = $False
        $DataGrid.SelectionMode = 'Single'
        $DataGrid
    }
    Function New-ContextMenu {
        Param()
        $ContextMenu = New-Object System.Windows.Controls.ContextMenu
        $ContextMenu
    }
    Function New-MenuItem {
        Param ($Name, $Header, [System.Windows.Visibility]$Visibility = 'Visible')
        $MenuItem = New-Object System.Windows.Controls.MenuItem
        $MenuItem.Name = $Name
        $MenuItem.Header = $Header
        $MenuItem.Visibility = $Visibility
        $MenuItem
    }
    Function New-Separator {
        Param ($Name, [System.Windows.Visibility]$Visibility = 'Visible')
        $Separator = New-Object System.Windows.Controls.Separator
        $Separator.Name = $Name
        $Separator.Visibility = $Visibility
        $Separator
    }
    #endregion Helper Functions

    #region InvokeSQL Parameters
    $Script:SQLParams = @{
        Computername = $SQLServer
        Database = 'ServerInventory'
        CommandType = 'Query'
        ErrorAction = 'Stop'
        #Credential = $Credential
    }
    #endregion InvokeSQL Parameters
    
    #region Build the GUI
    [xml]$Mainxaml = @"
    <Window Background="White"
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            Title="Server Inventory" Height="800" Width="1260">
        <Grid ShowGridLines="False" Grid.RowSpan="2">
            <Grid.RowDefinitions >
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Menu x:Name="menu" Height="20" Grid.Row="0">
            <Menu.Background>
                <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                    <LinearGradientBrush.GradientStops> 
                        <GradientStop Color='#C4CBD8' Offset='0' /> 
                        <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                        <GradientStop Color='#CFD7E2' Offset='0.9' /> 
                        <GradientStop Color='#C4CBD8' Offset='1' /> 
                    </LinearGradientBrush.GradientStops>
                </LinearGradientBrush>
            </Menu.Background> 
                <MenuItem Header="File">
                    <MenuItem x:Name = 'Exit_menu' Header = 'Exit' ToolTip = 'Exits application' InputGestureText ='Ctrl+E'/>
                </MenuItem>
                <MenuItem Header="Edit">
                    <MenuItem x:Name = 'GoToComputer_menu' Header = 'Go To Computer' ToolTip = 'Goes to specified computer on the treelist' InputGestureText ='Ctrl+G'/>
                </MenuItem>
                <MenuItem Header="Help">
                    <MenuItem x:Name = 'Help_menu' Header = 'Show Help' ToolTip = 'Shows information about application' InputGestureText ='F1'/>
                </MenuItem>
            </Menu>
            <ToolBarTray Height="28" Grid.Row="1">
            <ToolBarTray.Background>
                <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                    <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                    <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
                </LinearGradientBrush>
            </ToolBarTray.Background> 
                <ToolBar x:Name="toolBar1" Height="28" Background = "Transparent">
                    <Button x:Name= "ExportToExcel" Width = '25' Height = "25" Background = "Transparent" ToolTip = "Generate an Excel report" 
                    BorderThickness="0">
                        <Image x:Name = "Excel" />
                    </Button>
                </ToolBar>
                <ToolBar x:Name="toolBar3" Height="28" Background = "Transparent">
                    <Grid>
                    <TextBox x:Name = "Filter_txtbx"  ToolTip = "Type in a query to filter display" Width = "500"/>
                    <TextBlock IsHitTestVisible="False" Text="Enter Filter Here" Foreground="DarkGray"  VerticalAlignment="Center" HorizontalAlignment="Left" 
                    Margin="5,0,0,0">
                        <TextBlock.Style>
                            <Style TargetType="{x:Type TextBlock}">
                                <Setter Property="Visibility" Value="Collapsed"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding Text, ElementName=Filter_txtbx}" Value="">
                                        <Setter Property="Visibility" Value="Visible"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </TextBlock.Style>
                    </TextBlock>
                    </Grid>
                    <Button x:Name = "Filter_btn" Width = '25' Height = "25" ToolTip = "Apply filter to current tab" Background = "Transparent"
                    BorderThickness="0">
                        <Image x:Name = "ApplyFilter" />
                    </Button>
                    <Label Width = "3" Background = "Transparent" />
                    <Button x:Name = "ClearFilter_btn" Width = '25' Height = "25" ToolTip = "Clear filter on current tab" Background = "Transparent"
                    BorderThickness="0">
                        <Image x:Name = "Clear" />
                    </Button>
                </ToolBar>
            </ToolBarTray>
            <Grid x:Name = "SomeGrid" Grid.Row="2" ShowGridLines="False">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="200"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TreeView x:Name="treeView" Grid.Column="0" >
                    <TreeViewItem x:Name = "All_trvw" Header='All' IsExpanded = "True" />
                </TreeView>
                <GridSplitter x:Name="gridSplitter" Grid.Column="1" Width="3" Background = "Black" ResizeBehavior = "PreviousAndNext"/>
                <TabControl x:Name="tabControl" Grid.Column="2" />
            </Grid>
            <Grid ShowGridLines="False"  Grid.Row="3" Height="30">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>                             
                <ProgressBar x:Name = "ProgressBar" Grid.Column = "0"/>                
                <Viewbox Grid.Column = "0">
                    <TextBlock x:Name = "status_txt"/>
                </Viewbox> 
                <Separator Grid.Column = "1" Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Background = "DarkGray"/> 
                <Label x:Name = "Count_lbl" Content = 'Count: ' Grid.Column = "2" Width = "100"/>
            </Grid>                
        </Grid>
    </Window>
"@
 
    Try {
        $reader=(New-Object System.Xml.XmlNodeReader $Mainxaml)
        $UIHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    }
    Catch {
        Write-Warning $_
        BREAK
    }
    #endregion Build the GUI

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)          
    $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
    $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
    $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
        #Routine to handle completed runspaces
        Do {    
            Foreach($runspace in $jobs) {            
                If ($runspace.Runspace.isCompleted) {
                    $runspace.powershell.EndInvoke($runspace.Runspace) | Out-Null
                    $runspace.powershell.dispose()
                    $runspace.Runspace = $null
                    $runspace.powershell = $null               
                } 
            }
            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {
                $_.runspace -eq $Null
            } | ForEach {
                $jobs.remove($_)
            }        
            Start-Sleep -Seconds 1     
        } while ($jobCleanup.Flag)
    })
    $jobCleanup.PowerShell.Runspace = $newRunspace
    $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
    #endregion Background runspace to clean up jobs

    #region Connect to Controls    
    Write-Verbose "Connecting to controls"
    $Mainxaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {
        $uiHash[$_.Name] = $UIHash.Window.FindName($_.Name)
    }
    #endregion Connect to Controls

    #region Excel Report Config Build
    $ExcelReport = {
        Write-Verbose "Param: $($Tables | Out-String)"
        $ExcelUIHash = @{}
        [xml]$ExcelConfigxaml = @"
        <Window 
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                Title="Excel Report Selector" SizeToContent="Height" Width="450" ResizeMode="NoResize">
            <Grid ShowGridLines="False">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <GroupBox x:Name="groupBox" Header="Reports To Include" Grid.Row="0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="5"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="10"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <WrapPanel Grid.Row="1">
                            <CheckBox x:Name = "SelectAll_chkbx" Content="Select All"/>
                        </WrapPanel>
                        <WrapPanel x:Name = "ReportsPanel" Grid.Row="3" ItemWidth="125">
                        </WrapPanel>
                    </Grid>
                </GroupBox>
                <GroupBox Grid.Row = "1" Header = "Report File Path">
                    <TextBox x:Name = "ExcelReportPath_txtbx" Text = "$PWD\ServerInventory.xlsx"/>
                </GroupBox>
                <StackPanel FlowDirection="RightToLeft" Grid.Row="2" Orientation="Horizontal">
                    <Button x:Name = "XLCancel_btn" Content="Cancel" Width="50"/>
                    <Label />
                    <Button x:Name = "XLOK_btn" Content="OK" Width="50"/>
                </StackPanel>
            </Grid>
        </Window>
"@
        $reader=(New-Object System.Xml.XmlNodeReader $ExcelConfigxaml)
        $ExcelUIHash.ExcelWindow=[Windows.Markup.XamlReader]::Load( $reader )

        #region Connect to Controls
        Write-Verbose "Connecting to controls"
        $ExcelConfigxaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {
            $ExcelUIHash[$_.Name] = $ExcelUIHash.ExcelWindow.FindName($_.Name)
        }        
        #endregion Connect to Controls

        #region Event Handlers
        $ExcelUIHash.ExcelWindow.Add_Loaded({  
            #TODO Build the checkboxes for each tab  
            Write-Verbose "Creating additional checkboxes"          
            ForEach ($Table in $Tables) {                
                $CheckBox = New-Object System.Windows.Controls.CheckBox
                $Name = $Table.substring(2)
                Write-Verbose "Adding additional checkboxes <$Name>" 
                $CheckBox.Content = $Name
                $CheckBox.Name = $Table
                [void]$ExcelUIHash.ReportsPanel.AddChild($CheckBox)
            }
            $ExcelUIHash.SelectAll_chkbx.IsChecked = $True
        })

        $ExcelUIHash.XLCancel_btn.Add_Click({
            $Script:ExcelReports.Clear()
            $ExcelUIHash.ExcelWindow.Close()
        })

        $ExcelUIHash.XLOK_btn.Add_Click({ 
            If ($ExcelUIHash.ExcelReportPath_txtbx.text.length -ne 0) {
                ForEach ($Item in $ExcelUIHash.ReportsPanel.Children) {
                    If ($Item.IsChecked) {
                        Write-Verbose "Adding $($item.content)" 
                        [void]$Script:ExcelReports.Add($Item.Name)
                    }
                }
                $Script:ReportPath = $ExcelUIHash.ExcelReportPath_txtbx.text
                $ExcelUIHash.ExcelWindow.DialogResult = $True     
                $ExcelUIHash.ExcelWindow.Close()                
            }
            Else {
                $ExcelUIHash.ExcelReportPath_txtbx.Background = [System.Windows.Media.Brushes]::Red
            }

        })   
        
        $ExcelUIHash.ExcelReportPath_txtbx.Add_TextChanged({
            $ExcelUIHash.ExcelReportPath_txtbx.Background = [System.Windows.Media.Brushes]::White
        })

        $ExcelUIHash.SelectAll_chkbx.Add_Checked({
            $ExcelUIHash.ReportsPanel.Children | ForEach {
                $_.IsChecked = $True
            }
        })
        $ExcelUIHash.SelectAll_chkbx.Add_Unchecked({
            $ExcelUIHash.ReportsPanel.Children | ForEach {
                $_.IsChecked = $False
            }
        })
        #endregion Event Handlers
        $ExcelUIHash.ExcelWindow.Icon = $ExcelBitMap
        $ExcelUIHash.ExcelWindow.ShowDialog()     
    }
    #endregion Excel Report Config Build

    #region Event Handlers

    #region Window Loaded event
    $UIHash.Window.Add_Loaded({  
        $UIHash.Excel.Source = $ExcelBitMap       
        $UIHash.Clear.Source = $ClearBitMap   
        $UIHash.ApplyFilter.Source = $FilterBitMap  
        Write-Verbose "Running Window Loaded Event"
        ForEach ($Table in $Tables) {
            $UIHash["$($Table)_tab"] = New-Tab -TabControl $UIHash.TabControl -Name "$($Table)_tab" -Header $Table.Substring(2)
            $UIHash["$($Table)_datagrid"] = New-DataGrid -Tab $Tab -Name "$($Table)_datagrid"
            $UIHash["ParentContextMenu"] = New-ContextMenu
            $UIHash["AddToFilter_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_AND_$($Table)_menu" -Header "Add to AND Filter" -Visibility Collapsed
            $UIHash["AddToFilter_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_OR_$($Table)_menu" -Header "Add to OR Filter" -Visibility Collapsed
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_AND_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_OR_$($Table)_menu"])
            
            #region AddToFilter (EQUAL)
            $UIHash["AddToFilter_Equal_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Equal_$($Table)_menu" -Header "Add to Filter (EQUAL)"
            $UIHash["AddToFilter_Equal_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName = '$Value'"                
            })
            #endregion AddToFilter (EQUAL)

            #region AddToFilter (NOT EQUAL)
            $UIHash["AddToFilter_NotEqual_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotEqual_$($Table)_menu" -Header "Add to Filter (NOT EQUAL)"
            $UIHash["AddToFilter_NotEqual_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "NOT $ColumnName = '$Value'"                
            })
            #endregion AddToFilter (NOT EQUAL)

            #region AddToFilter (LIKE)
            $UIHash["AddToFilter_Like_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Like_$($Table)_menu" -Header "Add to Filter (LIKE)"
            $UIHash["AddToFilter_Like_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToFilter (LIKE)

            #region AddToFilter (NOT LIKE)
            $UIHash["AddToFilter_NotLike_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotLike_$($Table)_menu" -Header "Add to Filter (NOT LIKE)"
            $UIHash["AddToFilter_NotLike_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "NOT $ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToFilter (NOT LIKE)

            #region AddToFilter (IS NULL)
            $UIHash["AddToFilter_IsNull_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNull_$($Table)_menu" -Header "Add to Filter (IS NULL)"
            $UIHash["AddToFilter_IsNull_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName IS NULL"                
            })
            #endregion AddToFilter (IS NULL)

            #region AddToFilter (IS NOT NULL)
            $UIHash["AddToFilter_IsNotNull_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNotNull_$($Table)_menu" -Header "Add to Filter (IS NOT NULL)"
            $UIHash["AddToFilter_IsNotNull_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName IS NOT NULL"                
            })
            #endregion AddToFilter (IS NOT NULL)

            #region AddToFilter (GREATER THAN)
            $UIHash["AddToFilter_GreaterThan_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThan_$($Table)_menu" -Header "Add to Filter (GREATER THAN)"
            $UIHash["AddToFilter_GreaterThan_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName > '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN)

            #region AddToFilter (GREATER THAN OR EQUAL)
            $UIHash["AddToFilter_GreaterThanOrEqual_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThanOrEqual_$($Table)_menu" -Header "Add to Filter (GREATER THAN OR EQUAL)"
            $UIHash["AddToFilter_GreaterThanOrEqual_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName >= '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN OR EQUAL)

            #region AddToFilter (LESS THAN)
            $UIHash["AddToFilter_LessThan_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThan_$($Table)_menu" -Header "Add to Filter (LESS THAN)"
            $UIHash["AddToFilter_LessThan_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName < '$Value'"                
            })
            #endregion AddToFilter (LESS THAN)

            #region AddToFilter (LESS THAN OR EQUAL)
            $UIHash["AddToFilter_LessThanOrEqual_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThanOrEqual_$($Table)_menu" -Header "Add to Filter (LESS THAN OR EQUAL)"
            $UIHash["AddToFilter_LessThanOrEqual_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$ColumnName <= '$Value'"                
            })
            #endregion AddToFilter (LESS THAN OR EQUAL)
            
            #region AddToORFilter (EQUAL)
            $UIHash["AddToFilter_Equal_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Equal_OR_$($Table)_menu" -Header "Add to OR Filter (EQUAL)"
            $UIHash["AddToFilter_Equal_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName = '$Value'"                
            })
            #endregion AddToORFilter (EQUAL)

            #region AddToORFilter (NOT EQUAL)
            $UIHash["AddToFilter_NotEqual_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotEqual_OR_$($Table)_menu" -Header "Add to OR Filter (NOT EQUAL)"
            $UIHash["AddToFilter_NotEqual_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR NOT $ColumnName = '$Value'"                
            })
            #endregion AddToORFilter (NOT EQUAL)

            #region AddToORFilter (LIKE)
            $UIHash["AddToFilter_Like_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Like_OR_$($Table)_menu" -Header "Add to OR Filter (LIKE)"
            $UIHash["AddToFilter_Like_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToORFilter (LIKE)

            #region AddToORFilter (NOT LIKE)
            $UIHash["AddToFilter_NotLike_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotLike_OR_$($Table)_menu" -Header "Add to OR Filter (NOT)"
            $UIHash["AddToFilter_NotLike_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR NOT $ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToORFilter (NOT LIKE)

            #region AddToORFilter (IS NULL)
            $UIHash["AddToFilter_IsNull_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNull_OR_$($Table)_menu" -Header "Add to OR Filter (IS NULL)"
            $UIHash["AddToFilter_IsNull_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName IS NULL"                
            })
            #endregion AddToORFilter (IS NULL)

            #region AddToORFilter (IS NOT NULL)
            $UIHash["AddToFilter_IsNotNull_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNotNull_OR_$($Table)_menu" -Header "Add to OR Filter (IS NOT NULL)"
            $UIHash["AddToFilter_IsNotNull_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName IS NOT NULL"                
            })
            #endregion AddToORFilter (IS NOT NULL)

            #region AddToFilter (GREATER THAN)
            $UIHash["AddToFilter_GreaterThan_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThan_OR_$($Table)_menu" -Header "Add to OR Filter (GREATER THAN)"
            $UIHash["AddToFilter_GreaterThan_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName > '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN)

            #region AddToFilter (GREATER THAN OR EQUAL)
            $UIHash["AddToFilter_GreaterThanOrEqual_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThanOrEqual_OR_$($Table)_menu" -Header "Add to OR Filter (GREATER THAN OR EQUAL)"
            $UIHash["AddToFilter_GreaterThanOrEqual_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName >= '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN OR EQUAL)

            #region AddToFilter (LESS THAN)
            $UIHash["AddToFilter_LessThan_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThan_OR_$($Table)_menu" -Header "Add to OR Filter (LESS THAN)"
            $UIHash["AddToFilter_LessThan_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName < '$Value'"                
            })
            #endregion AddToFilter (LESS THAN)

            #region AddToFilter (LESS THAN OR EQUAL)
            $UIHash["AddToFilter_LessThanOrEqual_OR_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThanOrEqual_OR_$($Table)_menu" -Header "Add to OR Filter (LESS THAN OR EQUAL)"
            $UIHash["AddToFilter_LessThanOrEqual_OR_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) OR $ColumnName <= '$Value'"                
            })
            #endregion AddToFilter (LESS THAN OR EQUAL)

            #region AddToANDFilter (EQUAL)
            $UIHash["AddToFilter_Equal_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Equal_AND_$($Table)_menu" -Header "Add to AND Filter (EQUAL)"
            $UIHash["AddToFilter_Equal_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName = '$Value'"                
            })
            #endregion AddToANDFilter (EQUAL)

            #region AddToANDFilter (NOT EQUAL)
            $UIHash["AddToFilter_NotEqual_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotEqual_AND_$($Table)_menu" -Header "Add to AND Filter (NOT EQUAL)"
            $UIHash["AddToFilter_NotEqual_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND NOT $ColumnName = '$Value'"                
            })
            #endregion AddToANDFilter (NOT EQUAL)

            #region AddToANDFilter (LIKE)
            $UIHash["AddToFilter_Like_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_Like_AND_$($Table)_menu" -Header "Add to AND Filter (LIKE)"
            $UIHash["AddToFilter_Like_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToANDFilter (LIKE)

            #region AddToANDFilter (NOT LIKE)
            $UIHash["AddToFilter_NotLike_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_NotLike_AND_$($Table)_menu" -Header "Add to AND Filter (NOT LIKE)"
            $UIHash["AddToFilter_NotLike_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND NOT $ColumnName LIKE '%$Value%'"                
            })
            #endregion AddToANDFilter (NOT LIKE)

            #region AddToANDFilter (IS NULL)
            $UIHash["AddToFilter_IsNull_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNull_AND_$($Table)_menu" -Header "Add to AND Filter (IS NULL)"
            $UIHash["AddToFilter_IsNull_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName IS NULL"                
            })
            #endregion AddToANDFilter (IS NULL)

            #region AddToANDFilter (IS NOT NULL)
            $UIHash["AddToFilter_IsNotNull_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_IsNotNull_AND_$($Table)_menu" -Header "Add to AND Filter (IS NOT NULL)"
            $UIHash["AddToFilter_IsNotNull_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName IS NOT NULL"                
            })
            #endregion AddToANDFilter (IS NOT NULL)

            #region AddToFilter (GREATER THAN)
            $UIHash["AddToFilter_GreaterThan_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThan_AND_$($Table)_menu" -Header "Add to AND Filter (GREATER THAN)"
            $UIHash["AddToFilter_GreaterThan_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName > '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN)

            #region AddToFilter (GREATER THAN OR EQUAL)
            $UIHash["AddToFilter_GreaterThanOrEqual_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_GreaterThanOrEqual_AND_$($Table)_menu" -Header "Add to AND Filter (GREATER THAN OR EQUAL)"
            $UIHash["AddToFilter_GreaterThanOrEqual_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName >= '$Value'"                
            })
            #endregion AddToFilter (GREATER THAN OR EQUAL)

            #region AddToFilter (LESS THAN)
            $UIHash["AddToFilter_LessThan_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThan_AND_$($Table)_menu" -Header "Add to AND Filter (LESS THAN)"
            $UIHash["AddToFilter_LessThan_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName < '$Value'"                
            })
            #endregion AddToFilter (LESS THAN)

            #region AddToFilter (LESS THAN OR EQUAL)
            $UIHash["AddToFilter_LessThanOrEqual_AND_$($Table)_menu"] = New-MenuItem -Name "AddToFilter_LessThanOrEqual_AND_$($Table)_menu" -Header "Add to AND Filter (LESS THAN OR EQUAL)"
            $UIHash["AddToFilter_LessThanOrEqual_AND_$($Table)_menu"].Add_Click({
                $UIHash.Filter_txtbx.Text = "$($UIHash.Filter_txtbx.Text) AND $ColumnName <= '$Value'"                
            })
            #endregion AddToFilter (LESS THAN OR EQUAL)

            #region Add Menus to ContextMenu
            [void]$UIHash.tabControl.AddChild($UIHash["$($Table)_tab"])
            $UIHash["$($Table)_tab"].Content = $UIHash["$($Table)_datagrid"]
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_Equal_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_NotEqual_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_GreaterThan_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_GreaterThanOrEqual_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_LessThan_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_LessThanOrEqual_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_Like_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_NotLike_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_IsNull_$($Table)_menu"])
            [void]$UIHash["ParentContextMenu"].AddChild($UIHash["AddToFilter_IsNotNull_$($Table)_menu"])

            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_Equal_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_NotEqual_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_GreaterThan_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_GreaterThanOrEqual_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_LessThan_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_LessThanOrEqual_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_Like_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_NotLike_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_IsNull_AND_$($Table)_menu"])
            [void]$UIHash["AddToFilter_AND_$($Table)_menu"].AddChild($UIHash["AddToFilter_IsNotNull_AND_$($Table)_menu"])

            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_Equal_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_NotEqual_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_GreaterThan_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_GreaterThanOrEqual_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_LessThan_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_LessThanOrEqual_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_Like_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_NotLike_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_IsNull_OR_$($Table)_menu"])
            [void]$UIHash["AddToFilter_OR_$($Table)_menu"].AddChild($UIHash["AddToFilter_IsNotNull_OR_$($Table)_menu"])
            $UIHash["$($Table)_datagrid"].ContextMenu = $UIHash["ParentContextMenu"]            
            #endregion Add Menus to ContextMenu

            #region SQL Data Gathering
            Write-Verbose "Querying table: $Table"
            $Script:SQLParams.TSQL = "SELECT * FROM $Table"
            $Data = Invoke-SQLCmd @SQLParams
            $UIHash."$($Table)_Datagrid".ItemsSource = $Data.DefaultView
            #endregion SQL Data Gathering
        }

        #region Generate the TreeView List
        Write-Verbose "Generating TreeView"
        $UIHash.tbGeneral_Datagrid.ItemsSource | Select-Object -ExpandProperty Computername | 
        Sort-Object | ForEach {
            New-TreeItem -MainTree $UIHash.All_trvw -Computername $_
        }
        #endregion Generate the TreeView List
        $UIHash.tabControl.Items | ForEach {
            $Script:Filters[$Script:TabName] = $Null
            $Script:TempFilters[$Script:TabName] = $Null
        }
        $UIHash.All_trvw.Focus()
    })
    #endregion Window Loaded event

    #region TreeView event handler
    [System.Windows.RoutedEventHandler]$Global:TreeViewChangeHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.TreeViewItem]) {
            $Script:TreeItem = $_  
            Switch ($TreeItem.Source.Header) {
                'All' {
                    Write-Verbose "Setting view to all systems"
                    ForEach ($Table in $Tables) {
                        Write-Verbose "Set DataGrid <$($Table)_Datagrid> filter to: $($Script:Filters[$Table])"
                        $Filter = $Script:Filters[$Table]  
                        $__Filter = [regex]::Replace($Filter,"(\w+) (LIKE)",'CONVERT($1,System.String) $2')                    
                        $UIHash."$($Table)_Datagrid".ItemsSource.RowFilter = $__Filter
                    }
                }
                Default {
                    Write-Verbose "Setting view to $($_)"
                    ForEach ($Table in $Tables) {
                        If (-NOT [string]::IsNullOrEmpty($Script:Filters[$Table])) {
                            $Filter = "computername = '$($_)' AND ($($Script:Filters[$Table]))"
                            $__Filter = [regex]::Replace($Filter,"(\w+) (LIKE)",'CONVERT($1,System.String) $2')
                        }
                        Else {
                            $Filter = "computername = '$($_)'"
                            $__Filter = [regex]::Replace($Filter,"(\w+) (LIKE)",'CONVERT($1,System.String) $2')
                        }
                        Write-Verbose "Set DataGrid <$($Table)_Datagrid> filter to: $Filter"
                        $UIHash."$($Table)_Datagrid".ItemsSource.RowFilter = $__Filter
                    }
                }
            }
            $UIHash.Count_lbl.Content = "Count: {0}" -f $UIHash."$($Script:TabName)_Datagrid".ItemsSource.Count           
        }
    }
    $uiHash.treeView.AddHandler([System.Windows.Controls.TreeViewItem]::SelectedEvent, $TreeViewChangeHandler)
    #endregion TreeView event handler

    #region TabControl event handler
    [System.Windows.RoutedEventHandler]$Global:TabItemChangeHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.TabControl]) {            
            $Script:SelectedTab = $_.OriginalSource.Items | Where {
                $_.IsSelected
            } 
            Write-Verbose "Current tab: $($Script:SelectedTab.Header)" 
            $Script:TabName = $Script:SelectedTab.Name -replace '(.*)_tab','$1'
            $UIHash.Filter_txtbx.Text = If ($Script:Filters[$Script:TabName]) {
                $Script:Filters[$Script:TabName]
            }
            Else {
                $Script:TempFilters[$Script:TabName]
            }
             
            $UIHash.Count_lbl.Content = "Count: {0}" -f $UIHash."$($Script:TabName)_Datagrid".ItemsSource.Count   
        }
    }
    $uiHash.SomeGrid.AddHandler([System.Windows.Controls.TabControl]::SelectionChangedEvent, $TabItemChangeHandler)
    #endregion TabControl event handler

    #region DataGrid Right Click event handler
    [System.Windows.RoutedEventHandler]$Global:DataGridRightClickHandler = {
        $UIHash.Window.UpdateLayout()
        If ($_.OriginalSource -is [System.Windows.Controls.TextBlock]) {                      
            $Script:Value = $_.OriginalSource.Text
            $Script:ColumnName = $_.OriginalSource.Parent.Column.Header
            Write-Verbose "$($Value) | $($ColumnName)"
            If (-NOT [string]::IsNullOrEmpty($UIHash.Filter_txtbx.Text)) {
                $UIHash.GetEnumerator() | Where {
                    $_.name -match 'AddToFilter'
                } | ForEach {
                    If ($_.Name -match '_(OR|AND)_') {
                        $_.Value.Visibility = 'Visible'
                    }
                    Else {                    
                        $_.Value.Visibility = 'Collapsed'
                    }
                }
            }
            Else {
                $UIHash.GetEnumerator() | Where {
                    $_.name -match 'AddToFilter'
                } | ForEach {
                    If ($_.Name -match '_(OR|AND)_') {
                        $_.Value.Visibility = 'Collapsed'
                    }
                    Else {
                        $_.Value.Visibility = 'Visible'
                    }
                }            
            }
        }
    }
    $uiHash.Window.AddHandler([System.Windows.Controls.DataGrid]::MouseRightButtonDownEvent, $DataGridRightClickHandler)
    #endregion DataGrid Right Click event handler

    #region Filter box text changed event
    $UIHash.Filter_txtbx.Add_TextChanged({
        $UIHash.Filter_txtbx.Background = [System.Windows.Media.Brushes]::White
        $UIHash.Filter_txtbx.ToolTip = 'Type in a query to filter display'
        $Script:TempFilters[$Script:TabName] = $This.Text
    })
    #endregion Filter box text changed event

    #region Apply filter button click
    $UIHash.Filter_btn.Add_Click({
        Set-Filter
    })
    #endregion Apply filter button click

    #region Clear filter button click
    $UIHash.ClearFilter_btn.Add_Click({
        Clear-Filter 
    })
    #endregion Clear filter button click

    #region Export DataGrid to Excel click
    $UIHash.ExportToExcel.Add_Click({
        Invoke-ExcelReport
    })
    #endregion Export DataGrid to Excel click

    #region Exit Menu Close
    $UIHash.Exit_menu.Add_Click({
        $UIHash.Window.Close()
    })
    #endregion Exit Menu Close

    #region Window Close 
    $UIHash.Window.Add_Closed({
        Write-Verbose 'Halt runspace cleanup job processing'
        $jobCleanup.Flag = $False

        #Stop all runspaces
        $jobCleanup.PowerShell.Dispose()  
    
        $UIHash.Clear()
        Remove-Variable UIHash -Scope Script

        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()     
    })
    #endregion Window Close 

    #region Go To Computer Menu
    $UIHash.GoToComputer_menu.Add_Click({
        Select-Computer
    })
    #endregion Go To Computer Menu

    #region Show Help
    $UIHash.Help_menu.Add_Click({
        Show-AboutHelp
    })
    #endregion Show Help

    #region KeyDown Event
    $uiHash.Window.Add_KeyDown({ 
        $key = $_.Key  
        If ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) {
            Switch ($Key) {
            "E" {$This.Close()}
            "G" {Select-Computer}
            "A" {$UIHash.All_trvw.Focus()} 
            "R" {Invoke-ExcelReport}       
            Default {$Null}
            }
        } Else {
            Switch ($Key) {
                "F1" {Show-AboutHelp}
                "F5" {Set-Filter}
                "F8" {Clear-Filter}
            }
        }
    })
    #endregion KeyDown Event

    #endregion Event Handlers

    #region Display the Window
    Write-Verbose "Displaying the Window"
    [void]$uiHash.window.Dispatcher.InvokeAsync{$uiHash.window.ShowDialog()}.Wait()
    #endregion Display the Window
}).BeginInvoke()