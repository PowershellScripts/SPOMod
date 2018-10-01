
#
# Created by Arleta Wanat, 2015 
#
# The following cmdlets are a result of passion and hours of work and research. 
# They are distributed freely and happily to anyone who needs them in a day-to-day administration
# in hope they will make your work easier and allow you to manage your SharePoint Online 
# in ways not possible either through User Interface or Sharepoint Online Management Shell.
#
#
#
# The cmdlets can be used as basis for creating scripts and other solutions.
# If you are using the following code for any of your own works, please acknowledge my contribution.
#
#



function Get-SPOListCount
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32342.sharepoint-online-spomod-get-spolistcount.aspx

  #>
  $ctx.Load($ctx.Web.Lists)
  $ctx.ExecuteQuery()
  
  Write-Host $ctx.Web.Lists.Count
  <#
  $i=0

  foreach( $ll in $ctx.Web.Lists)
  {
            
        $i++

        
        }
  
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty Url($ctx.Web.Url)
        $obj | Add-Member NoteProperty Count($i)
        
        Write-Output $obj
  #>
  
  }




function Get-SPOList
{
  
   param (
        [Parameter(Mandatory=$false,Position=0)]
		[switch]$IncludeAllProperties
		)
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32335.sharepoint-online-spomod-get-spolist.aspx

  #>
  
  
  $ctx.Load($ctx.Web.Lists)
  $ctx.ExecuteQuery()
  Write-Host 
  Write-Host $ctx.Url -BackgroundColor White -ForegroundColor DarkGreen
  foreach( $ll in $ctx.Web.Lists)
  {     
        $ctx.Load($ll.RootFolder)
        $ctx.Load($ll.DefaultView)
        $ctx.Load($ll.Views)
        $ctx.Load($ll.WorkflowAssociations)
        try
        {
        $ctx.ExecuteQuery()
        }
        catch
        {
        }

        if($IncludeAllProperties)
        {
        
        $obj = New-Object PSObject
  $obj | Add-Member NoteProperty Title($ll.Title)
  $obj | Add-Member NoteProperty Created($ll.Created)
  $obj | Add-Member NoteProperty Tag($ll.Tag)
  $obj | Add-Member NoteProperty RootFolder.ServerRelativeUrl($ll.RootFolder.ServerRelativeUrl)
  $obj | Add-Member NoteProperty BaseType($ll.BaseType)
  $obj | Add-Member NoteProperty BaseTemplate($ll.BaseTemplate)
  $obj | Add-Member NoteProperty AllowContenttypes($ll.AllowContenttypes)
  $obj | Add-Member NoteProperty ContentTypesEnabled($ll.ContentTypesEnabled)
  $obj | Add-Member NoteProperty DefaultView.Title($ll.DefaultView.Title)
  $obj | Add-Member NoteProperty Description($ll.Description)
  $obj | Add-Member NoteProperty DocumentTemplateUrl($ll.DocumentTemplateUrl)
  $obj | Add-Member NoteProperty DraftVersionVisibility($ll.DraftVersionVisibility)
  $obj | Add-Member NoteProperty EnableAttachments($ll.EnableAttachments)
  $obj | Add-Member NoteProperty EnableMinorVersions($ll.EnableMinorVersions)
  $obj | Add-Member NoteProperty EnableFolderCreation($ll.EnableFolderCreation)
  $obj | Add-Member NoteProperty EnableVersioning($ll.EnableVersioning)
  $obj | Add-Member NoteProperty EnableModeration($ll.EnableModeration)
  $obj | Add-Member NoteProperty Fields.Count($ll.Fields.Count)
  $obj | Add-Member NoteProperty ForceCheckout($ll.ForceCheckout)
  $obj | Add-Member NoteProperty Hidden($ll.Hidden)
  $obj | Add-Member NoteProperty Id($ll.Id)
  $obj | Add-Member NoteProperty IRMEnabled($ll.IRMEnabled)
  $obj | Add-Member NoteProperty IsApplicationList($ll.IsApplicationList)
  $obj | Add-Member NoteProperty IsCatalog($ll.IsCatalog)
  $obj | Add-Member NoteProperty IsPrivate($ll.IsPrivate)
  $obj | Add-Member NoteProperty IsSiteAssetsLibrary($ll.IsSiteAssetsLibrary)
  $obj | Add-Member NoteProperty ItemCount($ll.ItemCount)
  $obj | Add-Member NoteProperty LastItemDeletedDate($ll.LastItemDeletedDate)
  $obj | Add-Member NoteProperty MultipleDataList($ll.MultipleDataList)
  $obj | Add-Member NoteProperty NoCrawl($ll.NoCrawl)
  $obj | Add-Member NoteProperty OnQuickLaunch($ll.OnQuickLaunch)
  $obj | Add-Member NoteProperty ParentWebUrl($ll.ParentWebUrl)
  $obj | Add-Member NoteProperty TemplateFeatureId($ll.TemplateFeatureId)
  $obj | Add-Member NoteProperty Views.Count($ll.Views.Count)
  $obj | Add-Member NoteProperty WorkflowAssociations.Count($ll.WorkflowAssociations.Count)



        Write-Output $obj

        }
        else
        {

        
       
        
        $obj = New-Object PSObject
  $obj | Add-Member NoteProperty Title($ll.Title)
  $obj | Add-Member NoteProperty Created($ll.Created)
  $obj | Add-Member NoteProperty RootFolder.ServerRelativeUrl($ll.RootFolder.ServerRelativeUrl)
        
        
        Write-Output $obj
        
        
     }  
        
        }
  
        

  
  
  }





function Set-SPOList
{

  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32335.sharepoint-online-spomod-set-spolist.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=1)]
		[bool]$NoCrawl,
[Parameter(Mandatory=$false,Position=2)]
		[string]$Title,
[Parameter(Mandatory=$false,Position=3)]
		[string]$Tag,
[Parameter(Mandatory=$false,Position=5)]
		[bool]$ContentTypesEnabled, 
[Parameter(Mandatory=$false,Position=6)]
		[string]$Description, 
[Parameter(Mandatory=$false,Position=7)]
[ValidateSet(0,1,2)]
		[Int]$DraftVersionVisibility, 
[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableAttachments,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableMinorVersions,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableFolderCreation,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableVersioning,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableModeration,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$ForceCheckout,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$Hidden,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$IRMEnabled,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$IsApplicationList,
[Parameter(Mandatory=$false,Position=8)]
		[bool]$OnQuickLaunch     
		)

$ll=$ctx.Web.Lists.GetByTitle($ListName)
    if($PSBoundParameters.ContainsKey("NoCrawl"))
  {$ll.NoCrawl=$NoCrawl}
  if($PSBoundParameters.ContainsKey("Title"))
  {$ll.Title=$Title}
  if($PSBoundParameters.ContainsKey("Tag"))
  {$ll.Tag=$Tag}
  if($PSBoundParameters.ContainsKey("ContentTypesEnabled"))
  {
  $ll.ContentTypesEnabled=$ContentTypesEnabled
  }
  if($PSBoundParameters.ContainsKey("Description"))
  {
  $ll.Description=$Description
  }
  if($PSBoundParameters.ContainsKey("DraftVersionVisibility"))
  {
  $ll.DraftVersionVisibility=$DraftVersionVisibility
  }
  if($PSBoundParameters.ContainsKey("EnableAttachments"))
  {
  $ll.EnableAttachments=$EnableAttachments
  }
  if($PSBoundParameters.ContainsKey("EnableMinorVersions"))
  {$ll.EnableMinorVersions=$EnableMinorVersions}
  if($PSBoundParameters.ContainsKey("EnableFolderCreation"))
  {$ll.EnableFolderCreation=$EnableFolderCreation}
  if($PSBoundParameters.ContainsKey("EnableVersioning"))
  {$ll.EnableVersioning=$EnableVersioning}
  if($PSBoundParameters.ContainsKey("EnableModeration"))
  {$ll.EnableModeration=$EnableModeration}
    if($PSBoundParameters.ContainsKey("ForceCheckout"))
  {$ll.ForceCheckout=$ForceCheckout}
    if($PSBoundParameters.ContainsKey("Hidden"))
  {$ll.Hidden=$Hidden}
    if($PSBoundParameters.ContainsKey("IRMEnabled"))
  {$ll.IRMEnabled=$IRMEnabled}
    if($PSBoundParameters.ContainsKey("IsApplicationList"))
  {$ll.IsApplicationList=$IsApplicationList}
        if($PSBoundParameters.ContainsKey("OnQuickLaunch"))
  {$ll.OnQuickLaunch=$OnQuickLaunch}

      $ll.Update()
    try
    {

        $ctx.ExecuteQuery()
        Write-Host "Done" -ForegroundColor Green
       }

       catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }


}


function New-SPOList
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32341.sharepoint-online-spomod-new-spolist.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$Title,
        [Parameter(Mandatory=$false,Position=1)]
		[int]$TemplateType=100,
        [Parameter(Mandatory=$false,Position=2)]
		[string]$Description="",
        [Parameter(Mandatory=$false,Position=3)]
		[Int]$DocumentTemplateType,
        [Parameter(Mandatory=$false,Position=4)]
		[GUID]$TemplateFeatureID,
        [Parameter(Mandatory=$false,Position=5)]
		[string]$ListUrl=""
		)

  $ListUrl=$Title
  
  $lci =New-Object Microsoft.SharePoint.Client.ListCreationInformation
  $lci.Description=$Description
  $lci.Title=$Title
  $lci.Templatetype=$TemplateType
  if($PSBoundParameters.ContainsKey("ListUrl"))
  {
  $lci.Url =$ListUrl
  }
  if($PSBoundParameters.ContainsKey("DocumentTemplateType"))
  {
  $lci.DocumentTemplateType=$DocumentTemplateType
  }
  if($PSBoundParameters.ContainsKey("TemplateFeatureID"))
  {
  $lci.TemplateFeatureID=$TemplateFeatureID
  }
  $list = $ctx.Web.Lists.Add($lci)
  $ctx.Load($list)
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host "List " $Title " has been added. "
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

     

}



#
#  The following cmdlets will be retired, as they are already included in Set-SPOList cmdlet
#


function Set-SPOListCheckout
{
  param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$ForceCheckout=$true
		)
 
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.ForceCheckout = $ForceCheckout
    $ll.Update()
    
        $listurl=$null
        if($ctx.Url.EndsWith("/")) {$listurl= $ctx.Url+$ll.Title}
        else {$listurl=$ctx.Url+"/"+$ll.Title}
        try
        {
        #$ErrorActionPreference="Stop"
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}

function Set-SPOListVersioning
{
  param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$Enabled=$true
		)
   
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.EnableVersioning=$Enabled
    $ll.Update()
    
       
        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Set-SPOListMinorVersioning
{
  param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$Enabled=$true
		)
  
  
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.EnableMinorVersions=$Enabled
    $ll.Update()
    

        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Remove-SPOListInheritance
{
  param (
		[Parameter(Mandatory=$true,Position=1)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$KeepPermissions=$true
		)
   
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.BreakRoleInheritance($KeepPermissions, $false)
    $ll.Update()
    

        try     {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {        
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Restore-SPOListInheritance
{
  param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ListName
		)
 
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.ResetRoleInheritance()
    $ll.Update()
    
        try        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Set-SPOListContentTypesEnabled
{
  param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=1)]
		[bool]$Enabled=$true
		)
  
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.ContentTypesEnabled=$Enabled
    $ll.Update()
    
        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Remove-SPOList
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32362.sharepoint-online-spomod-remove-spolist.aspx

  #>

  param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ListName
		)

  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.DeleteObject();
        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
           Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Set-SPOListFolderCreationEnabled
{
  param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=1)]
		[bool]$Enabled=$true
		)
  
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.EnableFolderCreation=$Enabled
    $ll.Update()
    
        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


function Set-SPOListIRMEnabled
{
  param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ListName,
        [Parameter(Mandatory=$false,Position=1)]
		[bool]$Enabled=$true
		)
   
  $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $ll.IrmEnabled=$Enabled
    $ll.Update()

        try
        {
        $ctx.ExecuteQuery() 
        Write-Host "Done!" -ForegroundColor DarkGreen             
        }

        catch [Net.WebException] 
        {
            
            Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
        }
          
  

}


#
#
#
#
#
#
# 
#
#
# View Cmdlets
#
#
#
#
#
#
#
#
#
#
#
#


function Get-SPOListView
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32363.sharepoint-online-spomod-get-spolistview.aspx

  #>



 param (
        [Parameter(ParameterSetName="seta", Mandatory=$true,Position=0)]
		[string]$ListName="",
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[string]$ListGUID="",
        [Parameter(Mandatory=$false,Position=0)]
		[switch]$IncludeAllProperties
		)
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32335.sharepoint-online-spomod-get-spolist.aspx

  #>
  switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  { $list=$ctx.Web.Lists.GetByTitle($ListName); break} 
    "setb"  { $list=$ctx.Web.Lists.GetByID($ListGUID); break} 
    } 
  $ctx.Load($list)
  $ctx.Load($list.Views)
  $ctx.ExecuteQuery()
  
  foreach($vv in $list.Views)
  {
    
    if($IncludeAllProperties){
    $ctx.Load($vv)
    $ctx.Load($vv.ViewFields)
    $ctx.ExecuteQuery()
    $vv | Add-Member NoteProperty List.Title($ListName)

    Write-Output $vv}
    else {Write-Output $vv.Title}
  }
}



function Remove-SPOListView
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32364.sharepoint-online-spomod-remove-spolistview.aspx

  #>

 param (
        [parameter(ParameterSetName="seta",ValueFromPipelineByPropertyName, Mandatory=$true,Position=0)]
        [Alias('List.Title')]
        [string]$ListName,
        [Parameter(ParameterSetName="seta", ValueFromPipelineByPropertyName, Mandatory=$true,Position=0)]
        [Alias('Title')]
		[string]$ViewName,
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[GUID]$ListGUID,
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[GUID]$ViewGUID
		)

Begin{
  }

  Process{
   switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  { 
    $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $vv=$ll.Views.GetByTitle($ViewName); break} 
    "setb"  { 
    $ll=$ctx.Web.Lists.GetByID($ListGUID)
    $vv=$ll.Views.GetByID($ViewGUID); break} 
    }
    $ctx.Load($vv)
    $ctx.ExecuteQuery()
    $vv.DeleteObject()
    Write-Verbose "Deleting the view"
    $ctx.ExecuteQuery
    }
}


function Set-SPOListView
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32365.sharepoint-online-spomod-set-spolistview.aspx

  #>


 param (
        [Parameter(ParameterSetName="seta", Mandatory=$true,Position=0)]
		[string]$ListName="",
        [Parameter(ParameterSetName="seta", Mandatory=$true,Position=0)]
		[string]$ViewName="",
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[string]$ListGUID="",
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[string]$ViewGUID="",
[Parameter( Mandatory=$false)]
		[bool]$Hidden,
[Parameter( Mandatory=$false)]
		[bool]$DefaultView,
[Parameter( Mandatory=$false)]
		[string]$AggregationsStatus,
[Parameter( Mandatory=$false)]
		[string]$Aggregations,
[Parameter( Mandatory=$false)]
		[string]$DefaultViewForContentType,
[Parameter( Mandatory=$false)]
		[bool]$EditorModfied,
[Parameter( Mandatory=$false)]
		[string]$Formats,
[Parameter( Mandatory=$false)]
		[bool]$IncludeRootFolder,
[Parameter( Mandatory=$false)]
		[string]$JSLink,
[Parameter( Mandatory=$false)]
		[bool]$MobileDefaultView,
[Parameter( Mandatory=$false)]
		[bool]$MobileView,
[Parameter( Mandatory=$false)]
		[bool]$Paged,
[Parameter( Mandatory=$false)]
		[bool]$PersonalView,
[Parameter( Mandatory=$false)]
		[bool]$RequiresClientIntegration,
[Parameter( Mandatory=$false)]
		[bool]$Threaded,
[Parameter( Mandatory=$false)]
		[Int]$RowLimit
		)

$ll
  $vv
   switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  { 
    $ll=$ctx.Web.Lists.GetByTitle($ListName)
    $vv=$ll.Views.GetByTitle($ViewName); break} 
    "setb"  { 
    $ll=$ctx.Web.Lists.GetByID($ListGUID)
    $vv=$ll.Views.GetByID($ViewGUID); break} 
    }
    $ctx.Load($vv)
    $ctx.ExecuteQuery()
    if($PSBoundParameters.ContainsKey("AggregationsStatus"))
  {
  $vv.AggregationsStatus=$AggregationsStatus
  }
 if($PSBoundParameters.ContainsKey("DefaultView"))
  {
  $vv.DefaultView=$DefaultView
  }
      if($PSBoundParameters.ContainsKey("Hidden"))
  {
  $vv.Hidden=$Hidden
  }
        if($PSBoundParameters.ContainsKey("Aggregations"))
  {
  $vv.Aggregations=$Aggregations
  }
        
        if($PSBoundParameters.ContainsKey("DefaultViewForContentType"))
  {
  $vv.DefaultViewForContentType=$DefaultViewForContentType
  }
        if($PSBoundParameters.ContainsKey("EditorModfied"))
  {
  $vv.EditorModified=$EditorModfied
  }
        if($PSBoundParameters.ContainsKey("Formats"))
  {
  $vv.Formats=$Formats
  }
        if($PSBoundParameters.ContainsKey("IncludeRootFolder"))
  {
  $vv.IncludeRootFolder=$IncludeRootFolder
  }
        if($PSBoundParameters.ContainsKey("JSLink"))
  {
  $vv.JSLink=$JSLink
  }
        if($PSBoundParameters.ContainsKey("MobileDefaultView"))
  {
  $vv.MobileDefaultView=$MobileDefaultView
  }
        if($PSBoundParameters.ContainsKey("MobileView"))
  {
  $vv.MobileView=$MobileView
  }

        if($PSBoundParameters.ContainsKey("Paged"))
  {
  $vv.Paged=$Paged
  }
        if($PSBoundParameters.ContainsKey("PersonalView"))
  {
  $vv.PersonalView=$PersonalView
  }

        if($PSBoundParameters.ContainsKey("RequiresClientIntegration"))
  {
  $vv.RequiresClientIntegration=$RequiresClientIntegration
  }
        if($PSBoundParameters.ContainsKey("Threaded"))
  {
  $vv.Threaded=$Threaded
  }
          if($PSBoundParameters.ContainsKey("RowLimit"))
  {
  $vv.RowLimit=$RowLimit
  }

    $vv.Update()
    $ctx.ExecuteQuery()
}


function New-SPOListView
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32366.sharepoint-online-spomod-new-spolistview.aspx

  #>



 param (
        [Parameter(ParameterSetName="seta", Mandatory=$true,Position=0)]
		[string]$ListName="",
        [Parameter(ParameterSetName="setb", Mandatory=$true,Position=0)]
		[string]$ListGUID="",
[Parameter(Mandatory=$true)]
		[string]$ViewName="DefaultName",
[Parameter(Mandatory=$false)]
		[string]$ViewQuery,
[Parameter(Mandatory=$false)]
		[string[]]$ViewFields,
[Parameter(Mandatory=$false)]
		[Int]$RowLimit
)

    $Vv = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
    $vv.Title=$ViewName
              if($PSBoundParameters.ContainsKey("viewQuery"))
  {
  $vv.Query=$viewQuery
  }
                if($PSBoundParameters.ContainsKey("RowLimit"))
  {
  $vv.RowLimit=$RowLimit
  }
                if($PSBoundParameters.ContainsKey("ViewFields"))
  {
  $vv.ViewFields=$ViewFields
  }

  $ll
   switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  { 
    $ll=$ctx.Web.Lists.GetByTitle($ListName); break} 
    "setb"  { 
    $ll=$ctx.Web.Lists.GetByID($ListGUID); break} 
    }
    $ctx.Load($ll)
    $ctx.Load($ll.Views)
    $ctx.ExecuteQuery()
$listViewToadd=$ll.Views.Add($vv)
$ctx.Load($listViewToadd)
$ctx.ExecuteQuery()

}
#
#
#
#
#
#
# 
#
#
# Column Cmdlets
#
#
#
#
#
#
#
#
#
#
#
#


function Get-SPOListColumn
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32397.sharepoint-online-spomod-get-spolistcolumn.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$true,Position=1)]
		[string]$FieldTitle

		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  
  $ctx.ExecuteQuery()
  $Field=$List.Fields.GetByInternalNameOrTitle($FieldTitle)
  $ctx.Load($Field)

  try
  {
   $ctx.ExecuteQuery()
   

   $obj = New-Object PSObject
   $obj | Add-Member NoteProperty CanBeDeleted($Field.CanBeDeleted)
   $obj | Add-Member NoteProperty DefaultValue($Field.DefaultValue)
        $obj | Add-Member NoteProperty Description($Field.Description)
        $obj | Add-Member NoteProperty Direction($Field.Direction)
        $obj | Add-Member NoteProperty EnforceUniqueValues($Field.EnforceUniqueValues)
        $obj | Add-Member NoteProperty EntityPropertyName($Field.EntityPropertyName)
        $obj | Add-Member NoteProperty Filterable($Field.Filterable)
        $obj | Add-Member NoteProperty FromBaseType($Field.FromBaseType)
        $obj | Add-Member NoteProperty Group($Field.Group)
        $obj | Add-Member NoteProperty Hidden($Field.Hidden)
        $obj | Add-Member NoteProperty ID($Field.Id)
        $obj | Add-Member NoteProperty Indexed($Field.Indexed)
        $obj | Add-Member NoteProperty InternalName($Field.InternalName)
        $obj | Add-Member NoteProperty JSLink($Field.JSLink)
        $obj | Add-Member NoteProperty ReadOnlyField($Field.ReadOnlyField)
        $obj | Add-Member NoteProperty Required($Field.Required)
        $obj | Add-Member NoteProperty SchemaXML($Field.SchemaXML)
        $obj | Add-Member NoteProperty Scope($Field.Scope)
        $obj | Add-Member NoteProperty Sealed($Field.Sealed)
        $obj | Add-Member NoteProperty StaticName($Field.StaticName)
        $obj | Add-Member NoteProperty Sortable($Field.Sortable)
        $obj | Add-Member NoteProperty Tag($Field.Tag)
        $obj | Add-Member NoteProperty Title($Field.Title)
        $obj | Add-Member NoteProperty FieldType($Field.FieldType)
        $obj | Add-Member NoteProperty TypeAsString($Field.UIVersionLabel)
        $obj | Add-Member NoteProperty TypeDisplayName($Field.UIVersionLabel)
        $obj | Add-Member NoteProperty TypeShortDescription($Field.UIVersionLabel)
        $obj | Add-Member NoteProperty ValidationFormula($Field.UIVersionLabel)
        $obj | Add-Member NoteProperty ValidationMessage($Field.UIVersionLabel)
        

        Write-Output $obj
  }
  catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
 



}





function New-SPOListColumn
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32403.sharepoint-online-spomod-new-spolistcolumn.aspx

  #>


param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$true,Position=1)]
		[string]$FieldDisplayName,
  [Parameter(Mandatory=$true, Position=2)]
        [ValidateSet('AllDayEvent','Attachments','Boolean', 'Calculate', 'Choice', 'Computed', 'ContenttypeID', 'Counter', 'CrossProjectLink', 'Currency', 'DateTime', 'Error', 'File', 'Geolocation', 'GridChoice', 'Guid', 'Integer', 'Invalid', 'Lookup', 'MaxItems', 'ModStat', 'MultiChoice', 'Note', 'Number', 'OutcomeChoice', 'PageSeparator', 'Recurrence', 'Text', 'ThreadIndex', 'Threading', 'Url','User', 'WorkflowEventType', 'WorkflowStatus')]
        [System.String]$FieldType,
[Parameter(Mandatory=$false,Position=3)]
		[string]$Description="",
[Parameter(Mandatory=$false,Position=4)]
		[string]$Required="false",
[Parameter(Mandatory=$false,Position=5)]
		[string]$Group="",
[Parameter(Mandatory=$false,Position=6)]
		[string]$StaticName,
[Parameter(Mandatory=$false,Position=7)]
		[string]$Name,
[Parameter(Mandatory=$false,Position=8)]
		[string]$Version="1",
[Parameter(Mandatory=$false,Position=9)]
		[bool]$AddToDefaultView=$false,
[Parameter(Mandatory=$false,Position=10)]
		[string]$AddToView="",
[Parameter(Mandatory=$false,Position=11)]
		[string]$LookupListGUID="",
[Parameter(Mandatory=$false,Position=12)]
		[string]$LookupField="Title"           
		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()

  if($PSBoundParameters.ContainsKey("StaticName")) {$StaticName=$StaticName}
  else {$StaticName=$FieldDisplayName}
  if($PSBoundParameters.ContainsKey("Name")) {$Name=$Name}
  else {$Name=$FieldDisplayName}

   $FieldOptions=[Microsoft.SharePoint.Client.AddFieldOptions]::AddToAllContentTypes 
   $xml="<Field Type='"+$FieldType+"' Description='"+$Description+"' Required='"+$Required+"' Group='"+$Group+"' StaticName='"+$StaticName+"' Name='"+$Name+"' DisplayName='"+$FieldDisplayName+"' Version='"+$Version+"'></Field>"    
   if($LookupListGUID)
   {$xml=$xml.Replace("></Field>"," List='"+$LookupListGUID+"' ShowField='"+$LookupField+"'></Field>")}
   Write-Host $xml
$List.Fields.AddFieldAsXml($xml,$true,$FieldOptions) 
$List.Update() 
 
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host "Field " $FieldDisplayName " has been added to " $ListTitle
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()-ForegroundColor Red
     }

     if($AddToDefaultView -eq $true)
     {
      $ctx.Load($List.DefaultView)
      $ctx.ExecuteQuery()
      if($List.DefaultView -eq $null){ Write-Verbose "There is no default view set for this list"}
      $DefaultViewFields=$List.DefaultView.ViewFields
      $ctx.Load($DefaultViewFields)
      $ctx.ExecuteQuery()
      $List.DefaultView.ViewFields.Add($Name)
      $List.DefaultView.Update()
      $ctx.ExecuteQuery()
      Write-Verbose "Adding to the default view"
     }

     if($AddToView -ne "")
     {
       $ctx.Load($List.Views)
       $ctx.ExecuteQuery()

       $vv=$List.Views.GetByTitle($AddToView.Trim())
       $ctx.Load($vv)
          $ctx.ExecuteQuery()
          $vv.ViewFields.Add($Name)
          $vv.Update()
          $ctx.ExecuteQuery()
          Write-Verbose "Adding to the view "

     
     }


}






function Set-SPOListColumn
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32398.sharepoint-online-spomod-set-spolistcolumn.aspx

  #>


param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$false,Position=11)]
		[string]$DefaultValue,
        [Parameter(Mandatory=$false,Position=12)]
		[string]$Description="",
        [Parameter(Mandatory=$false,Position=13)]
        [ValidateSet('LTR','RTL','none')]
		[string]$Direction,
        [Parameter(Mandatory=$false,Position=14)]
		[bool]$EnforceUniqueValues,
[Parameter(Mandatory=$false,Position=15)]
		[string]$Group="",
[Parameter(Mandatory=$false,Position=16)]
		[bool]$Hidden,
[Parameter(Mandatory=$false,Position=17)]
		[bool]$Indexed,
[Parameter(Mandatory=$false,Position=18)]
		[string]$JSLink="",
[Parameter(Mandatory=$false,Position=19)]
		[bool]$ReadOnlyField,
[Parameter(Mandatory=$false,Position=110)]
		[bool]$Required,
[Parameter(Mandatory=$false,Position=111)]
		[string]$SchemaXML,
[Parameter(Mandatory=$false,Position=112)]
		[string]$StaticName,
[Parameter(Mandatory=$false,Position=113)]
		[string]$Tag,
[Parameter(Mandatory=$true,Position=1)]
		[string]$FieldTitle
		)


  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()
  $lci=$List.Fields.GetByInternalNameOrTitle($FieldTitle)
   $ctx.ExecuteQuery()
  if($PSBoundParameters.ContainsKey("Description"))
  {
  $lci.Description=$Description
  }
  if($PSBoundParameters.ContainsKey("DefaultValue"))
  {
  $lci.DefaultValue=$DefaultValue
  }

  if($PSBoundParameters.ContainsKey("Direction"))
  {
  $lci.Direction=$Direction
  }
  if($PSBoundParameters.ContainsKey("EnforceUniqueValues"))
  {
  $lci.EnforceUniqueValues=$EnforceUniqueValues
  }
  
  if($PSBoundParameters.ContainsKey("Group"))
  {
  $lci.Group=$Group
  }
  if($PSBoundParameters.ContainsKey("Hidden")){
  $lci.Hidden=$Hidden
  }
  if($PSBoundParameters.ContainsKey("Indexed"))
  {
  $lci.Indexed=$Indexed
  }
  
  if($PSBoundParameters.ContainsKey("JSLink"))
  {
  $lci.JSLink=$JSLink
  }
  if($PSBoundParameters.ContainsKey("ReadOnlyField"))
  {
  $lci.ReadOnlyField=$ReadOnlyField
  }
  if($PSBoundParameters.ContainsKey("Required"))
  {
  $lci.Required=$Required
  }
  if($PSBoundParameters.ContainsKey("SchemaXML"))
  {
  $lci.SchemaXML=$SchemaXML
  }
 
  
  if($PSBoundParameters.ContainsKey("StaticName"))
  {
  $lci.StaticName=$StaticName
  }
 
  if($PSBoundParameters.ContainsKey("Tag"))
  {
  $lci.Tag=$Tag
  }


  $lci.Update()
  $ctx.load($lci)
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host $FieldTitle " has been updated"
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

     



}



function Remove-SPOListColumn
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32475.sharepoint-online-spomod-remove-spolistcolumn.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$false,Position=1)]
		[string]$FieldTitle

		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()
  $Field=$List.Fields.GetByTitle($FieldTitle)
   $ctx.ExecuteQuery()
   $Field.DeleteObject()
   $ctx.ExecuteQuery()

}


function Get-SPOListColumnFieldIsObjectPropertyInstantiated
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$false,Position=1)]
		[string]$FieldTitle,
[Parameter(Mandatory=$false,Position=2)]
		[string]$FieldID,
[Parameter(Mandatory=$false,Position=3)]
		[string]$ObjectPropertyName

		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()
  if($PSBoundParameters.ContainsKey("FieldTitle"))
  {
  $Field=$List.Fields.GetByInternalNameorTitle($FieldTitle)
  }
  if($PSBoundParameters.ContainsKey("FieldID"))
  {
  $Field=$List.Fields.GetById($FieldID)
  }
   $ctx.ExecuteQuery()
   $Field.IsObjectPropertyInstantiated($ObjectPropertyName)
   $ctx.ExecuteQuery()

}



function Get-SPOListColumnFieldIsPropertyAvailable
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$false,Position=1)]
		[string]$FieldTitle,
[Parameter(Mandatory=$false,Position=2)]
		[string]$FieldID,
[Parameter(Mandatory=$false,Position=3)]
		[string]$PropertyName

		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()
  if($PSBoundParameters.ContainsKey("FieldTitle"))
  {
  $Field=$List.Fields.GetByInternalNameorTitle($FieldTitle)
  }
  if($PSBoundParameters.ContainsKey("FieldID"))
  {
  $Field=$List.Fields.GetById($FieldID)
  }
   $ctx.ExecuteQuery()
   $Field.IsPropertyAvailable($PropertyName)
   $ctx.ExecuteQuery()

}



function New-SPOListChoiceColumn
{


<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
[Parameter(Mandatory=$true,Position=1)]
		[string]$FieldDisplayName,
[parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [String[]]
            $ChoiceNames,
            [Parameter(Mandatory=$false,Position=2)]
		[string]$Description="",
[Parameter(Mandatory=$false,Position=3)]
		[string]$Required="false",
[Parameter(Mandatory=$false,Position=4)]
[ValidateSet('Dropdown','RadioButtons')]
		[string]$Format="Dropdown",
[Parameter(Mandatory=$false,Position=5)]
		[string]$Group="",
[Parameter(Mandatory=$true,Position=6)]
		[string]$StaticName,
[Parameter(Mandatory=$true,Position=7)]
		[string]$Name,
[Parameter(Mandatory=$false,Position=8)]
		[string]$Version="1",
[Parameter(Mandatory=$false,Position=9)]
[ValidateSet('MultiChoice')]
		[string]$Type
          
		)

  $List=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.ExecuteQuery()
   $FieldOptions=[Microsoft.SharePoint.Client.AddFieldOptions]::AddToAllContentTypes 
    if($PSBoundParameters.ContainsKey("Type"))
   {
    $xml="<Field Type='MultiChoice' Description='"+$Description+"' Required='"+$Required+"' FillInChoice='FALSE' "
   }
   else
   {
   $xml="<Field Type='Choice' Description='"+$Description+"' Required='"+$Required+"' FillInChoice='FALSE' "
   }
   if($PSBoundParameters.ContainsKey("Format"))
   {
     $xml+="Format='"+$Format+"' "
     }
     
     $xml+="Group='"+$Group+"' StaticName='"+$StaticName+"' Name='"+$Name+"' DisplayName='"+$FieldDisplayName+"' Version='"+$Version+"'>
   <CHOICES>"
     
   foreach($choice in $ChoiceNames)
   {
   $xml+="<CHOICE>"+$choice+"</CHOICE>
   "
   
   }
   
   $xml+="</CHOICES>
   </Field>"
   
   
   Write-Host $xml
$List.Fields.AddFieldAsXml($xml,$true,$FieldOptions) 
$List.Update() 
 
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host "Field " $FieldDisplayName " has been added to " $ListTitle
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString() -ForegroundColor
     }

     



}




function Get-SPOListFields
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


 param (
        [Parameter(Mandatory=$true,Position=3)]
		[string]$ListTitle,
        [Parameter(Mandatory=$false,Position=4)]
		[bool]$IncludeSubsites=$false
		)

  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.Fields)
  $ctx.ExecuteQuery()


  $fieldsArray=@()
  $fieldslist=@()
 foreach ($fiel in $ll.Fields)
 {
  #Write-Host $fiel.Description `t $fiel.EntityPropertyName `t $fiel.Id `t $fiel.InternalName `t $fiel.StaticName `t $fiel.Tag `t $fiel.Title  `t $fiel.TypeDisplayName

  $array=@()
  $array+="InternalName"
    $array+="StaticName"
       $array+="Title"
              $array+="SchemaXML"

  $obj = New-Object PSObject
  $obj | Add-Member NoteProperty $array[0]($fiel.InternalName)
  $obj | Add-Member NoteProperty $array[1]($fiel.StaticName)
  $obj | Add-Member NoteProperty $array[2]($fiel.Title)
  $obj | Add-Member NoteProperty $array[3]($fiel.SchemaXML)
  $fieldsArray+=$obj
  $fieldslist+=$fiel.InternalName
  Write-Output $obj
 }
 

 $ctx.Dispose()
  return $fieldsArray

}



function Get-SPOListItems
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

   param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$false,Position=1)]
		[bool]$IncludeAllProperties=$false,
        [switch]$Recursive
		)
  
  
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.Fields)
  $ctx.ExecuteQuery()
  $i=0
  $NumberOfItemsInTheList=$ll.ItemCount
  $itemki=@()
  $spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
  
  if($Recursive)
    {
        $spqQuery.ViewXml ="<View Scope='RecursiveAll' />";
    }

  if($NumberOfItemsInTheList -gt 5000)
  {
    [decimal]$NoOfRuns=($NumberOfItemsInTheList/5000)
    $NoOfRuns=[math]::Ceiling($NoOfRuns)

    for($WhichRun=0; $WhichRun -lt $NoOfRuns; $WhichRun++)
    {
        $startIndex=$WhichRun*5000
        $endIndex=$startIndex+5000
        if($Recursive)
        {
        $spqQuery.ViewXml="<View Scope='RecursiveAll'><Query><Where><And>"+
		    "<Geq><FieldRef Name='ID'></FieldRef><Value Type='Number'>"+$startIndex+"</Value></Geq>"+
			"<Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>"+$endIndex+"</Value></Lt>"+
		  "</And></Where></Query></View>"
        }
        else
        {
        $spqQuery.ViewXml="<View><Query><Where><And>"+
		    "<Geq><FieldRef Name='ID'></FieldRef><Value Type='Number'>"+$startIndex+"</Value></Geq>"+
			"<Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>"+$endIndex+"</Value></Lt>"+
		  "</And></Where></Query></View>"
        }

    
  #  Write-Host $spqQuery.ViewXml
    $partialItems=$ll.GetItems($spqQuery)
    $ctx.Load($partialItems)
    $ctx.ExecuteQuery()

    foreach($partialItem in $partialItems)
    {
        $itemki+=$partialItem
    }
    }
  }

  else
  {
    $itemki=$ll.GetItems($spqQuery)
    $ctx.Load($itemki)
    $ctx.ExecuteQuery()
  }


 #$spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
# $spqQuery.ViewAttributes = "Scope='Recursive'"

Write-Verbose ("Items are ready. Retrieving their properties. It may take a while. My tests indicate 30min per 25 000 items. "+ $itemki.Count)
   
  $bobo=Get-SPOListFields -ListTitle $ListTitle 




 
 $objArray=@()

  for($j=0;$j -lt $itemki.Count ;$j++)
  {
        
        $obj = New-Object PSObject
        
        if($IncludeAllProperties)
        {

        for($k=0;$k -lt $bobo.Count ; $k++)
        {
          
         # Write-Host $k
         $name=$bobo[$k].InternalName
         $value=$itemki[$j][$name]
        # Write-Host $bobo[$k].SchemaXML
         if($bobo[$k].SchemaXML.Contains('Field Type="Lookup"'))
         {
           Write-Host "Contains lookup"
           $value=$itemki[$j][$name].LookupValue
         }
         if($bobo[$k].SchemaXML.Contains('V3Comments'))
         {
           Write-Host "Contains V3Comments"
           $value=$itemki[$j][$name][0]+$itemki[$j][$name][1]+" bb "+$itemki[$j][$name][2]+$itemki[$j][$name][3]+$itemki[$j][$name][4]
         }
          $obj | Add-Member NoteProperty $name($value) -Force
          
        }

        }
        else
        {
          $obj | Add-Member NoteProperty ID($itemki[$j]["ID"])
          $obj | Add-Member NoteProperty Title($itemki[$j]["Title"])

        }

      #  Write-Host $obj.ID `t $obj.Title
        $objArray+=$obj
    
   
  }

 
  
  return $objArray
  
  
  }



  function Get-SPOListItemVersions
  {
   param([Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
        [int]$ItemID=0,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$IncludeAllProperties=$false)

$ll=$ctx.Web.Lists.GetByTitle($ListTitle)
$item=$ll.GetItemByID($ItemID)
$ctx.Load($item)
$ctx.ExecuteQuery()
Write-Host $item["FileRef"]
  $file =$ctx.Web.GetFileByServerRelativeUrl($item["FileRef"]);
        $ctx.Load($file)
        $ctx.Load($file.Versions)
        try{
        $ctx.ExecuteQuery() }
        catch
        {

        }
  if($file.Versions.Count -eq 0)
  {
   Write-Output "No versions available"
  }
  else{

  foreach($vers in $file.Versions)
  {
    Write-Output $vers
  }
  }
  }
#
#
#
#
#
#
# 
#
#
# Item Cmdlets
#
#
#
#
#
#
#
#
#
#
#
#












function New-SPOListItem
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$ItemTitle,
[Parameter(Mandatory=$false,Position=2)]
		[string]$FolderUrl,
[Parameter(Mandatory=$false,Position=3)]
		$AdditionalMultipleFields="",
[Parameter(Mandatory=$false,Position=4)]
		[string]$AdditionalValue="",
[Parameter(Mandatory=$false)]
		[string]$AdditionalField=""
		)


  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()

  $lici =New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
  $lici.FolderUrl=$FolderUrl
  
  $listItem = $ll.AddItem($lici)
  $listItem["Title"]=$ItemTitle
  if($AdditionalField -ne "")
  {
   $listItem[$AdditionalField]=$AdditionalValue
  }

  # The following function I owe to Loic Michel. Thanks!
  if($AdditionalMultipleFields -ne "")
  {
   $additionalMultipleFields |%{
		write-verbose "fieldname :  $($_.fieldname), fieldvalue  $($_.fieldvalue)"	
		$listItem[$_.fieldname]=$_.fieldvalue
   }   }
  $listItem.Update()
  $ll.Update()
  
  try
     {      
         $ctx.ExecuteQuery()
         Write-Host "Item " $ItemTitle " has been added to list " $ListTitle
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }


}



function Remove-SPOListItemInheritance
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


   param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[Int]$ItemID,
        [Parameter(Mandatory=$true,Position=2)]
		[bool]$KeepPermissions
		)
  
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()


  $itemek=$ll.GetItemByID($ItemID)
  $ctx.Load($itemek)
  $ctx.ExecuteQuery()
  $itemek.BreakRoleInheritance($KeepPermissions, $false)
  try
  {
  $ctx.ExecuteQuery()
  write-host $itemek.Name " Success"
  }
 catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
  
  }

  <# Deprecated
  function Remove-SPOListItemPermissions
{
  
   param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[Int]$ItemID
		)
  
  
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()


  $itemek=$ll.GetItemByID($ItemID)
  $ctx.Load($itemek)
  $ctx.ExecuteQuery()
  $itemek.BreakRoleInheritance($false, $false)
  try
  {
  $ctx.ExecuteQuery()
  write-host $itemek.Name " Success"
  }
catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
  
  }
  #>

  function Restore-SPOListItemInheritance
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


   param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[Int]$ItemID
		)
  
  
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()


  $itemek=$ll.GetItemByID($ItemID)
  $ctx.Load($itemek)
  $ctx.ExecuteQuery()
  $itemek.ResetRoleInheritance()
  try
  {
  $ctx.ExecuteQuery()
  write-host $itemek.Name " Success"
  }
 catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
  
  }

  function Remove-SPOListItem
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


   param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[Int]$ItemID
		)
  
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()


  $itemek=$ll.GetItemByID($ItemID)
  $ctx.Load($itemek)
  $ctx.ExecuteQuery()
  $itemek.DeleteObject()
  try
  {
  $ctx.ExecuteQuery()
  write-host $itemek.Name " Success"
  }
catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
  
  }




  function Set-SPOListItem
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


   param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=5)]
		[Int]$ItemID,
[Parameter(Mandatory=$true,Position=6)]
		[string]$FieldToUpdate,
[Parameter(Mandatory=$true,Position=7)]
		[string]$ValueToUpdate
		)
  

  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.ExecuteQuery()


  $itemek=$ll.GetItemByID($ItemID)
  $ctx.Load($itemek)
  $ctx.ExecuteQuery()
  $itemek[$FieldToUpdate] = $ValueToUpdate
  $itemek.Update()
  try
  {
  $ctx.ExecuteQuery()
  write-host $itemek.Name " Success"
  }
  catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
  
  
  }







  #
#
#
#
#
#
# 
#
#
# File Cmdlets
#
#
#
#
#
#
#
#
#
#
#
#







  function Set-SPOFileCheckout
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.CheckOut()
  $ctx.Load($file)
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host $file.Name " has been checked out"   -ForegroundColor DarkGreen 
       }
       catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

}



function Approve-SPOFile
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$false,Position=1)]
		[string]$ApprovalComment=""    
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.Approve($ApprovalComment)
  $ctx.Load($file)

  try
  {
  $ctx.ExecuteQuery()        
        

        Write-Host $file.Name " has been approved"  -ForegroundColor DarkGreen 
        }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
}



function Set-SPOFileCheckin
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>


param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$true,Position=5)]
        [ValidateSet('MajorCheckIn','MinorCheckIn','OverwriteCheckIn')]
        [System.String]$CheckInType,
        [Parameter(Mandatory=$false,Position=6)]
		[string]$CheckinComment=""     
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.CheckIn($CheckInComment, $CheckInType)
  $ctx.Load($file)
  try
  {
  $ctx.ExecuteQuery()        
  Write-Host $file.Name " has been checked in"     -ForegroundColor DarkGreen 
  }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }


}




function Copy-SPOFile
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$true,Position=5)]
		[string]$DestinationLibrary,
        [Parameter(Mandatory=$false,Position=6)]
		[bool]$Overwrite=$true,
        [Parameter(Mandatory=$false,Position=7)]
		[string]$NewName=""
    
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

        if($NewName -eq "")
        {
           $NewName=$file.Name

        }

        if($DestinationLibrary.EndsWith("/")){}
        else {$DestinationLibrary=$DestinationLibrary+"/"}

$file.CopyTo($DestinationLibrary+$NewName, $Overwrite)
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host $file.Name " has been copied to" $DestinationLibrary   -ForegroundColor DarkGreen 
       }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
}



function Remove-SPOFile
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.DeleteObject()
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host $file.Name " has been deleted"   -ForegroundColor DarkGreen 
       }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
}




function Deny-SPOFileApproval
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$false,Position=1)]
		[string]$ApprovalComment=""    
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.Deny($ApprovalComment)
  $ctx.Load($file)

  try
  {
  $ctx.ExecuteQuery()        
        

        Write-Host $file.Name " has been denied"  -ForegroundColor DarkGreen 
        }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
}



function Get-SPOFileIsPropertyAvailable
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl, 
[Parameter(Mandatory=$true,Position=1)]
		[string]$propertyName    
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  if($file.IsPropertyAvailable($propertyName))
  {
  Write-Host "True"
  }
  else
  {
  Write-Host "False"
  }
  

}


function Move-SPOFile
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$DestinationLibrary,
        [Parameter(Mandatory=$false,Position=2)]
		[bool]$Overwrite=$false,
        [Parameter(Mandatory=$false,Position=3)]
		[string]$NewName=""     
		)



  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

        if($PSBoundParameters.ContainsKey("NewName"))
        {
         $DestinationLibrary+=$NewName

        }
        else
        {
        $DestinationLibrary+=$file.Name

        }

        if($PSBoundParameters.ContainsKey("Overwrite"))
        {

  $file.MoveTo($DestinationLibrary,"Overwrite")
  }
  else
  {
  $file.MoveTo($DestinationLibrary,"none")
  }
  
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host $file.Name " has been moved to "  $DestinationLibrary -ForegroundColor DarkGreen 
       }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

}



function Publish-SPOFile
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$false,Position=5)]
		[string]$Comment=""    
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.Publish($Comment)
  $ctx.Load($file)

  try
  {
  $ctx.ExecuteQuery()        
  Write-Host $file.Name " has been published"  -ForegroundColor DarkGreen 
        }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }
}



function Undo-SPOFileCheckout
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)

  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.UndoCheckOut()
  $ctx.Load($file)
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host "Checkout for " $file.Name " has been undone"   -ForegroundColor DarkGreen 
       }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

}


function Undo-SPOFilePublish
{
<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl,
        [Parameter(Mandatory=$false,Position=1)]
		[string]$Comment     
		)


  $file =
        $ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()

  $file.Unpublish($Comment)
  $ctx.Load($file)
  try
  {
  $ctx.ExecuteQuery()        
        
       Write-Host $file.Name " has been unpublished"   -ForegroundColor DarkGreen 
       }
        catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

}



function Get-SPOFolderFilesCount
{
<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)


  $fileCollection =
        $ctx.Web.GetFolderByServerRelativeUrl($ServerRelativeUrl).Files;
        $ctx.Load($fileCollection)
        $ctx.ExecuteQuery()

        
        return $fileCollection.Count

}




function Get-SPOFolderFiles
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)



  $fileCollection =
        $ctx.Web.GetFolderByServerRelativeUrl($ServerRelativeUrl).Files;
        $ctx.Load($fileCollection)
        $ctx.ExecuteQuery()

        
        foreach ($file in $fileCollection)
        {

        $ctx.Load($file.ListItemAllFields)
        $Author=$file.Author
        $CheckedOutByUser=$file.CheckedOutByUser
        $LockedByUser=$file.LockedByUser
        $ModifiedBy=$file.ModifiedBy
        $ctx.Load($Author)
        $ctx.Load($CheckedOutByUser)
        $ctx.Load($LockedByUser)
        $ctx.Load($ModifiedBy)
        $ctx.ExecuteQuery()
        
        
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty Name($file.Name)
        $obj | Add-Member NoteProperty Author.LoginName($file.Author.LoginName)
        $obj | Add-Member NoteProperty CheckedOutByUser.LoginName($file.CheckedOutByUser.LoginName)
        $obj | Add-Member NoteProperty CheckinComment($file.CheckinComment)
        $obj | Add-Member NoteProperty ContentTag($file.ContentTag)
        $obj | Add-Member NoteProperty ETag($file.ETag)
        $obj | Add-Member NoteProperty Exists($file.Exists)
        $obj | Add-Member NoteProperty Length($file.Length)
        $obj | Add-Member NoteProperty LockedByUser.LoginName($file.LockedByUser.LoginName)
        $obj | Add-Member NoteProperty MajorVersion($file.MajorVersion)
        $obj | Add-Member NoteProperty MinorVersion($file.MinorVersion)
        $obj | Add-Member NoteProperty ModifiedBy.LoginName($file.ModifiedBy.LoginName)
        $obj | Add-Member NoteProperty ServerRelativeUrl($file.ServerRelativeUrl)
        $obj | Add-Member NoteProperty Tag($file.Tag)
        $obj | Add-Member NoteProperty TimeCreated($file.TimeCreated)
        $obj | Add-Member NoteProperty TimeLastModified($file.TimeLastModified)
        $obj | Add-Member NoteProperty Title($file.Title)
        $obj | Add-Member NoteProperty UIVersion($file.UIVersion)
        $obj | Add-Member NoteProperty UIVersionLabel($file.UIVersionLabel)
        

        Write-Output $obj
        }



}



function Get-SPOFileByServerRelativeUrl
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=0)]
		[string]$ServerRelativeUrl     
		)


  $file =$ctx.Web.GetFileByServerRelativeUrl($ServerRelativeUrl);
        $ctx.Load($file)
        $ctx.ExecuteQuery()
        $Author=$file.Author
        
        $CheckedOutByUser=$file.CheckedOutByUser
        $LockedByUser=$file.LockedByUser
        $ModifiedBy=$file.ModifiedBy
        $ctx.Load($Author)
        $ctx.Load($CheckedOutByUser)
        $ctx.Load($LockedByUser)
        $ctx.Load($ModifiedBy)
        $ctx.Load($file.EffectiveInformationRightsManagementSettings)
        $ctx.Load($file.InformationRightsManagementSettings)
        $ctx.Load($file.ListItemAllFields)
        $ctx.Load($file.Properties)
        $ctx.Load($file.VersionEvents)
        $ctx.Load($file.Versions)
        $ctx.ExecuteQuery()
<#
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty Name($file.Name)
        $obj | Add-Member NoteProperty Author.LoginName($file.Author.LoginName)
        $obj | Add-Member NoteProperty CheckedOutDate($file.CheckedOutDate)
        $obj | Add-Member NoteProperty CheckedOutByUser.LoginName($file.CheckedOutByUser.LoginName)
        $obj | Add-Member NoteProperty CheckinComment($file.CheckinComment)
        $obj | Add-Member NoteProperty ContentTag($file.ContentTag)
        $obj | Add-Member NoteProperty ETag($file.ETag)
        $obj | Add-Member NoteProperty Exists($file.Exists)
        $obj | Add-Member NoteProperty Length($file.Length)
        $obj | Add-Member NoteProperty LockedByUser.LoginName($file.LockedByUser.LoginName)
        $obj | Add-Member NoteProperty MajorVersion($file.MajorVersion)
        $obj | Add-Member NoteProperty MinorVersion($file.MinorVersion)
        $obj | Add-Member NoteProperty ModifiedBy.LoginName($file.ModifiedBy.LoginName)
        $obj | Add-Member NoteProperty ServerRelativeUrl($file.ServerRelativeUrl)
        $obj | Add-Member NoteProperty Tag($file.Tag)
        $obj | Add-Member NoteProperty TimeCreated($file.TimeCreated)
        $obj | Add-Member NoteProperty TimeLastModified($file.TimeLastModified)
        $obj | Add-Member NoteProperty Title($file.Title)
        $obj | Add-Member NoteProperty UIVersion($file.UIVersion)
        $obj | Add-Member NoteProperty UIVersionLabel($file.UIVersionLabel)
        #>
        Write-Output $file



}








function Get-SPOFolderByServerRelativeUrl
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param (
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ServerRelativeUrl     
		)



  $folderCollection =
        $ctx.Web.GetFolderByServerRelativeUrl($ServerRelativeUrl).Folders;
        $ctx.Load($folderCollection)
        $ctx.ExecuteQuery()


        
        foreach ($fof in $folderCollection)
        {
        $obj = New-Object PSObject
        $ctx.Load($fof.ListItemAllFields)
        $ctx.ExecuteQuery()
        $obj | Add-Member NoteProperty Name($fof.Name)
        $obj | Add-Member NoteProperty Itemcount($fof.ItemCount)
        $obj | Add-Member NoteProperty WelcomePage($fof.WelcomePage)

        Write-Output $obj
        }



}




#
#
#
#
#
#
# 
#
#
# Content Types
#
#
#
#
#
#
#
#
#
#
#
#







function New-SPOListContentType
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param(
[Parameter(Mandatory=$false,Position=4)]
		[string]$Description,
[Parameter(Mandatory=$true,Position=5)]
		[string]$Name,
[Parameter(Mandatory=$false,Position=6)]
		[string]$Group,
[Parameter(ParameterSetName="setb", Mandatory=$true)]
[Parameter(ParameterSetName="seta", Mandatory=$true)]
		[string]$ParentContentTypeID,
[Parameter(ParameterSetName="setc", Mandatory=$true)]
[Parameter(ParameterSetName="setd", Mandatory=$true)]
		[string]$ContentTypeID="",
[Parameter(ParameterSetName="setc", Mandatory=$true)]
[Parameter(ParameterSetName="seta", Mandatory=$true)]
		[string]$ListID,
[Parameter(ParameterSetName="setd", Mandatory=$true)]
[Parameter(ParameterSetName="setb", Mandatory=$true)]
		[string]$ListName=""

		)
    


  $lci =New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
  if($PSBoundParameters.ContainsKey("Description"))
  {$lci.Description=$Description}
   if($PSBoundParameters.ContainsKey("Group"))
  {$lci.Group=$Group}
  $lci.Name=$Name  
  switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  { $lci.ParentContentType=$ctx.Web.ContentTypes.GetById($ParentContentTypeID);
     $ContentType = $ctx.Web.Lists.GetByID($ListID).ContentTypes.Add($lci); break} 
    "setb"  { $lci.ParentContentType=$ctx.Web.ContentTypes.GetById($ParentContentTypeID);
    $ContentType = $ctx.Web.Lists.GetByTitle($ListName).ContentTypes.Add($lci); break} 
    "setc"  { $lci.ID=$ContentTypeID;
    $ContentType = $ctx.Web.Lists.GetByID($ListID).ContentTypes.Add($lci); break} 
    "setd"  { $lci.ID=$ContentTypeID;
    $ContentType = $ctx.Web.Lists.GetByTitle($ListName).ContentTypes.Add($lci); break} 
    } 


  $ctx.Load($contentType)
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host "Content Type " $Name " has been added to the list"
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

     

}




function New-SPOSiteContentType
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

param(
[Parameter(Mandatory=$false,Position=4)]
		[string]$Description,
[Parameter(Mandatory=$true,Position=5)]
		[string]$Name,
[Parameter(Mandatory=$false,Position=6)]
		[string]$Group,
[Parameter(ParameterSetName="seta", Mandatory=$true,Position=7)]
		[string]$ParentContentTypeID,
[Parameter(ParameterSetName="setb", Mandatory=$true,Position=8)]
		[string]$ContentTypeID=""
		)
 

  $lci =New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
  $lci.Name=$Name  
  if($PSBoundParameters.ContainsKey("Description"))
  {$lci.Description=$Description}
   if($PSBoundParameters.ContainsKey("Group"))
  {$lci.Group=$Group}
  switch ($PsCmdlet.ParameterSetName) 
    { 
    "seta"  {$lci.ParentContentType=$ctx.Web.ContentTypes.GetById($ParentContentTypeID); break} 
    "setb"  {$lci.ID=$ContentTypeID; break} 
    } 
  $ContentType = $ctx.Web.ContentTypes.Add($lci)
  $ctx.Load($contentType)
  try
     {
       
         $ctx.ExecuteQuery()
         Write-Host "Content Type " $Name " has been added to the site"
     }
     catch [Net.WebException]
     { 
        Write-Host $_.Exception.ToString()
     }

     

}







function New-SPOListContentTypeColumn
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
   param (

        [Parameter(Mandatory=$true,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$ColumnName,
        [Parameter(Mandatory=$true,Position=2)]
		[string]$ContentTypeID,
        [Parameter(Mandatory=$false,Position=5)]
		[switch]$ListColumn
		)
  

  $ctx.Load($ctx.Web.Lists)
  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
  $ctx.Load($ll)
  $ctx.Load($ll.ContentTypes)
  $ctType=$ll.ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  if($ListColumn)
  {
     $field=$ll.Fields.GetByInternalNameOrTitle($ColumnName)
  }
  else{  $field=$ctx.Web.Fields.GetByInternalNameOrTitle($ColumnName) }

    
     $link=new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
     $link.Field=$field
     $fielsie=$ctType.FieldLinks.Add($link)
     $ctType.Update($false)
     $ctx.ExecuteQuery()
   
  
     
     
} 






function New-SPOSiteContentTypeColumn
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
   param (

        [Parameter(Mandatory=$true,Position=0)]
		[string]$ContentTypeID,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$ColumnName,
        [Parameter(Mandatory=$true,Position=3)]
		[bool]$UpdateChildren
		)
  

  $ctType=$ctx.Web.ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  $field=$ctx.Web.Fields.GetByInternalNameOrTitle($ColumnName)

     $link=new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
     $link.Field=$field
     $fielsie=$ctType.FieldLinks.Add($link)
     $ctType.Update($UpdateChildren)
     $ctx.ExecuteQuery()
   
     
  
     
}




function Get-SPOContentType
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
   param (

        [Parameter(Mandatory=$false,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$false,Position=1)]
		[switch]$Available
		)
  

  if($PSBoundParameters.ContainsKey("ListTitle"))
  {
  $ctTypes=$ctx.Web.Lists.GetByTitle($ListTitle).ContentTypes
  $ctx.Load($ctTypes)
  $ctx.ExecuteQuery()
  }
  elseif($Available)
  {
  $ctTypes=$ctx.Web.AvailableContentTypes
  $ctx.Load($ctTypes)
  $ctx.ExecuteQuery()

  }
  else
  {
  $ctTypes=$ctx.Web.ContentTypes
  $ctx.Load($ctTypes)
  $ctx.ExecuteQuery()
  }



  foreach($cc in $ctTypes)
  {

       $ctx.Load($cc)
          $ctx.Load($cc.FieldLinks)
     $ctx.Load($cc.Fields)
     $ctx.Load($cc.WorkflowAssociations)
     $ctx.ExecuteQuery()
      foreach($field in $cc.Fields)
     {
      $PropertyName="Field "+$field.ID
      $cc | Add-Member NoteProperty $PropertyName($field.Title)
     }
     foreach($fieldlink in $cc.FieldLinks)
     {
      $PropertyName="Fieldlink "+$fieldlink.ID
      $cc | Add-Member NoteProperty $PropertyName($fieldlink.Name)
     }
     foreach($workflow in $cc.WorkflowAssociations)
     {
      $PropertyName="Workflow "+$workflow.ID
      $cc | Add-Member NoteProperty $PropertyName($workflow.Name)
     }

     Write-Output $cc

  }
  
  
   
     
  
     
}


function Remove-SPOContentType
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
   param (

        [Parameter(Mandatory=$false,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$ContentTypeID
		)
  

  if($PSBoundParameters.ContainsKey("ListTitle"))
  {
  $ctType=$ctx.Web.Lists.GetByTitle($ListTitle).ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  $ctType.DeleteObject()
  $ctx.ExecuteQuery()
  }
  else
  {
  $ctType=$ctx.Web.ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  $ctType.DeleteObject()
  $ctx.ExecuteQuery()

  }


  }



function Set-SPOContentType
{
  <#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
   param (

        [Parameter(Mandatory=$false,Position=0)]
		[string]$ListTitle,
        [Parameter(Mandatory=$true,Position=1)]
		[string]$ContentTypeID,
        [Parameter(Mandatory=$false)]
		[string]$Group,
        [Parameter(Mandatory=$false)]
		[string]$DisplayFormUrl,
        [Parameter(Mandatory=$false)]
		[string]$EditFormUrl,
        [Parameter(Mandatory=$false)]
		[bool]$Hidden,
        [Parameter(Mandatory=$false)]
		[string]$JSLink,
        [Parameter(Mandatory=$false)]
		[string]$NewFormUrl,
        [Parameter(Mandatory=$false)]
		[bool]$ReadOnly,
        [Parameter(Mandatory=$false)]
		[bool]$Sealed=$false,
        [Parameter(Mandatory=$true)]
		[bool]$UpdateChildren

		)
  

  if($PSBoundParameters.ContainsKey("ListTitle"))
  {
  $ctType=$ctx.Web.Lists.GetByTitle($ListTitle).ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  }
  else
  {
  $ctType=$ctx.Web.ContentTypes.GetByID($ContentTypeID)
  $ctx.Load($ctType)
  $ctx.ExecuteQuery()
  }

    if($PSBoundParameters.ContainsKey("Group"))
  {$ctType.Group=$Group}
      if($PSBoundParameters.ContainsKey("DisplayFormUrl"))
  {$ctType.DisplayFormUrl=$DisplayFormUrl}
      if($PSBoundParameters.ContainsKey("EditFormUrl"))
  {$ctType.EditFormUrl=$EditFormUrl}
      if($PSBoundParameters.ContainsKey("Hidden"))
  {$ctType.Hidden=$Hidden}
      if($PSBoundParameters.ContainsKey("JSLink"))
  {$ctType.JSLink=$JSLink}
      if($PSBoundParameters.ContainsKey("NewFormUrl"))
  {$ctType.NewFormUrl=$NewFormUrl}
      if($PSBoundParameters.ContainsKey("ReadOnly"))
  {$ctType.ReadOnly=$ReadOnly}
        if($PSBoundParameters.ContainsKey("Sealed"))
  {$ctType.Sealed=$Sealed}

  $ctType.Update($UpdateChildren)
  $ctx.ExecuteQuery()

  }



#
#
#
#
#
#
# 
#
#
# Taxonomy
#
#
#
#
#
#
#
#
#
#
#
#



function New-SPOTerm
{
param (
		#[Parameter(Mandatory=$true,Position=4)]
		#[string]$TermSetGuid,
		[Parameter(Mandatory=$true,Position=5)]
		[string]$Term,
		[Parameter(Mandatory=$true,Position=6)]
		[string]$TermLanguage
		)


  $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

         $termstore = $session.GetDefaultSiteCollectionTermStore();
         $ctx.Load($termstore)
         $ctx.ExecuteQuery()

  Write-Host "Termstore" -ForegroundColor Green
  Write-Host "Term1"
  $set=$termstore.GetTermSet($TermSetGuid)
  $ctx.Load($set)
  $ctx.Load($set.GetAllTerms())
  $ctx.ExecuteQuery()
  $guid = [guid]::NewGuid()
  Write-Host $guid
  $term=$set.CreateTerm($Term, $TermLanguage,$guid)
 
  $termstore.CommitAll()
  
  $ctx.ExecuteQuery()

  }


function Get-SPOTermGroups
{

  $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

         $termstore = $session.GetDefaultSiteCollectionTermStore();
         $ctx.Load($termstore)
         $ctx.ExecuteQuery()

  $groups=$termstore.Groups
  $ctx.Load($groups)

  $ctx.ExecuteQuery()

  foreach($group in $groups)
  {
    $ctx.Load($group)
    $ctx.Load($group.TermSets)
    $ctx.ExecuteQuery()
    Write-Output $group

  }

  }

  
function Get-SPOTermSets
{
param (
		[Parameter(ParameterSetName="groupName",Mandatory=$true)]
		[string]$TermGroupName="",
        [Parameter(ParameterSetName=’groupId’, Mandatory=$false)]
        [string]$TermGroupId=""
		)
  
  
  $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

         $termstore = $session.GetDefaultSiteCollectionTermStore();
         $ctx.Load($termstore)
         $ctx.ExecuteQuery()
         if($TermGroupName -eq "" -and ($TermGroupId -eq ""))
         {
            $groups=$termstore.Groups
            $ctx.Load($groups)
            $ctx.ExecuteQuery()

            foreach($group in $groups)
                {
                    $ctx.Load($group)
                    $ctx.Load($group.TermSets)
                    $ctx.ExecuteQuery()
                    foreach($termset in $group.TermSets)
                    {
                        $ctx.Load($termset)
                        $ctx.Load($termset.Terms)
                        $ctx.ExecuteQuery()
                        Write-Output $termset  

                    }

                }
           }
           else 
           {
            $group;
            if($TermGroupName -ne ""){
                $group=$termstore.Groups.GetByName($TermGroupName)
            }
            elseif($TermGroupId -ne ""){
                $group=$termstore.Groups.GetById($TermGroupId)
            }
            else{
            $group=$termstore.Groups[0]
            }
            $ctx.Load($group)
            $ctx.Load($group.TermSets)
            $ctx.ExecuteQuery()
            foreach($termset in $group.TermSets)
                {
                    $ctx.Load($termset)
                    $ctx.Load($termset.Terms)
                    $ctx.ExecuteQuery()
                    Write-Output $termset  

                }

           }


  }


  function Get-SPOTermStore
  {
        
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        $ctx.Load($termstore)
        $ctx.Load($termstore.Groups)
        $ctx.ExecuteQuery() 

        Write-Output $termstore

  }

 function Get-SPOHashTagsTermSet
  {
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        $ctx.Load($termstore)
        $ctx.Load($termstore.HashTagsTermSet)
        $ctx.Load($termstore.HashTagsTermSet.Terms)
        $ctx.ExecuteQuery()
        Write-Output $termstore.HashTagsTermSet
  }

function  Get-SPOHashTagsTerms
  {
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        $hashtagtermset=$termstore.HashTagsTermSet
        $ctx.Load($termstore)
        $ctx.Load($hashtagtermset)
        $ctx.Load($hashtagtermset.Terms)
        $ctx.ExecuteQuery()
        foreach($term in $hashtagtermset.Terms)
        {
          $ctx.Load($term)
          $ctx.Load($term.Terms)
          $ctx.Load($term.TermSets)
          $ctx.Load($term.Labels)
          $ctx.Load($term.ReusedTerms)
          $ctx.ExecuteQuery()
          Write-Output $term
        }
  }

  function Get-SPOKeyWordsTermSet
  {
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        $keywordsTermStore=$termstore.KeywordsTermSet
        $ctx.Load($termstore)
        $ctx.Load($keywordsTermStore)
        $ctx.Load($keywordsTermStore.Terms)
        $ctx.ExecuteQuery()
        Write-Output $keywordsTermStore
  }



  function New-SPOTermGroup
  {
        param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$false)]
        [string]$GUID=""
        )
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        if($GUID -eq ""){$GUID = [guid]::NewGuid()}
        $group=$termstore.CreateGroup($Name,$GUID)
        try
        {
        $ctx.ExecuteQuery()
        Write-Host "Group " $Name " created successfully."
        
        }
        catch [Net.WebException] 
        {
            
            Write-Host "Couldn't create a group "$Name $_.Exception.ToString() -ForegroundColor Red
        }

  }

  function Set-SPOTermGroup
  {
        param(
        [Parameter(ParameterSetName="ByGUID",Mandatory=$true)]
        [string]$GUID="",
        [Parameter(ParameterSetName="FromPipeline", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        $group=$null,
        [Parameter(Mandatory=$false)]
        [string]$Description
        )
        BEGIN{
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()

        $termstore = $session.GetDefaultSiteCollectionTermStore();
        }
        PROCESS{
                if($group -eq $null)
                    {
                        $group=$termstore.GetGroup($GUID)
                        $ctx.Load($termstore)
                        $ctx.Load($group)
                        $ctx.ExecuteQuery()
                    }
                $group.Description=$Description
                $ctx.ExecuteQuery()
        }
  }

  function New-SPOTermSet
  {
    param(
    [Parameter(ParameterSetName="ByName",Mandatory=$true)]
    [string]$TermGroupName="",
    [Parameter(ParameterSetName="ByID",Mandatory=$true)]
    [string]$TermGroupID="",
    [Parameter(ParameterSetName="ByGroup",Mandatory=$true,ValueFromPipeline=$true)]
    $TermGroup="",
    [Parameter(Mandatory=$true)]
    [string]$TermSetName,
    [Parameter(Mandatory=$false)]
    [int]$LanguageID=1033,
    [Parameter(Mandatory=$false)]
    $GUID=""
    
    )

    BEGIN{
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()
        
        }
    PROCESS{
        $termstore = $session.GetDefaultSiteCollectionTermStore()
        if($TermGroupName -ne "")
        { $group=$termstore.Groups.GetByName($TermGroupName)}
        elseif($TermGroupID -ne "")
        { $group=$termstore.Groups.GetById($TermGroupID)}
        elseif($group -ne "")
        { 
            $group=$TermGroup
        }
        else
        {
          Write-Host "Could not retrieve the group. Missing parameters"
        }
        
        $ctx.Load($group)
        $ctx.Load($group.TermSets)
        $ctx.ExecuteQuery()
        if($GUID -eq ""){$GUID = [guid]::NewGuid()}
        
        $termSet=$group.CreateTermSet($TermSetName, $GUID, $LanguageID)
        $ctx.ExecuteQuery()
        return $termset
        }
        

  }

  function Set-SPOTermSet
  {
    param(
    [Parameter(Mandatory=$true, Position=1)]
    [Microsoft.SharePoint.Client.Taxonomy.TermSet]$TermSet,
    [Parameter(Mandatory=$false)]
    [string]$Description="",
    [Parameter(Mandatory=$false)]
    [bool]$IsOpenForTermCreation,
    [Parameter(Mandatory=$false)]
    [bool]$IsAvailableForTagging,
    [Parameter(Mandatory=$false)]
    [string]$Name,
    [Parameter(Mandatory=$false)]
    [string]$Owner,
    [Parameter(Mandatory=$false)]
    [string]$StakeholderToAdd,
    [Parameter(Mandatory=$false)]
    [string]$StakeholderToRemove
    )
    if($Description -ne "")
    {
        $TermSet.Description=$Description
    }
    if($PSBoundParameters.ContainsKey("IsOpenForTermCreation"))
    {
        $TermSet.IsOpenForTermCreation=$IsOpenForTermCreation
    }
        if($PSBoundParameters.ContainsKey("IsOpenForTermCreation"))
    {
        $TermSet.IsAvailableForTagging=$IsAvailableForTagging
    }
        if($PSBoundParameters.ContainsKey("Name"))
    {
        $TermSet.Name=$Name
    }
            if($PSBoundParameters.ContainsKey("Owner"))
    {
        $TermSet.Owner=$Owner
    }
        if($PSBoundParameters.ContainsKey("StakeholderToAdd"))
    {
        $TermSet.AddStakeholder($StakeholderToAdd)
    }
        if($PSBoundParameters.ContainsKey("StakeholderToRemove"))
    {
        $TermSet.DeleteStakeholder($StakeholderToRemove)
    }

    try
    {
        $ctx.ExecuteQuery()
    }
   catch [Net.WebException] 
   {
      Write-Host "Could not update the termset. " $_.Exception.ToString() -ForegroundColor Red
   }

  }

  function Get-SPOTerm
  {
    param(
    [Parameter(Mandatory=$true, Position=1)]
    [GUID]$Guid
    )
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
        $ctx.Load($session)
        $ctx.ExecuteQuery()
        $termstore = $session.GetDefaultSiteCollectionTermStore();
        $term=$termstore.GetTerm($Guid)
        $ctx.Load($termstore)
        $ctx.Load($term)
        $ctx.Load($term.Terms)
        $ctx.Load($term.TermSets)
        $ctx.ExecuteQuery()
        Write-Output $term

  }



#
#
#
#
#
#
# 
#
#
# Web
#
#
#
#
#
#
#
#
#
#
#
#

function Get-SPOWeb
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>
param (
        [Parameter(Mandatory=$false,Position=4)]
		[bool]$IncludeSubsites=$false
		)


  $ctx.Load($ctx.Web)
  $ctx.Load($ctx.Web.Webs)
  $ctx.ExecuteQuery()
  



# Get the root
        $obj = new-Object PSOBject
        $obj | Add-Member NoteProperty AllowRSSFeeds($ctx.Web.Webs[$i].AllowRssFeeds)
        $obj | Add-Member NoteProperty Created($ctx.Web.Webs[$i].Created)
        $obj | Add-Member NoteProperty CustomMasterUrl($ctx.Web.Webs[$i].CustomMasterUrl)
        $obj | Add-Member NoteProperty Description($ctx.Web.Webs[$i].Description)
        $obj | Add-Member NoteProperty EnableMinimalDownload($ctx.Web.Webs[$i].EnableMinimalDownload)
        $obj | Add-Member NoteProperty ID($ctx.Web.Webs[$i].Id)
        $obj | Add-Member NoteProperty Language($ctx.Web.Webs[$i].Language)
        $obj | Add-Member NoteProperty LastItemModifiedDate($ctx.Web.Webs[$i].LastItemModifiedDate)
        $obj | Add-Member NoteProperty MasterUrl($ctx.Web.Webs[$i].MasterUrl)
        $obj | Add-Member NoteProperty QuickLaunchEnabled($ctx.Web.Webs[$i].QuickLaunchEnabled)
        $obj | Add-Member NoteProperty RecycleBinEnabled($ctx.Web.Webs[$i].RecycleBinEnabled)
        $obj | Add-Member NoteProperty ServerRelativeUrl($ctx.Web.Webs[$i].ServerRelativeUrl)
        $obj | Add-Member NoteProperty Title($ctx.Web.Webs[$i].Title)
        $obj | Add-Member NoteProperty TreeViewEnabled($ctx.Web.Webs[$i].TreeViewEnabled)
        $obj | Add-Member NoteProperty UIVersion($ctx.Web.Webs[$i].UIVersion)
        $obj | Add-Member NoteProperty UIVersionConfigurationEnabled($ctx.Web.Webs[$i].UIVersionConfigurationEnabled)
        $obj | Add-Member NoteProperty Url($ctx.Web.Webs[$i].Url)
        $obj | Add-Member NoteProperty WebTemplate($ctx.Web.Webs[$i].WebTemplate)

        Write-Output $obj

        # Get the subsites
        if($IncludeSubsites){
                if($ctx.Web.Webs.Count -eq 0)
                        {
                        Write-Host "No subsites found" 

                        }
                for($i=0;$i -lt $ctx.Web.Webs.Count ;$i++)
                    {
                    $obj = new-Object PSOBject
                    $obj | Add-Member NoteProperty AllowRSSFeeds($ctx.Web.Webs[$i].AllowRssFeeds)
                    $obj | Add-Member NoteProperty Created($ctx.Web.Webs[$i].Created)
                    $obj | Add-Member NoteProperty CustomMasterUrl($ctx.Web.Webs[$i].CustomMasterUrl)
                    $obj | Add-Member NoteProperty Description($ctx.Web.Webs[$i].Description)
                    $obj | Add-Member NoteProperty EnableMinimalDownload($ctx.Web.Webs[$i].EnableMinimalDownload)
                    $obj | Add-Member NoteProperty ID($ctx.Web.Webs[$i].Id)
                    $obj | Add-Member NoteProperty Language($ctx.Web.Webs[$i].Language)
                    $obj | Add-Member NoteProperty LastItemModifiedDate($ctx.Web.Webs[$i].LastItemModifiedDate)
                    $obj | Add-Member NoteProperty MasterUrl($ctx.Web.Webs[$i].MasterUrl)
                    $obj | Add-Member NoteProperty QuickLaunchEnabled($ctx.Web.Webs[$i].QuickLaunchEnabled)
                    $obj | Add-Member NoteProperty RecycleBinEnabled($ctx.Web.Webs[$i].RecycleBinEnabled)
                    $obj | Add-Member NoteProperty ServerRelativeUrl($ctx.Web.Webs[$i].ServerRelativeUrl)
                    $obj | Add-Member NoteProperty Title($ctx.Web.Webs[$i].Title)
                    $obj | Add-Member NoteProperty TreeViewEnabled($ctx.Web.Webs[$i].TreeViewEnabled)
                    $obj | Add-Member NoteProperty UIVersion($ctx.Web.Webs[$i].UIVersion)
                    $obj | Add-Member NoteProperty UIVersionConfigurationEnabled($ctx.Web.Webs[$i].UIVersionConfigurationEnabled)
                    $obj | Add-Member NoteProperty Url($ctx.Web.Webs[$i].Url)
                    $obj | Add-Member NoteProperty WebTemplate($ctx.Web.Webs[$i].WebTemplate)

                    Write-Output $obj
                    }

     }
     

   
     



}











#
#
#
#
#
#
# 
#
#
# Connect
#
#
#
#
#
#
#
#
#
#
#
#


function Connect-SPOCSOM
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

[CmdletBinding(DefaultParameterSetName="Credential")]
	param (
		[Parameter(Mandatory = $True, Position=1, ParameterSetName = "Credential")]
		$Credential,
		[Parameter(Mandatory = $True, Position=1, ParameterSetName = "Username")]
		[string]$Username,
		[Parameter(Mandatory = $True, Position=2)]
		[string]$Url
	)

	Switch ($PSCmdlet.ParameterSetName) {
		"Credential" {
			$Username = $Credential.Username
			$Password = $Credential.Password	
		}
		"Username" {
			$password = Read-Host "Password" -AsSecureString
		}
	}

 
  $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
  $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $password)
  $ctx.ExecuteQuery()  
$global:ctx=$ctx
}

function Connect-SPCSOM
{

<#
	.link
	http://social.technet.microsoft.com/wiki/contents/articles/32334.sharepoint-online-spomod-cmdlets-resources.aspx

  #>

	[CmdletBinding(DefaultParameterSetName="Credential")]
	param (
		[Parameter(Mandatory = $True, Position=1, ParameterSetName = "Credential")]
		$Credential,
		[Parameter(Mandatory = $True, Position=1, ParameterSetName = "Username")]
		[string]$Username,
		[Parameter(Mandatory = $True, Position=2)]
		[string]$Url
	)

	Switch ($PSCmdlet.ParameterSetName) {
		"Credential" {
			$Username = $Credential.Username
			$Password = $Credential.Password	
		}
		"Username" {
			$password = Read-Host "Password" -AsSecureString
		}
	}

  $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
  $ctx.Credentials = New-Object System.Net.NetworkCredential($Username, $password)
  $ctx.ExecuteQuery()  
$global:ctx=$ctx
}


$global:ctx






# Paths to SDK. Please verify location on your computer.
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll" 



Export-ModuleMember -Function "Connect-SPCSOM","New-SPOTerm", "Get-SPOTermGroups", "Get-SPOTermSets", "Get-SPOTermStore", "Get-SPOHashTagsTermSet", "Get-SPOHashTagsTerms", "Get-SPOKeyWordsTermSet", "New-SPOTermGroup", "Set-SPOTermGroup", "New-SPOTermSet", "Get-SPOTerm", "Set-SPOTermSet", "New-SPOListContentType","Get-SPOListItemVersions","Connect-SPOCSOM","New-SPOSiteContentType","New-SPOSiteContentTypeColumn","New-SPOListContentTypeColumn", "Get-SPOContentType", "Remove-SPOContentType","Set-SPOContentType","New-SPOListView","Set-SPOListView","Remove-SPOListView","Get-SPOListView","Get-SPOWeb","Get-SPOListCount","Get-SPOList", "Set-SPOList", "New-SPOList","Set-SPOListCheckout","Set-SPOListVersioning","Set-SPOListMinorVersioning","Remove-SPOListInheritance","Restore-SPOListInheritance","Set-SPOListContentTypesEnabled","Remove-SPOList","Set-SPOListFolderCreationEnabled","Set-SPOListIRMEnabled","Get-SPOListColumn","New-SPOListColumn","Set-SPOListColumn","Remove-SPOListColumn","Get-SPOListColumnFieldIsObjectPropertyInstantiated","Get-SPOListColumnFieldIsPropertyAvailable","New-SPOListChoiceColumn","Get-SPOListFields","Get-SPOListItems","New-SPOListItem","Remove-SPOListItemInheritance","Remove-SPOListItemPermissions","Restore-SPOListItemInheritance","Remove-SPOListItem","Set-SPOListItem","Set-SPOFileCheckout","Approve-SPOFile","Set-SPOFileCheckin","Copy-SPOFile","Remove-SPOFile","Deny-SPOFileApproval","Get-SPOFileIsPropertyAvailable","Move-SPOFile","Publish-SPOFile","Undo-SPOFileCheckout","Undo-SPOFilePublish","Get-SPOFolderFilesCount","Get-SPOFolderFiles","Get-SPOFileByServerRelativeUrl","Get-SPOFolderByServerRelativeUrl","Connect-SPOCSOM"
