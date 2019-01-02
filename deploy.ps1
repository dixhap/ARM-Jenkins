try{

    #$cred = Get-Credential   #Getting the credentials from pop-up
    #Login-AzureRmAccount -Credential $cred  # Loging to the Azure Subscription
    
    
    #$file = "Cloud_Foundation.xlsx"  
    $TemplateFileLocation="Template.json" 
    $TemplateVNetFileName="azure_vnetdeploy.json"    
    $TemplateSubnetFileName="azure_subnetdeploy.json" 
    $TemplateVNetFile= $TemplateFileLocation + $TemplateVNetFileName
    $TemplateSubnetFile= $TemplateFileLocation + $TemplateSubnetFileName
    $sheetName = "Vnet-Subnet"   
    $objExcel = New-Object -ComObject Excel.Application     
    $workbook = $objExcel.Workbooks.Open($file)             
    $sheet = $workbook.Worksheets.Item($sheetName)
    $objExcel.Visible=$false 
    $colMax = ($sheet.UsedRange.Columns).count            
    
    if($colMax -ge 2)
    {
    for($j= 0 ; $j -le $colMax -2 ; $j++)
    {
    
    $DeploymentType = $sheet.Cells.Item(2,$j+2).text 
    
    if(($DeploymentType).ToLower() -eq 'yes')
    {
    $rgName = $sheet.Cells.Item(3,$j+2).text         
    $location= $sheet.Cells.Item(4,$j+2).text
    
    
    $get =Get-AzureRmResourceGroup -Name $rgName -Location $location -ErrorAction SilentlyContinue
    if($get.count -eq 0)
    {  
      Write-Host 'creating Resource group...'
      New-AzureRmResourceGroup -Name $rgName -Location $location
      Write-Host "Resource Group '$RGName' is created"
    }
    
    $vnetCount = ($colMax - 1)
    
    if($vnetCount -gt 0)
    {   
        $vnetName = $sheet.Cells.Item(5,$j+2).text
        $vnetCIDRRange = $sheet.Cells.Item(6,$j+2).text          
    
          $parameters = @{"vnetName"=$vnetName.Trim()
                          "vnetCIDRRange"=$vnetCIDRRange.Trim()
                          "vnetTagValues"=@{}   
                          } 
          New-AzureRmResourceGroupDeployment -Name vnet -ResourceGroupName $rgName -TemplateFile $TemplateVNetFile -TemplateParameterObject $parameters -Verbose -Mode Incremental
          Write-Host "Vnet '$vnetName' is created"
    
        $subnetcount = $sheet.Cells.Item(7,$j+2).text
    
        for($m=0;$m -lt $subnetcount;$m++)
        {
        $subnetName = $sheet.Cells.Item(8+($m*2),$j+2).text
        $subnetRange = $sheet.Cells.Item(9+($m*2),$j+2).text     
            
    
            $parameters = @{"vnetName"=$vnetName.Trim()
                            "subnetName"=$subnetName.Trim()
                            "subnetCIDRRange"=$subnetRange.Trim()
                            "subnetnetTagValues"=@{}   
                            } 
             New-AzureRmResourceGroupDeployment -Name subnet -ResourceGroupName $rgName -TemplateFile $TemplateSubnetFile -TemplateParameterObject $parameters -Verbose -Mode Incremental
             Write-Host "Subnet '$subnetName' is created"         
                   
        }  
    }
    }
    }
    }
    $objExcel.Workbooks.Close()
    
}
    catch
    {
    $objExcel.Workbooks.Close() 
    Write-Output $_.Exception
    }
