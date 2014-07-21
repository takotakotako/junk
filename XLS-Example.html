$config = @{
    "csvdir" = @(
        "パス",
        "パス"
    );
    
    "hcode" = @("コード","コード");

    "basexls" = @{
        "path"="ベースのパス";
        "sheet"="ベースのシート";
        "x"=3;
        "y"=2;
    }

    "xlsoutput" = "アウトプットパス"
}


$config.csvdir | % {
    echo "csvdir:$_" # debug
    (Get-ChildItem $_).Fullname | ? {$_ -like "*.csv"} | % {
        $csvpath = $_ 

        $nodename = (Split-Path $csvpath -Leaf).split("_")[0]
 
        $config.hcode | ? { (Split-Path $nodename -Leaf) -like ("*" + $_ + "*")} | % {
            $hcode = $_

            echo "" #debug
            echo "csvpath:$csvpath" #debug
            echo "nodename:$nodename" #debug
            echo "" #debug

            $csvcontent = Get-Content $csvpath

            if ( $csvcontent[0] -like "*CPU*") {
                $resourcetype = "CPU"
            } elseif ( ($csvcontent[0] -like "*メモリ*") -or ($csvcontent[0] -like "*Memory*") ) {
                $resourcetype = "Memory"
            } else {
                $resourcetype = "Disk"
            }

            echo "ResourceType:$resourcetype" #debug


			if ( $csvcontent[1] -like "Agents:*5A1*") {
                $nodetype = "物理サーバー"
                $resourcedata = $csvcontent[6..($csvcontent.Length-1)]
            } elseif ( $csvcontent[1] -like "Agents:*7A1*") {
                $nodetype = "仮想マシン"
                $resourcedata = $csvcontent[6..($csvcontent.Length-1)]
            } else {
                $nodetype = "物理マシン"
                $resourcedata = $csvcontent[2..($csvcontent.Length-1)]
            }

            echo "NodeType:$nodetype " #debug

        }
    }
}


echo "" #debug
echo "xls pert" #debug
echo "" #debug

$config.hcode | ? {$_ -ne $null} | % {
    $hcode = $_

    echo $hcode #debug

    $xlscom = New-Object -ComObject Excel.Application
    $xlscom.Visible = $True
    $book = $xlscom.Workbooks.Open($config.basexls.path)
    $book.Activate()

    $sheet = $book.Worksheets.Item($config.basexls.sheet)
    $sheet.Activate()

    $sheet.copy($sheet)
    $sheet = $book.Worksheets.Item($config.basexls.sheet + " (2)")
    $sheet.Activate()
#    $book.Worksheets.Item($config.basexls.sheet + " (2)").Remove()

    $sheet.name = "ノード名"
    $sheet.Cells.Item($config.basexls.y, $config.basexls.x) = "書き込み"        
    $sheet.PageSetup.RightHeader = "右ヘッダー書き込み"

#    $sheet.PrintOut()
    $book.SaveAs($config.xlsoutput + $hcode + ".xlsx")    
    $xlscom.Quit()

}
