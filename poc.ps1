cd C:\Users\user\Desktop\

$analysispath = 'c:\temp\analysis\'
$pattern = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"

$emails = @()
$urls = @()

mkdir $analysispath


Remove-Item $analysispath\* -Force -Recurse

$files = Get-ChildItem *

foreach($file in $files){
Remove-Item $analysispath\* -Force -Recurse -ErrorAction SilentlyContinue
#Expand-Archive -Path

$contents = [string](get-content -raw -Encoding Unknown -path $file.PSPath).ToCharArray();
$inttype = [convert]::tostring([convert]::toint32($contents[0]),16);
$inttype

    #if its a zip then extract 
    if($inttype -eq '4b50'){

    write-host 'zip file found' -ForegroundColor DarkRed
        if($file.Extension -eq '.zip'){  
        expand-archive -Path $file.FullName -DestinationPath $analysispath

        }
        else
        {
        write-host 'need to rename the file' -ForegroundColor DarkRed
        $path = $file.PSPath
        #$newname = $file.PSPath -replace '.doc',".zip"
    
        Copy-Item -Path $file -Destination $analysispath
            

        }
    
    Dir $analysispath | rename-item -newname { [io.path]::ChangeExtension($_.name, "zip") }
    
    $expand = Get-childItem $analysispath

    foreach($item in $expand){

    Expand-Archive -Path $item.FullName -DestinationPath $analysispath\expand\ -Force

    #read-host -Prompt 'PAUSE'
    }

    write-host "analyzing compressed items for interesting content" -ForegroundColor Green

    #Get any content in the expand folder in analysis path

    $analyze = Get-ChildItem "$analysispath\expand\" -Recurse -File
 
        foreach($afile in $analyze){

        write-host "SEARCHING for strings.........." -ForegroundColor DarkGreen

        #read-host "PAUSE"

        $message = Get-Content $afile.FullName

        $results = ($message | Select-String $pattern -AllMatches).Matches

        $results | Get-Unique

        $emails += ,$results


        if($message -match "\b(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])"){
            $urls += ,$_

            write-host "FOUND URL" -ForegroundColor Red 
                $urlpattern = "\b(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])"

    
            $url = ($message | Select-String $urlpattern -AllMatches).Matches
    
            $urls += ,$url
    

            }


    }

    }







    #end of archive found loops
}


write-host "############################################EMAILS###################################################" -ForegroundColor Gray

$emails.value #| Sort-Object -Unique
$emails.value | group | Sort-Object -Property Count -Descending

write-host "############################################URLS###################################################" -ForegroundColor Gray

$urls.Count
$urls.value #| Sort-Object -Unique #| Set-Clipboard
$urls.value | group | Sort-Object -Property Count -Descending

$uniq = $urls.value | Sort-Object -Unique
$uniq.count
    

    #clean the analysis path
    #Remove-Item $analysispath\* -Force -Recurse


