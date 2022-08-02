Clear-Host
$workingPath = 'C:\Users\opervpn\Desktop\teste\selenium'


# Add the working directory to the environment path.
# This is required for the ChromeDriver to work.
if (($env:Path -split ';') -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}


# OPTION 1: Import Selenium to PowerShell using the Add-Type cmdlet.
Add-Type -Path "$($workingPath)\WebDriver.dll"


$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
                        
$ChromeOptions.AddArgument('start-maximized')

$ChromeOptions.AcceptInsecureCertificates = $True

#$ChromeOptions.addArguments('headless')

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)

$ChromeDriver.Navigate().GoToURL('https://outlook.live.com/owa/')

sleep -Seconds 3

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('/html/body/header/div/aside/div/nav/ul/li[2]/a')).Click()

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="i0116"]')).SendKeys('ContaMail')

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="idSIButton9"]')).Click()

sleep -Seconds 3

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="i0118"]')).SendKeys('Password')

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="idSIButton9"]')).Click()

sleep -Seconds 3

$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="idSIButton9"]')).Click()

sleep -Seconds 1

$ChromeDriver.Navigate().GoToURL('https://outlook.office365.com/mail/AQMkADRiYTA0MDY1LTYwYjktNDBkYy05N2Y4LWI1YmYxZDY0OGQyNwAuAAAD0WDYIMPuQUOTJpQrDai78wEAmg%2B%2B%2FCfYR0asIVao1SBr%2BQAAAjFLAAAA')



sleep -Seconds 5
$emails = $ChromeDriver.FindElement([OpenQA.Selenium.By]::CssSelector('#MainModule > div > div > div.AsZHetgbr0sNPWTic7WD.css-179 > div > div > div.IrLT6us4d4A53AD27CcM.css-183 > div.oQNugVUw3c8iisQbidOJ.customScrollBar > div > div > div > div'))
$emails = $emails.Text
$emails = $emails.split("`n")

if ($emails[3].split(":")[0] -eq (get-date -UFormat %H)){

        VerifyExel
}
else {Write-Output "Sem nada a fazer"
    #VerifyExel

}


function VerifyExel
{


sleep -Seconds 4

$css = $ChromeDriver.FindElement([OpenQA.Selenium.By]::CssSelector('#MainModule > div > div > div.AsZHetgbr0sNPWTic7WD.css-179 > div > div > div.IrLT6us4d4A53AD27CcM.css-183 > div.oQNugVUw3c8iisQbidOJ.customScrollBar > div > div > div > div > div > div > div')).Click()

sleep -Seconds 3

$seta = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="ReadingPaneContainerId"]/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/button')).Click()

#fluent-default-layer-host > div:nth-child(2) > div > div > div > div.ms-Callout-main.calloutMain-394 > div > div > ul > li:nth-child(5) > button > div


$seta1 = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//*[@id="fluent-default-layer-host"]/div[2]/div/div/div/div[3]'))

[OpenQA.Selenium.Interactions.Actions]$actions = New-Object OpenQA.Selenium.Interactions.Actions($ChromeDriver)

$actions.MoveByOffset(667,392)

sleep -Seconds 2

$actions.Click().Perform()


exit 
sleep -Seconds 3

[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')


sleep -Seconds 2


# path do folder download do user
$caminho = "C:\Users\opervpn\Downloads\"
# procurar os ficheiros com o nome parcial do ficheiro
$procurar = Get-ChildItem -Path $caminho -Recurse -Filter "FileName_*" | Select-Object Name
# se existir mais que um, utilizar o primeiro (mais recente)
if ($procurar.Length -gt 1){
    $file = $caminho+$procurar.Name[0]
}
else {
   $file = $caminho+$procurar.Name 
}
# ler o CSV e eliminar as primeiras 5 linhas
$csv = Get-Content $file | Select-Object -skip 4 
# transformar o output anterior em array separado por linha
$csv = $csv.split("`n")
foreach ($x in $csv){
   # retirar os espaços em branco e separar a linha por ;
   $linha = $x.split(";")
   #Write-Host $linha
   #sleep -Seconds 1
   # [-2] é a "ultima posição" para apanhar o valor OK | Not OK
   if($linha[-2] -eq " NOT OK"){
   #Write-Host "-2"
        $x
    }
    elseif($linha[-1] -eq " NOT OK"){
        
        #Write-Host "-1" 
        $x 
    }
}

# depois de todo o processo se tudo correr bem, eliminar o ficheiro :D
Write-Output "APAGAR: "$file  
Remove-Item $file

}

#$ChromeDriver.Close()
#$ChromeDriver.Quit()