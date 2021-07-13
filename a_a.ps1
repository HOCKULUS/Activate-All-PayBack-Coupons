$ie = New-Object -ComObject 'internetExplorer.Application' -ErrorAction Ignore -ErrorVariable global:Fehler
$ie.Visible = $true
$ie.Navigate("https://www.payback.de/login?") #Logindaten eingeben - kann bei weiterer Nutzung auskommentiert werden, wenn "angemeldet bleiben" ausgewählt wurde
While($ie.Busy -eq $true){Start-Sleep -s 3}
do{
Start-Sleep -Milliseconds 1
}until($ie.LocationURL -ne "https://www.payback.de/info/datenschutz/hinweise-cookies/wall?redirectUrl=https%3A%2F%2Fwww.payback.de%2Flogin" -and $ie.LocationURL -ne "https://www.payback.de/login?")
$ie.Visible = $false #Browserfenster wird ausgeblendet
While($ie.Busy -eq $true){Start-Sleep -s 3}
$ie.Navigate2("https://www.payback.de/coupons?partnerId=")
While($ie.Busy -eq $true){Start-Sleep -s 3}
Start-Sleep -s 30 #Zeitpuffer bis alle Coupons geladen wurden
$coupons = $ie.Document.getElementsByClassName("not-activated")
foreach($coupon in $coupons){
if($coupon.innerText -match "JETZT AKTIVIEREN"){
$coupon.click()
Start-Sleep -s 2
}
}
