$ds1="130d388d6bb444"
$ds2="f7b6e4ce1937e91aed"
$exa1="087c3d3f-f03c"
$exa2="-464f-ac04-914de869696f"
$c="DEEPSEEK_API_KEY=sk-$ds1$ds2`r`nEXA_API_KEY=$exa1$exa2`r`n"
Set-Content -Path 'C:\Users\Rei\AppData\Local\hermes\.env' -Value $c -Force
Write-Output "Done"
