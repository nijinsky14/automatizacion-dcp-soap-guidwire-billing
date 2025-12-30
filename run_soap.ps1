# URL del servicio SOAP
$soapUrl = "http://bcqa.gruposancristobal.org.ar/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"


#QA  ---  $soapUrl ="http://bcqa.gruposancristobal.org.ar/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"
#UAT ---  $soapUrl ="https://bcuat.gruposancristobal.org.ar/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"
#DEV ---  $soapUrl ="http://bcdev/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"
#GW01---  $soapUrl ="http://diwin10gw01:8580/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"
#GW02---  $soapUrl ="http://diwin10gw02:8580/bc/ws/sancristobal/bc/webservice/producerpromises/PaymentsPromiseAPI?wsdl"

Write-Host "Utilizando ambiente:`n$soapUrl"
 
# Ruta base del script actual
$basePath = $PSScriptRoot

# Leer las plantillas XML desde archivos (relativas al script)
$soapCreateTemplate = Get-Content (Join-Path $basePath "soap_create.xml") -Raw
$soapAddItemTemplate = Get-Content (Join-Path $basePath "soap_add_item.xml") -Raw
$soapSetStatusTemplate = Get-Content (Join-Path $basePath "soap_set_status.xml") -Raw

# Abrir Excel
$excelPath = Join-Path $basePath "data1.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)

$row = 2 # Asumiendo encabezados en fila 1

while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
    # Leer datos de Excel

    $username = $sheet.Cells.Item($row, 1).Text.Trim()
    $password = $sheet.Cells.Item($row, 2).Text.Trim()
    $nivel = $sheet.Cells.Item($row, 3).Text.Trim()
    $producerCode = $sheet.Cells.Item($row, 4).Text.Trim()
    $organizerCode = $sheet.Cells.Item($row, 5).Text.Trim()
    $currency = $sheet.Cells.Item($row, 6).Text.Trim()
    $country = $sheet.Cells.Item($row, 7).Text.Trim()
    $branchOffice = $sheet.Cells.Item($row, 8).Text.Trim()
    $alias = $sheet.Cells.Item($row, 9).Text.Trim()
    $userPortal = $sheet.Cells.Item($row, 10).Text.Trim()

    $policyNumber = $sheet.Cells.Item($row, 11).Text.Trim()
    $valueAmount = $sheet.Cells.Item($row, 12).Text.Trim()
    $nroCuota = $sheet.Cells.Item($row, 13).Text.Trim()

    # --- Crear createProducerPromise ---
    $soapBodyCreate = $soapCreateTemplate -replace "{{username}}", $username `
                                          -replace "{{password}}", $password `
                                          -replace "{{nivel}}", $nivel `
                                          -replace "{{producerCode}}", $producerCode `
                                          -replace "{{organizerCode}}", $organizerCode `
                                          -replace "{{currency}}", $currency `
                                          -replace "{{country}}", $country `
                                          -replace "{{branchOffice}}", $branchOffice `
                                          -replace "{{alias}}", $alias `
                                          -replace "{{userPortal}}", $userPortal
    Write-Host "XML que se envía:`n$soapBodyCreate"
    try {
        # Enviar request createProducerPromise
        $responseCreateRaw = Invoke-RestMethod -Uri $soapUrl -Method Post -Body $soapBodyCreate -ContentType "application/soap+xml"
        [xml]$xmlResponse = $responseCreateRaw

        # Setup namespaces para XPath
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlResponse.NameTable)
        $nsMgr.AddNamespace("tns", "http://schemas.xmlsoap.org/soap/envelope/")
        $nsMgr.AddNamespace("san", "http://www.sancristobal.com.ar")

        $nodeReturn = $xmlResponse.SelectSingleNode("//san:createProducerPromiseResponse/san:return", $nsMgr)

if ($nodeReturn -ne $null -and $nodeReturn.InnerText) {
    $publicId = $nodeReturn.InnerText
    Write-Host "✅ publicIdPromise: $publicId para $producerCode"

    # Agregar ítems a la promesa mientras haya datos válidos en la fila
    while ($sheet.Cells.Item($row, 11).Value() -ne $null) {
        $policyNumber = $sheet.Cells.Item($row, 11).Text.Trim()
        $valueAmount = $sheet.Cells.Item($row, 12).Text.Trim()
        $nroCuota = $sheet.Cells.Item($row, 13).Text.Trim()

        $soapBodyItem = $soapAddItemTemplate -replace "{{username}}", $username `
                                             -replace "{{password}}", $password `
                                             -replace "{{publicIdPromise}}", $publicId `
                                             -replace "{{policyNumber}}", $policyNumber `
                                             -replace "{{valueAmount}}", $valueAmount `
                                             -replace "{{nroCuota}}", $nroCuota

        Write-Host "XML que se envía:`n$soapBodyItem"
        Start-Sleep -Seconds 2

        $responseItem = Invoke-RestMethod -Uri $soapUrl -Method Post -Body $soapBodyItem -ContentType "application/soap+xml"
        Write-Host "✅ Item $policyNumber agregado para $producerCode"

        $row++
    }

    # Después de agregar todos los ítems, cambiar el estado
    $soapBodyStatus = $soapSetStatusTemplate -replace "{{username}}", $username `
                                             -replace "{{password}}", $password `
                                             -replace "{{publicIdPromise}}", $publicId
    Write-Host "XML que se envía:`n$soapBodyStatus"

    try {
        $responseStatus = Invoke-RestMethod -Uri $soapUrl -Method Post -Body $soapBodyStatus -ContentType "application/soap+xml"
        Write-Host "✅ Estado cambiado a Pending para $producerCode"
    }
    catch {
        Write-Host "❌ Error en setStatusToPending:"
        if ($_.Exception.Response -ne $null) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody
        }
        else {
            Write-Host $_.Exception.Message
        }
    }
}
        else {
            Write-Host "❌ No se encontró publicIdPromise para $producerCode"
        }
    }
    catch {
        Write-Host "❌ Error con $producerCode"
        Write-Host $_.Exception.Message
    }

    $row++
}

# Cerrar Excel
$workbook.Close($false)
$excel.Quit()

