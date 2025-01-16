function Get-ApplicationName {
    param (
        [string]$ProcessPath
    )
    try {
        if ($ProcessPath -ne "Tidak Diketahui" -and (Test-Path $ProcessPath)) {
            $versionInfo = (Get-Item $ProcessPath).VersionInfo
            $appName = $versionInfo.ProductName
            if (-not $appName) {
                $appName = $versionInfo.FileDescription
            }
            if (-not $appName) {
                $appName = "Tidak Diketahui"
            }
            return $appName
        }
        else {
            return "Tidak Diketahui"
        }
    }
    catch {
        return "Tidak Diketahui"
    }
}

function Get-PortUsage {
    param (
        [int]$Port
    )
    $connections = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    $results = @()

    if ($connections) {
        foreach ($conn in $connections) {
            try {
                $process = Get-Process -Id $conn.OwningProcess -ErrorAction Stop
                $processPath = $process.Path
                if (-not $processPath) {
                    $processPath = "Tidak Diketahui"
                }
                $appName = Get-ApplicationName -ProcessPath $processPath

                $results += [PSCustomObject]@{
                    ProsesId      = $process.Id
                    NamaProses    = $process.ProcessName
                    NamaAplikasi  = $appName
                    AlamatLocal   = $conn.LocalAddress
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    ProsesId      = $conn.OwningProcess
                    NamaProses    = "Tidak Dikenal"
                    NamaAplikasi  = "Tidak Diketahui"
                    AlamatLocal   = $conn.LocalAddress
                }
            }
        }
    }

    return $results
}

function Show-MainMenu {
    Clear-Host
    Write-Host "=======================" -ForegroundColor Cyan
    Write-Host "PORT MANAGER TOOLS"
    Write-Host "=======================`n" -ForegroundColor Cyan
    Write-Host "1. Tampilkan Semua Aplikasi yang Menggunakan Port"
    Write-Host "2. Nonaktifkan Aplikasi yang Menggunakan Port"
    Write-Host "3. Keluar`n"
    $choice = Read-Host "Pilih opsi (1-3)"
    return $choice
}

function Show-PortMenu {
    Write-Host "`nPilih Port yang Ingin Diperiksa:"
    Write-Host "1. Port 80"
    Write-Host "2. Port 443"
    Write-Host "3. Port 3306"
    Write-Host "4. Kembali ke Menu Utama`n"
    $portChoice = Read-Host "Pilih port (1-4)"
    return $portChoice
}

function Disable-Application {
    param (
        [int]$Port
    )

    $portUsage = Get-PortUsage -Port $Port

    if ($portUsage.Count -eq 0) {
        Write-Host "`nTidak ada aplikasi yang menggunakan port $Port.`n" -ForegroundColor Yellow
        Read-Host "Tekan Enter untuk kembali..."
        return
    }

    Write-Host "`nAplikasi yang menggunakan port ${Port}:`n" -ForegroundColor Green
    $i = 1
    $portUsageIndexed = @()
    foreach ($item in $portUsage) {
        $portUsageIndexed += [PSCustomObject]@{
            No           = $i
            ProsesId     = $item.ProsesId
            NamaProses   = $item.NamaProses
            NamaAplikasi = $item.NamaAplikasi
            AlamatLocal  = $item.AlamatLocal
        }
        $i++
    }

    $portUsageIndexed | Format-Table -AutoSize

    Write-Host "`nPilih nomor aplikasi yang ingin dinonaktifkan (pisahkan dengan koma jika lebih dari satu, misal: 1,3):"
    $selected = Read-Host "Nomor Aplikasi"

    # Memproses input
    $selectedIndices = $selected -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }

    foreach ($index in $selectedIndices) {
        if ($index -ge 1 -and $index -le $portUsageIndexed.Count) {
            $app = $portUsageIndexed[$index - 1]
            Write-Host "`nMenonaktifkan aplikasi:" -ForegroundColor Cyan
            Write-Host "Proses ID     : $($app.ProsesId)"
            Write-Host "Nama Proses   : $($app.NamaProses)"
            Write-Host "Nama Aplikasi : $($app.NamaAplikasi)"
            Write-Host "Alamat Local  : $($app.AlamatLocal)`n"

            $confirm = Read-Host "Apakah Anda yakin ingin menonaktifkan aplikasi ini? (y/n)"
            if ($confirm -eq 'y' -or $confirm -eq 'Y') {
                try {
                    # Coba untuk menghentikan proses
                    Stop-Process -Id $app.ProsesId -Force -ErrorAction Stop
                    Write-Host "Aplikasi dengan Proses ID $($app.ProsesId) berhasil dinonaktifkan." -ForegroundColor Green
                }
                catch {
                    Write-Host "Gagal menonaktifkan aplikasi dengan Proses ID $($app.ProsesId). Mungkin perlu menjalankan script ini sebagai Administrator." -ForegroundColor Red
                }
            }
            else {
                Write-Host "Pembatalan menonaktifkan aplikasi." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Nomor aplikasi $index tidak valid." -ForegroundColor Red
        }
    }

    Read-Host "`nTekan Enter untuk kembali ke menu port..."
}

function Show-PortUsage {
    param (
        [int]$Port
    )

    $portUsage = Get-PortUsage -Port $Port

    Clear-Host
    Write-Host "`nAplikasi yang menggunakan port ${Port}:`n" -ForegroundColor Green

    if ($portUsage.Count -eq 0) {
        Write-Host "Tidak ada aplikasi yang menggunakan port $Port.`n" -ForegroundColor Yellow
    }
    else {
        $i = 1
        $portUsageIndexed = @()
        foreach ($item in $portUsage) {
            $portUsageIndexed += [PSCustomObject]@{
                No           = $i
                ProsesId     = $item.ProsesId
                NamaProses   = $item.NamaProses
                NamaAplikasi = $item.NamaAplikasi
                AlamatLocal  = $item.AlamatLocal
            }
            $i++
        }
        $portUsageIndexed | Format-Table -AutoSize
    }

    Read-Host "`nTekan Enter untuk kembali ke menu PORT"
}

function Main {
    do {
        $choice = Show-MainMenu

        switch ($choice) {
            '1' {
                do {
                    $portChoice = Show-PortMenu
                    switch ($portChoice) {
                        '1' { Show-PortUsage -Port 80 }
                        '2' { Show-PortUsage -Port 443 }
                        '3' { Show-PortUsage -Port 3306 }
                        '4' { break }
                        default { Write-Host "Pilihan tidak valid. Silakan pilih antara 1-4." -ForegroundColor Red }
                    }
                } while ($portChoice -ne '4')
            }
            '2' {
                # Nonaktifkan aplikasi yang menggunakan port
                do {
                    $portChoice = Show-PortMenu
                    switch ($portChoice) {
                        '1' { Disable-Application -Port 80 }
                        '2' { Disable-Application -Port 443 }
                        '3' { Disable-Application -Port 3306 }
                        '4' { break }
                        default { Write-Host "Pilihan tidak valid. Silakan pilih antara 1-4." -ForegroundColor Red }
                    }
                } while ($portChoice -ne '4')
            }
            '3' {
                Write-Host "Keluar dari Program. Terima kasih!" -ForegroundColor Cyan
                break
            }
            default {
                Write-Host "Pilihan tidak valid. Silakan pilih antara 1-3." -ForegroundColor Red
            }
        }
    } while ($choice -ne '3')
}

Main
