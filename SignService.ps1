# Полный путь до установленной КриптоПРО PDF
$CPPDFutil = "C:\example\cppdfutil.exe"
# Полный путь до рабочей папки со структурой:
# Фамилия1
#    На_Подпись
#    Подписано
# Сертификаты
#    Фамилия1.cer
$WorkDir = "C:\SignService\"
# Путь до папки с сертификатами
$CertsDir = $WorkDir + "Сертификаты\"
# Путь до папки с логами
$logDir = "C:\SignService\!Логи\"
# Пути до сетевой папки
$NetPathForSign = "M:\example\"
$NetPathForArc = "M:\example\Подписано\"

# Пересоздавать файл лога
function ManageLogs{
    
    $logFileName = $logDir + "$(Get-Date -Format 'yyyy-MM-dd')-SignService.txt"

    if(Test-Path $logFileName){
    
        return $logFileName
    
    } else {

        $logfile = New-Item -Path $logFileName -ItemType File -ErrorAction Ignore

        return $logfile.FullName

    }

}

# Конвертация из WORD в PDF
function WordToPDF{

    Param(
        $logFileName,
        $WordFile
    
    )

    $pdfNewName = $WordFile.DirectoryName + "\" + $WordFile.BaseName + ".pdf"

    try{
        
        Write-Output "$(Get-Date)    Попытка конвертации WORD в PDF" | Tee-Object -Append $logFileName

        $word = New-Object -ComObject Word.Application
        $document = $word.Documents.Open($WordFile.FullName)
        $document.SaveAs([ref] $pdfNewName, [ref] 17)
        $document.Close()
        $word.Quit()

        Write-Output "$(Get-Date)    Документ $WordFile успешно сконвертирован в PDF" | Tee-Object -Append $logFileName

    } catch {
    
        Write-Output "$(Get-Date)    Не сработала конвертация из WORD в PDF!" | Tee-Object -Append $logFileName
        $word.Quit()

    }

}

# Конвертация из EXCEL в PDF
function ExcelToPDF{

    Param(
        $logFileName,
        $ExcelFile
    
    )

    $pdfNewName = $ExcelFile.DirectoryName + "\" + $ExcelFile.BaseName + ".pdf"

    try{
        
        Write-Output "$(Get-Date)    Попытка конвертации EXCEL в PDF" | Tee-Object -Append $logFileName

        $excel = New-Object -ComObject Excel.Application
        $ExcelFilePath = $ExcelFile.FullName
        $document = $excel.workbooks.open($ExcelFilePath, 3)
        $document.Saved = $true
        $document.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfNewName)
        $excel.Workbooks.close()
        $excel.Quit()

        Write-Output "$(Get-Date)    Документ $ExcelFile успешно сконвертирован в PDF" | Tee-Object -Append $logFileName

    } catch {
    
        Write-Output "$(Get-Date)    Не сработала конвертация из EXCEL в PDF!" | Tee-Object -Append $logFileName
        $excel.Quit()
    
    }

}

# Функция определения является ли файл PDF, если нет произвести конвертацию
function MakePDF{

    Param(
    
        $logFileName,
        $fileobj

    )

    $fileext = $fileobj.Extension
    
    if($fileext.ToUpper() -match "DOC" -or $fileext.ToUpper() -match "DOCX"){
    
        Write-Output "$(Get-Date)    Обнаружен документ типа WORD - $fileobj" | Tee-Object -Append $logFileName
        WordToPDF $logFileName $fileobj
        Write-Output "$(Get-Date)    Удаление файла $fileobj" | Tee-Object -Append $logFileName
        # Удаление Файла WORD
        Remove-Item $fileobj.FullName
        Write-Output "$(Get-Date)    Файл $fileobj успешно удален" | Tee-Object -Append $logFileName
    
    } elseif($fileext.ToUpper() -match "XLS" -or $fileext.ToUpper() -match "XLSX"){
    
        Write-Output "$(Get-Date)    Обнаружен документ типа EXCEL - $fileobj" | Tee-Object -Append $logFileName
        ExcelToPDF $logFileName $fileobj
        Write-Output "$(Get-Date)    Удаление файла $fileobj" | Tee-Object -Append $logFileName
        # Удаление Файла EXCEL
        Remove-Item $fileobj.FullName
        Write-Output "$(Get-Date)    Файл $fileobj успешно удален" | Tee-Object -Append $logFileName
    
    }

}

# Перенос файлов с сетевой папки на локал
function DownloadFromNet{
    
    Write-Output "$(Get-Date)    Начинаю проверку сетевого расположения на наличие новых файлов" | Tee-Object -Append $logFileName
    $NetDirsForSign = Get-ChildItem $NetPathForSign

    foreach($NetDirForSign in $NetDirsForSign){
        
        Write-Output "$(Get-Date)    Просматриваю папку $($NetDirForSign.Name)" | Tee-Object -Append $logFileName
        $NetFiles = Get-ChildItem $NetDirForSign.FullName -File

        if($NetFiles){
            
            Write-Output "$(Get-Date)    Обнаружены документы на подпись в папке $($NetDirForSign.Name)" | Tee-Object -Append $logFileName
            $localDirPath = $WorkDir + $NetDirForSign.Name + "\На_Подпись"
        
            foreach($NetFile in $NetFiles){
                
                Write-Output "$(Get-Date)    Попытка переноса файла $($NetDirForSign.Name) на подпись в локальную папку" | Tee-Object -Append $logFileName
                
                try{

                    Move-Item $NetFile.FullName $localDirPath -ErrorAction SilentlyContinue
                    Write-Output "$(Get-Date)    Файл $($NetFile.Name) перенесен успешно!" | Tee-Object -Append $logFileName
                
                } catch{
                
                    Write-Output "$(Get-Date)    Файл $($NetFile.Name) не удалось перенести, похоже файл открыт в другой программе!" | Tee-Object -Append $logFileName

                }

            }
        
        } else {
        
            Write-Output "$(Get-Date)    Файлы для обработки в сетевой папке $NameDir не найдены" | Tee-Object -Append $logFileName

        }
    
    }

}

# Перенос файлов с локала на сетевой диск
function UploadToNet{

    $dirs = Get-ChildItem $WorkDir -Exclude "!Логи", "Сертификаты"

    foreach($dir in $dirs){
    
        Write-Output "$(Get-Date)    Начинаю просмотр локальных папок" | Tee-Object -Append $logFileName
        $localDirToUpload = $dir.FullName + "\Подписано"

        Write-Output "$(Get-Date)    Просматриваю локальную папку $($dir.Name) на наличие обработанных файлов" | Tee-Object -Append $logFileName
        $filesToUpload = Get-ChildItem $localDirToUpload

        if($filesToUpload){
            Write-Output "$(Get-Date)    Обнаружены обработанные документы в папке $($dir.Name)" | Tee-Object -Append $logFileName
            $remoteDirToUpload = "M:\example\" + $dir.Name + "\Подписано"

            foreach($fileToUpload in $filesToUpload){
            
                Write-Output "$(Get-Date)    Попытка переноса файла $($fileToUpload.Name) в сетевую папку" | Tee-Object -Append $logFileName
                try{
                
                    Write-Output "$(Get-Date)    Файл $($fileToUpload.Name) перенесен успешно!!" | Tee-Object -Append $logFileName
                    Move-Item $filesToUpload.FullName $remoteDirToUpload -ErrorAction SilentlyContinue

                } catch{
                
                    Write-Output "$(Get-Date)    Файл $($fileToUpload.Name) не удалось перенести, похоже файл открыт в другой программе!" | Tee-Object -Append $logFileName
                
                }
            
            }
        
        }
    
    }


}

function SignFiles{

    Param(
    
        $logFileName
    
    )

    DownloadFromNet

    $NamesDir = Get-ChildItem $WorkDir -Directory

    foreach($NameDir in $NamesDir){

        if(($NameDir -notlike "Сертификаты") -and ($NameDir -notlike "!Логи")){
    
            $SignDir = $NameDir.FullName + "\На_Подпись\"
            $DoneDir = $NameDir.FullName + "\Подписано\"
            $Cert = $CertsDir + $NameDir.Name + ".cer"
            $SignFiles = Get-ChildItem $SignDir
        
            if($SignFiles){

                Write-Output "$(Get-Date)    Обнаружены документы на подпись в папке $NameDir" | Tee-Object -Append $logFileName

                foreach($SignFile in $SignFiles){
                
                    if(!($SignFile.Extension.ToUpper() -match "PDF")){

                        MakePDF $logFileName $SignFile

                    }
                
                }
        
                try{
            
                    Write-Output "$(Get-Date)    Попытка подписи файлов" | Tee-Object -Append $logFileName
                    $errDir = $NameDir.FullName + "\Ошибки_подписи"
                    Start-Process -FilePath $CPPDFutil -ArgumentList "sign -i $SignDir -o $DoneDir -c $Cert -r $logDir -e $errDir -E -A law --X 0 --Y -0" -Wait
                    Write-Output "$(Get-Date)    Файлы успешно подписаны" | Tee-Object -Append $logFileName
                    Write-Output "$(Get-Date)    Очистка директории на подпись" | Tee-Object -Append $logFileName
                    $FilesToRemove = Get-ChildItem $SignDir
                    foreach($FileToRemove in $FilesToRemove){
                        
                        Write-Output "$(Get-Date)    Удаление файла $FileToRemove" | Tee-Object -Append $logFileName
                        Remove-Item $FileToRemove.FullName

                    }

                    Write-Output "$(Get-Date)    Директория очищена" | Tee-Object -Append $logFileName

                } catch {
            
                    Write-Output "$(Get-Date)    Что-то пошло не так!" | Tee-Object -Append $logFileName
            
                }

            }


            
            Write-Output "$(Get-Date)    Файлы для обработки в папке $NameDir не найдены" | Tee-Object -Append $logFileName

        }

    }

    UploadToNet

}

while($true){
    
    $logFileName = ManageLogs
    Write-Output "$(Get-Date)    __________________________" | Tee-Object -Append $logFileName
    Write-Output "$(Get-Date)    Поиск файлов для обработки" | Tee-Object -Append $logFileName
    SignFiles $logFileName
    [System.GC]::Collect()
    Write-Output "$(Get-Date)    Пауза 10 секунд" | Tee-Object -Append $logFileName
    Start-Sleep 10

}