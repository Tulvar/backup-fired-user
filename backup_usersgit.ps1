Import-Module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn 
#OU с заблокированными пользователями
$TargetOU= "OU=Уволенные,OU=Заблокированные,DC=domain,DC=local"

$date=Get-Date -Format g

#вычисление даты бекапа юзеров
$Buckdate=(get-date).AddDays(-7)

#пусть куда будет копироваться пользовательский профиль
$destpath="\\PC1\backup_users$"

#OU с забекапленными пользователями
$backupOU= "OU=Backup,OU=Заблокированные,DC=domain,DC=local"

#Поиск юзеров в OU
$blockusers=Get-ADUser -Filter * -SearchBase $TargetOU -Properties HomeDirectory,PasswordLastSet,Enabled
#перебираем пользователей в OU
foreach ($user in $blockusers)
{
    #выбираем заблокированных и тех у кого пароль менялся больше необходимого времени назад
    if (($user.Enabled -eq $false) -and ($user.PasswordLastSet -le $Buckdate ))
    {
    $san=$user.SamAccountName
    #путь куда будут положены архивы почтового ящика
    $PrimaryPath = $destpath + "\" + $san + "\" + $san + ".pst"
    #Создаем папку для бекапа пользователя
    New-Item -ItemType Directory -Force -Path $destpath\$san
        if ($user.HomeDirectory -ne $null)
        {         
            #Переносим пользовательскую папку в другое место    
            Move-Item -Path $user.HomeDirectory -Destination $destpath\$san
            #Отключаем пользовательскую папку
            Set-ADUser $user.SamAccountName -HomeDirectory $null
        } 
  
    #Делаем бекап почтового ящика пользователя и кидаем к папке с данными пользователя
    New-MailboxExportRequest -Mailbox $san -BatchName $san -FilePath $PrimaryPath
    #Проверяем статус выполнения бекапа почты
    $i=1;
    #Если статус Queued или InProgress цикл ожидает выполнения бекапа
        while ((Get-MailboxExportRequest -BatchName $san | Where {($_.Status -eq “Queued”) -or ($_.Status -eq “InProgress”)})) 
        {
        sleep 60
        Write-Host "Скрипт работает $i минут. Ожидаем завершения.."
        $i=$i+1
        }
    #Подчищаем за собой созданные запросы на выполнение бекапа почты
    Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false
    #записываем в поле описание дату бекапа
    Set-ADUser $user.SamAccountName -Description "backup at $date"
    #переносим пользователя в OU забекапленных учеток
    Move-ADObject -Identity $user -TargetPath $backupOU
    }
} 