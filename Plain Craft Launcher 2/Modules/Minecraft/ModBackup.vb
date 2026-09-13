Public Module ModBackup

    '启动游戏前将版本下的每个存档分别打包备份到版本文件夹下的 Backup 中，并只保留最新的若干份。
    '内容没有变动的存档不会被重复备份，备份失败绝不能阻止游戏启动，因此所有异常都只会记录日志。

    ''' <summary>
    ''' 记录存档最近一次成功备份的指纹的清单文件名。
    ''' 它只会由 PCL 写入，因此也用于区分 PCL 建立的备份文件夹与用户手动放入的文件夹。
    ''' </summary>
    Private Const ManifestFileName As String = "PCL-Backup.ini"

#Region "对外接口"

    ''' <summary>
    ''' 指定版本在启动时是否应当自动备份存档。
    ''' 会考虑版本独立设置中的三态覆盖。
    ''' </summary>
    Public Function McBackupSavesAble(Instance As McInstance) As Boolean
        If Instance Is Nothing Then Return False
        Select Case Settings.Get(Of Integer)("VersionAdvanceBackupSaves", Instance:=Instance)
            Case 0 '关闭
                Return False
            Case 1 '开启
                Return True
            Case Else '跟随全局设置
                Return Settings.Get(Of Boolean)("LaunchAdvanceBackupSaves")
        End Select
    End Function

    ''' <summary>
    ''' 指定版本的每一个存档在启动时应当保留的备份份数。
    ''' 只有在该版本的独立设置中自行开启备份时才使用版本自己的份数，其余情况一律沿用全局设置，
    ''' 这样 “跟随全局设置” 才是彻底跟随。
    ''' </summary>
    Public Function McBackupSavesKeep(Instance As McInstance) As Integer
        If Instance IsNot Nothing AndAlso Settings.Get(Of Integer)("VersionAdvanceBackupSaves", Instance:=Instance) = 1 Then
            Return Settings.Get(Of Integer)("VersionAdvanceBackupSavesCount", Instance:=Instance)
        End If
        Return Settings.Get(Of Integer)("LaunchAdvanceBackupSavesCount")
    End Function

    ''' <summary>
    ''' 指定版本存放存档备份的文件夹，以 \ 结尾。
    ''' 该文件夹与 PCL 文件夹平级，因此不会被导出整合包时打包。
    ''' </summary>
    Public Function McBackupFolder(Instance As McInstance) As String
        Return Instance.PathVersion & "Backup\"
    End Function

#End Region

#Region "启动时备份"

    ''' <summary>
    ''' 启动流程中的存档备份加载器。
    ''' </summary>
    Public Sub McLaunchBackupSaves(Loader As LoaderTask(Of Integer, Integer))
        Try
            Dim Instance = McInstanceSelected
            If Not McBackupSavesAble(Instance) Then Return
            Dim SavesFolder As String = Instance.PathIndie & "saves\"
            Dim BackupFolder As String = McBackupFolder(Instance)
            '存档不存在或为空时不备份，但不能就此返回：
            '否则最后一个存档被删除后，它留下的备份将永远得不到清理
            Dim SaveFolders = New List(Of String)
            If DirectoryUtils.Exists(SavesFolder) AndAlso Not DirectoryUtils.IsEmpty(SavesFolder) Then
                SaveFolders = DirectoryUtils.EnumerateDirectories(SavesFolder).ToList()
            Else
                Logger.Info($"未找到任何存档，跳过自动备份：{SavesFolder}")
            End If
            If Loader.IsCanceled Then Return
            '逐个存档备份，内容没有变动的存档会自行跳过
            Dim Keep As Integer = McBackupSavesKeep(Instance)
            For Index = 0 To SaveFolders.Count - 1
                If Loader.IsCanceled Then Return
                BackupOneSave(SaveFolders(Index), BackupFolder, Keep)
                Loader.Progress = 0.05 + 0.9 * (Index + 1) / Math.Max(1, SaveFolders.Count)
            Next
            If Loader.IsCanceled Then Return
            '存档已被删除或移走时，它留下的备份也不再保留
            CleanUpBackups(SavesFolder, BackupFolder)
            Loader.Progress = 1
        Catch ex As Exception
            Logger.Warn(ex, "自动备份存档失败，将直接开始启动游戏")
        End Try
    End Sub

    ''' <summary>
    ''' 将单个存档打包备份到 Backup\存档名\ 中，并在打包成功后记录其指纹。
    ''' 指纹与上次备份时一致时不会产生任何新文件。
    ''' </summary>
    Private Sub BackupOneSave(SaveFolder As String, BackupFolder As String, Keep As Integer)
        Try
            Dim SaveName As String = PathUtils.GetLastPart(SaveFolder)
            Dim SaveBackupFolder As String = BackupFolder & SaveName & "\"
            Dim ManifestPath As String = SaveBackupFolder & ManifestFileName
            Dim Fingerprint As String = GetSavesFingerprint(SaveFolder)
            '指纹一致，且上次打包的文件确实还在，说明这个存档没有变动
            Dim LastBackup As String = ReadIni(ManifestPath, "File")
            If ReadIni(ManifestPath, "Fingerprint") = Fingerprint AndAlso
               LastBackup <> "" AndAlso FileUtils.Exists(SaveBackupFolder & LastBackup) Then
                Logger.Info($"存档没有变动，跳过备份：{SaveFolder}")
                Return
            End If
            '打包
            Dim ZipPath As String = SaveBackupFolder & SaveName & " " & Date.Now.ToString("yyyy'-'MM'-'dd HH'-'mm'-'ss") & ".zip"
            Logger.Info($"正在备份存档：{SaveFolder} → {ZipPath}")
            FileUtils.CreateZipFromDirectory(ZipPath, SaveFolder)
            '打包成功后才记录指纹，这样打包失败时不会留下 “已经备份过” 的记录
            WriteIni(ManifestPath, "Fingerprint", Fingerprint)
            WriteIni(ManifestPath, "File", PathUtils.GetLastPart(ZipPath))
            Logger.Info($"存档备份完成：{ZipPath}")
            '删除过旧的备份
            DeleteOldBackups(SaveBackupFolder, Keep)
        Catch ex As Exception
            Logger.Warn(ex, $"自动备份存档失败：{SaveFolder}")
        End Try
    End Sub

    ''' <summary>
    ''' 计算存档文件夹的指纹，由其中所有文件的相对路径、长度与最后写入时间构成。
    ''' 只读取目录项而不读取文件内容，因此即使存档很大，开销也可以忽略。
    ''' </summary>
    Private Function GetSavesFingerprint(SaveFolder As String) As String
        Dim Entries = DirectoryUtils.EnumerateFiles(SaveFolder, includeSubDirectories:=True).
                      OrderBy(Function(FilePath) FilePath, StringComparer.OrdinalIgnoreCase).
                      Select(Function(FilePath)
                                 Dim Info As New FileInfo(FilePath)
                                 '使用相对路径，这样存档文件夹本身被移动或改名后指纹也不会改变
                                 Dim RelativePath = If(FilePath.StartsWith(SaveFolder, StringComparison.OrdinalIgnoreCase),
                                                       FilePath.Substring(SaveFolder.Length), FilePath)
                                 Return $"{RelativePath}|{Info.Length}|{Info.LastWriteTimeUtc.Ticks}"
                             End Function)
        Return String.Join(";", Entries).GetStableHashCode().ToString("X16")
    End Function

    ''' <summary>
    ''' 删除指定存档的备份文件夹中除了最新的 Keep 份以外的所有备份。
    ''' </summary>
    Private Sub DeleteOldBackups(SaveBackupFolder As String, Keep As Integer)
        Keep = Math.Max(1, Keep)
        '先完整枚举再删除，避免边遍历文件夹边删除文件
        Dim OldFiles = DirectoryUtils.EnumerateFiles(SaveBackupFolder, searchPattern:="*.zip").
                       OrderByDescending(Function(Path) New FileInfo(Path).LastWriteTime).
                       Skip(Keep).
                       ToList()
        For Each OldFile In OldFiles
            Try
                Logger.Info($"删除过旧的存档备份：{OldFile}")
                FileUtils.Delete(OldFile)
            Catch ex As Exception
                Logger.Warn(ex, $"删除过旧的存档备份失败：{OldFile}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 清除存档已被删除或移走时留下的备份文件夹。
    ''' </summary>
    Private Sub CleanUpBackups(SavesFolder As String, BackupFolder As String)
        For Each BackupFolderItem In DirectoryUtils.EnumerateDirectories(BackupFolder)
            Try
                '只处理 PCL 自己建立的备份文件夹，避免误删用户手动放入的内容
                If Not FileUtils.Exists(BackupFolderItem & "\" & ManifestFileName) Then Continue For
                Dim SaveName As String = PathUtils.GetLastPart(BackupFolderItem)
                If DirectoryUtils.Exists(SavesFolder & SaveName) Then Continue For
                Logger.Info($"存档已不存在，删除其备份：{BackupFolderItem}")
                DirectoryUtils.Delete(BackupFolderItem)
            Catch ex As Exception
                Logger.Warn(ex, $"删除已不存在的存档的备份失败：{BackupFolderItem}")
            End Try
        Next
    End Sub

#End Region

End Module
