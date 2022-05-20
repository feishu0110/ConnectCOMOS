Class MainWindow
    '{3BDD1BA3-E09C-11CF-9BB7-00A0248A9BEE} 
    Private Workset As Plt.IComosDWorkset
    Private CurUser As Plt.IComosDUser
    Private PrjCol As Plt.IComosDCollection '项目集合
    Private WLCol As Plt.IComosDCollection '工作层集合
    Private clsStdMod = New ComosXStdMod.XStdModClass()



    Private Sub GetQryData_Click(sender As Object, e As RoutedEventArgs) '获取指定项目指定Query的数据
        txtMsgInfo.Text = "..."
        If Workset Is Nothing Then
            Dim blnCheck As Boolean
            blnCheck = GetWorkset()
            If blnCheck = False Then Exit Sub
        End If
        Dim strPrjName As String
        Dim strQryPFN As String
        Dim strWLID As String
        strPrjName = txtPrjName.SelectedValue
        If String.IsNullOrEmpty(strPrjName) Then
            MessageBox.Show(“请先输入项目号”, "提示", MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If

        strWLID = txtWorkingLayerID.SelectedValue


        strQryPFN = txtQryPath.Text
        If String.IsNullOrEmpty(strQryPFN) Then
            MessageBox.Show("请先输入Query相对项目的PathfullName", "提示"， MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        Dim Proj As Plt.IComosDProject
        If Workset.GetAllProjects.ItemExist(strPrjName) Then
            Proj = Workset.GetAllProjects.Item(strPrjName)
        Else
            MessageBox.Show(“COMOS中未找到项目号：” + strPrjName, "提示"， MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        Call Workset.SetCurrentProject(Proj)
        If String.IsNullOrEmpty(strWLID) = False Then
            Dim oWl As Plt.IComosDWorkingOverlay
            oWl = Workset.GetCurrentProject().GetWorkingOverlay(strWLID)
            If oWl Is Nothing Then
                MessageBox.Show("工作层ID不正确，请勿输入工作层Name", "提示"， MessageBoxButton.OK, MessageBoxImage.Information)
                Exit Sub
            End If
            Workset.GetCurrentProject().CurrentWorkingOverlay = oWl
        End If
        Dim qryObj As Plt.IComosDDevice
        qryObj = Proj.GetObjectByPathFullName(strQryPFN)
        If qryObj Is Nothing Then
            MessageBox.Show(“项目” + strPrjName + "中未找到Query：" + vbCrLf + strQryPFN, "提示", MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        If qryObj.SystemType <> 8 Then
            MessageBox.Show("QryPathFullName对应的COMOS对象不是工程对象，请重新输入。", "提示", 64, MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        If qryObj.Class <> "J" Then
            MessageBox.Show("QryPathFullName对应的COMOS对象Query类型，请重新输入。", "提示", 64, MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        Dim topQry As ComosQueryInterface.ITopQuery
        Dim qry As ComosQueryInterface.IQuery
        Dim strResult As String
        topQry = qryObj.XObj.topQuery
        topQry.Execute()
        qry = topQry.Query
        strResult = ""
        '//输出Query查询到的结果
        For i = 1 To qry.RowCount
            For j = 1 To qry.BaseQuery.ColumnCount
                strResult = strResult + qry.Cell(i, j).Text + " "
            Next
            strResult += vbCrLf
        Next
        txtMsgInfo.Text = strResult
        '//输出Query查询到的结果
    End Sub
    Private Sub ConnectCOMOS_Click(sender As Object, e As RoutedEventArgs)  '测试数据库连接获取具有权限的全部项目名称
        txtMsgInfo.Text = "执行中..."
        Workset = Nothing
        Comos.Global.AppGlobal.Workset = Nothing
        If Workset Is Nothing Then
            Dim blnCheck As Boolean
            blnCheck = GetWorkset()
            If blnCheck = False Then
                txtMsgInfo.Text = "获取workset失败"
                Exit Sub
            End If
        End If
        'get project
        Dim Proj As Plt.IComosDProject
        Dim strMsgInfo As String
        strMsgInfo = ""

        For i = 1 To PrjCol.Count
            Proj = PrjCol.Item(i)

            strMsgInfo += CStr(i) + ". " + Proj.Name + Proj.Description + vbCrLf


        Next
        txtMsgInfo.Text = "COMOS项目连接成功,项目列表如下" + vbCrLf + strMsgInfo
    End Sub

    Function GetWorkset()
        GetWorkset = False

        Workset = Nothing

        If (Comos.Global.AppGlobal.Workset Is Nothing) Then
            Workset = New Plt.CPLTWorksetClass
            If Not Workset Is Nothing Then
                Comos.Global.AppGlobal.Workset = Workset
            End If
        End If

        Dim DBPath As String
        DBPath = txtDBPath.Text

        'check if a path has been entered
        If DBPath = "" Then
            MsgBox（"请输入SQL ODBC配置名称或数据库地址!" + vbCrLf + "SQL示例:[SQL - SERVER]pt_sql_server" + vbCrLf + "Access示例:C:\Comos.mdb" + vbCrLf + "Oracle示例:[ORACLE]pt_oracle" _
                , 64, "提示")
            Exit Function
        End If
        'Start instance of Comos
        'Try to open the database
        Dim msgErr As Boolean

        If (Not Workset.IsInitialized()) Then
            msgErr = Workset.Init(String.Empty, String.Empty, DBPath)


            If (Not Comos.Global.AppGlobal.Workset.IsInitialized()) Then

                msgErr = True

            Else

                msgErr = False

            End If

        End If

        'If Not Workset.Init(String.Empty, String.Empty, DBPath) And Not Workset.IsInitialized() Then

        If msgErr = True Then

            '    MsgBox("无法打开数据库:  " + vbCrLf + DBPath, 64, "提示")
            '    Exit Sub
        End If
        PrjCol = Workset.GetTempCollection '项目集合
        WLCol = Workset.GetTempCollection '工作层集合
        ' the setup user
        Dim strUserName As String
        strUserName = System.Environment.GetEnvironmentVariable("UserName").ToString()
        strUserName = UCase(strUserName) ' UCase(strUserName + "_2")


        CurUser = Workset.GetAllUsers.Item(strUserName)
        'Safety check if user exists or not
        If CurUser Is Nothing Then
            MsgBox（"用户" + strUserName + " 不存在,请联系COMOS管理员." + vbCrLf + DBPath, vbCritical + vbOKOnly)

            Exit Function
        End If
        'Set user
        Workset.SetCurrentUser(CurUser)
        Dim prjTemp As Plt.IComosDProject


        For i = 1 To Workset.GetAllProjects().Count
            prjTemp = Workset.GetAllProjects().Item(i)
            If clsStdMod.RightsCheck(prjTemp, 1) And prjTemp.Name.Substring(0, 1) <> "@" Then '只显示有权限的项目
                PrjCol.Add(prjTemp)
            End If
        Next

        Dim dic As Dictionary(Of String, String)
        dic = New Dictionary(Of String, String)

        For i = 1 To PrjCol.Count
            dic.Add(PrjCol.Item(i).name, PrjCol.Item(i).name + "-" + PrjCol.Item(i).Description)
        Next
        txtPrjName.ItemsSource = dic
        txtPrjName.SelectedValuePath = "Key"
        txtPrjName.DisplayMemberPath = "Value"
        GetWorkset = True
    End Function
    Function GetWokingLayers(oWL)
        Dim oSubWL As Plt.IComosDWorkingOverlay
        For i = 1 To oWL.WorkingOverlays.Count
            oSubWL = CType(oWL.WorkingOverlays.Item(i), Plt.IComosDWorkingOverlay)
            If clsStdMod.RightsCheck(oSubWL, 1) Then '如对工作层无权限则跳过
                WLCol.Add(oSubWL)
                ' dic.Add(oSubWL.ID, oSubWL.Name + " " + oSubWL.Description)
                GetWokingLayers(oSubWL)
            End If
        Next
    End Function

    Private Sub txtPrjName_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim strPrjName As String
        strPrjName = txtPrjName.SelectedValue
        If String.IsNullOrEmpty(strPrjName) Then
            Exit Sub
        End If
        Dim Proj As Plt.IComosDProject
        If Workset.GetAllProjects.ItemExist(strPrjName) Then
            Proj = Workset.GetAllProjects.Item(strPrjName)
        Else
            MessageBox.Show(“COMOS中未找到项目号：” + strPrjName, "提示"， MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub
        End If
        Call Workset.SetCurrentProject(Proj)
        GetWokingLayers(Proj)
        Dim dic As Dictionary(Of String, String)
        dic = New Dictionary(Of String, String)
        Dim oWl As Plt.IComosDWorkingOverlay
        For i = 1 To WLCol.Count
            oWl = WLCol.Item(i)
            dic.Add(oWl.ID, oWl.Name + " " + oWl.Description)
        Next
        txtWorkingLayerID.ItemsSource = dic
        txtWorkingLayerID.SelectedValuePath = "Key"
        txtWorkingLayerID.DisplayMemberPath = "Value"

    End Sub

    Private Sub txtDBPath_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Workset = Nothing
    End Sub

    Private Sub StackPanel_Loaded(sender As Object, e As RoutedEventArgs)
        Dim dic = New Dictionary(Of String, String)
        dic.Add("[SQL - SERVER]pt_sql_server", "[SQL - SERVER]pt_sql_server")
        dic.Add("[SQL - SERVER]pt_sql_server_1", "[SQL - SERVER]pt_sql_server_1")
        dic.Add("[SQL - SERVER]pt_sql_server_2", "[SQL - SERVER]pt_sql_server_2")
        txtDBPath.ItemsSource = dic
        txtDBPath.DisplayMemberPath = "Value"
        txtDBPath.SelectedValuePath = "Key"
    End Sub
End Class

