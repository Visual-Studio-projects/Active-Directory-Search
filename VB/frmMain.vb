
#Region "  Authors / Credit  "
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              This application uses LDAP (Lightweight Directory Access Protocol) to access Active Directory
' References:           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=5006&lngWId=10
'
' Ver.  Date            Author              Details
' 1.00  18-AUG-2006     Joseph Beeler       Initial version.
' 1.01  12-APR-2010     Anthony Duguid      formatting changes and save file function without the need for the Excel dll
'--------------------------------------------------------------------------------------------------------------------
#End Region

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.DirectoryServices
Imports System.IO.StreamWriter
'Imports ActiveDs 'commented out so I wouldn't have to reference Interop.ActiveDs.dll anymore

Public Class frmMain
    Inherits System.Windows.Forms.Form
    Private deBase As DirectoryEntry
    Public listboxSortOrder As Integer = 0  'used in the sort of listboxes
    Public exportFileName As String

    Class ListViewItemComparer
        Implements IComparer
        Private col As Integer
        Private order As SortOrder

        Public Sub New()
            col = 0
            order = SortOrder.Ascending
        End Sub

        Public Sub New(ByVal column As Integer, ByVal order As SortOrder)
            col = column
            Me.order = order
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements System.Collections.IComparer.Compare
            Dim returnVal As Integer = -1
            returnVal = [String].Compare(CType(x,
                            ListViewItem).SubItems(col).Text,
                            CType(y, ListViewItem).SubItems(col).Text)
            ' Determine whether the sort order is descending.
            If order = SortOrder.Descending Then
                ' Invert the value returned by String.Compare.
                returnVal *= -1
            End If

            Return returnVal
        End Function

    End Class

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents tlbActiveDirectory As System.Windows.Forms.ToolStrip
    Friend WithEvents tltControlTips As System.Windows.Forms.ToolTip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents tlpMain As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents IltGroups As System.Windows.Forms.ImageList
    Friend WithEvents tvwGroups As System.Windows.Forms.TreeView
    Friend WithEvents lvwAttributes As System.Windows.Forms.ListView
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbSave As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbSearch As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbUserList As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsp1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsp2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsp3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFind As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tlpMain = New System.Windows.Forms.TableLayoutPanel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.tvwGroups = New System.Windows.Forms.TreeView()
        Me.IltGroups = New System.Windows.Forms.ImageList(Me.components)
        Me.lvwAttributes = New System.Windows.Forms.ListView()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.tlbActiveDirectory = New System.Windows.Forms.ToolStrip()
        Me.tsbExit = New System.Windows.Forms.ToolStripButton()
        Me.tsp1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSave = New System.Windows.Forms.ToolStripButton()
        Me.tsp2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSearch = New System.Windows.Forms.ToolStripButton()
        Me.tsp3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbUserList = New System.Windows.Forms.ToolStripButton()
        Me.mnuMain = New System.Windows.Forms.MenuStrip()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.tltControlTips = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.tlpMain.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.tlbActiveDirectory.SuspendLayout()
        Me.mnuMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.tlpMain)
        Me.Panel1.Controls.Add(Me.Splitter1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(903, 698)
        Me.Panel1.TabIndex = 0
        '
        'tlpMain
        '
        Me.tlpMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tlpMain.ColumnCount = 3
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tlpMain.Controls.Add(Me.SplitContainer1, 1, 1)
        Me.tlpMain.Location = New System.Drawing.Point(3, 48)
        Me.tlpMain.Name = "tlpMain"
        Me.tlpMain.RowCount = 3
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.tlpMain.Size = New System.Drawing.Size(897, 650)
        Me.tlpMain.TabIndex = 7
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(13, 13)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.tvwGroups)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lvwAttributes)
        Me.SplitContainer1.Size = New System.Drawing.Size(871, 624)
        Me.SplitContainer1.SplitterDistance = 310
        Me.SplitContainer1.TabIndex = 8
        '
        'tvwGroups
        '
        Me.tvwGroups.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvwGroups.ImageIndex = 0
        Me.tvwGroups.ImageList = Me.IltGroups
        Me.tvwGroups.Location = New System.Drawing.Point(0, 0)
        Me.tvwGroups.Name = "tvwGroups"
        Me.tvwGroups.SelectedImageIndex = 0
        Me.tvwGroups.Size = New System.Drawing.Size(306, 620)
        Me.tvwGroups.TabIndex = 7
        '
        'IltGroups
        '
        Me.IltGroups.ImageStream = CType(resources.GetObject("IltGroups.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.IltGroups.TransparentColor = System.Drawing.Color.Transparent
        Me.IltGroups.Images.SetKeyName(0, "folder.png")
        Me.IltGroups.Images.SetKeyName(1, "folder_add.png")
        Me.IltGroups.Images.SetKeyName(2, "Help2.png")
        Me.IltGroups.Images.SetKeyName(3, "world.png")
        Me.IltGroups.Images.SetKeyName(4, "monitor.png")
        Me.IltGroups.Images.SetKeyName(5, "user.png")
        Me.IltGroups.Images.SetKeyName(6, "group.png")
        '
        'lvwAttributes
        '
        Me.lvwAttributes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvwAttributes.FullRowSelect = True
        Me.lvwAttributes.GridLines = True
        Me.lvwAttributes.Location = New System.Drawing.Point(0, 0)
        Me.lvwAttributes.MultiSelect = False
        Me.lvwAttributes.Name = "lvwAttributes"
        Me.lvwAttributes.Size = New System.Drawing.Size(553, 620)
        Me.lvwAttributes.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvwAttributes.TabIndex = 6
        Me.tltControlTips.SetToolTip(Me.lvwAttributes, "Double-Click to view values from either member or memberOf")
        Me.lvwAttributes.UseCompatibleStateImageBehavior = False
        Me.lvwAttributes.View = System.Windows.Forms.View.Details
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 698)
        Me.Splitter1.TabIndex = 5
        Me.Splitter1.TabStop = False
        '
        'tlbActiveDirectory
        '
        Me.tlbActiveDirectory.AutoSize = False
        Me.tlbActiveDirectory.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.tsp1, Me.tsbSave, Me.tsp2, Me.tsbSearch, Me.tsp3, Me.tsbUserList})
        Me.tlbActiveDirectory.Location = New System.Drawing.Point(0, 0)
        Me.tlbActiveDirectory.Name = "tlbActiveDirectory"
        Me.tlbActiveDirectory.Size = New System.Drawing.Size(903, 45)
        Me.tlbActiveDirectory.TabIndex = 111
        Me.tlbActiveDirectory.Text = "tlbActiveDirectory"
        '
        'tsbExit
        '
        Me.tsbExit.Image = Global.ADS.My.Resources.Resources.cross
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(50, 42)
        Me.tsbExit.Text = "    Exit   "
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbExit.ToolTipText = "Exit  (Ctrl + End)"
        '
        'tsp1
        '
        Me.tsp1.Name = "tsp1"
        Me.tsp1.Size = New System.Drawing.Size(6, 45)
        '
        'tsbSave
        '
        Me.tsbSave.Image = Global.ADS.My.Resources.Resources.disk
        Me.tsbSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbSave.Name = "tsbSave"
        Me.tsbSave.Size = New System.Drawing.Size(53, 42)
        Me.tsbSave.Text = "   Save   "
        Me.tsbSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbSave.ToolTipText = "Save file to desktop (Ctrl + S)"
        '
        'tsp2
        '
        Me.tsp2.Name = "tsp2"
        Me.tsp2.Size = New System.Drawing.Size(6, 45)
        '
        'tsbSearch
        '
        Me.tsbSearch.Image = Global.ADS.My.Resources.Resources.zoom
        Me.tsbSearch.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbSearch.Name = "tsbSearch"
        Me.tsbSearch.Size = New System.Drawing.Size(58, 42)
        Me.tsbSearch.Text = "    Find    "
        Me.tsbSearch.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbSearch.ToolTipText = "Find anything in Active Directory (Ctrl + F)"
        '
        'tsp3
        '
        Me.tsp3.Name = "tsp3"
        Me.tsp3.Size = New System.Drawing.Size(6, 45)
        '
        'tsbUserList
        '
        Me.tsbUserList.Image = Global.ADS.My.Resources.Resources.vcard
        Me.tsbUserList.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbUserList.Name = "tsbUserList"
        Me.tsbUserList.Size = New System.Drawing.Size(61, 42)
        Me.tsbUserList.Text = " User List "
        Me.tsbUserList.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbUserList.ToolTipText = "Show all users"
        Me.tsbUserList.Visible = False
        '
        'mnuMain
        '
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuExit, Me.mnuSave, Me.mnuFind})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(554, 24)
        Me.mnuMain.TabIndex = 112
        Me.mnuMain.Text = "MenuStrip1"
        Me.mnuMain.Visible = False
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.[End]), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(37, 20)
        Me.mnuExit.Text = "Exit"
        '
        'mnuSave
        '
        Me.mnuSave.Name = "mnuSave"
        Me.mnuSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnuSave.Size = New System.Drawing.Size(43, 20)
        Me.mnuSave.Text = "Save"
        '
        'mnuFind
        '
        Me.mnuFind.Name = "mnuFind"
        Me.mnuFind.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.mnuFind.Size = New System.Drawing.Size(42, 20)
        Me.mnuFind.Text = "Find"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(903, 698)
        Me.Controls.Add(Me.tlbActiveDirectory)
        Me.Controls.Add(Me.mnuMain)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnuMain
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Active Directory Search"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.tlpMain.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.tlbActiveDirectory.ResumeLayout(False)
        Me.tlbActiveDirectory.PerformLayout()
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "  Form  "

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        SaveSetting(Application.ProductName, "Settings", "LastObject", Me.tvwGroups.SelectedNode.FullPath)
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        subRootNode()
        ExpandNode(Me.tvwGroups, GetSetting(Application.ProductName, "Settings", "LastObject", ""))
    End Sub

#End Region

#Region "  Toolbar  "

    Private Sub tsbExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Application.Exit()
    End Sub

    Private Sub tsbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSave.Click
        ExportListViewToCSV(exportFileName)
    End Sub

    Private Sub tsbFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSearch.Click
        FindItem()
    End Sub

    Private Sub tsbUserList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbUserList.Click
        FindAllObjectClass()
    End Sub

#End Region

#Region "  Listview  "

    Private Sub lvwAttributes_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvwAttributes.ColumnClick
        If e.Column <> listboxSortOrder Then
            ' Set the sort column to the new column.
            listboxSortOrder = e.Column
            ' Set the sort order to ascending by default.
            Me.lvwAttributes.Sorting = SortOrder.Ascending
        Else
            ' Determine what the last sort order was and change it.
            If Me.lvwAttributes.Sorting = SortOrder.Ascending Then
                Me.lvwAttributes.Sorting = SortOrder.Descending
            Else
                Me.lvwAttributes.Sorting = SortOrder.Ascending
            End If
        End If
        ' Call the sort method to manually sort.
        Me.lvwAttributes.Sort()
        ' Set the ListViewItemSorter property to a new ListViewItemComparer object.
        Me.lvwAttributes.ListViewItemSorter = New ListViewItemComparer(e.Column, Me.lvwAttributes.Sorting)

    End Sub

    Private Sub lvwAttributes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvwAttributes.DoubleClick
        Dim I As Integer
        For I = 0 To Me.lvwAttributes.SelectedItems.Count - 1
            Dim strName As String = lvwAttributes.SelectedItems(I).Text
            If strName = "memberOf" Or strName = "member" Then
                Dim strTemp As String = Me.lvwAttributes.SelectedItems(I).SubItems(1).Text.ToString
                FindItem(strTemp)
            End If
        Next

    End Sub

#End Region

#Region "  Treeview  "

    Private Sub tvwGroups_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvwGroups.BeforeExpand
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        If e.Node.ImageIndex > 3 Then Exit Sub
        Try
            If e.Node.GetNodeCount(False) = 1 And e.Node.Nodes(0).Text = "" Then
                e.Node.Nodes(0).Remove()
                subEnumerateChildren(e.Node)
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to expand " & e.Node.FullPath & ":" & ex.ToString)
        End Try
        If e.Node.GetNodeCount(False) > 0 Then
            e.Node.ImageIndex = 1
            e.Node.SelectedImageIndex = 1
        End If
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub tvwGroups_BeforeCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvwGroups.BeforeCollapse
        e.Node.ImageIndex = 0
        e.Node.SelectedImageIndex = 0
    End Sub

    Private Sub tvwGroups_BeforeSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvwGroups.AfterSelect
        Try
            Dim list As DirectoryEntry = DirectCast(e.Node.Tag, DirectoryEntry)
            exportFileName = Me.tvwGroups.SelectedNode.Text.ToString
            Me.lvwAttributes.Clear()
            Me.lvwAttributes.Columns.Add("Attribute", 150, HorizontalAlignment.Left)
            Me.lvwAttributes.Columns.Add("Format Value", 150, HorizontalAlignment.Left) 'used in the double click search
            Me.lvwAttributes.Columns.Add("Value", 475, HorizontalAlignment.Left)
            Me.lvwAttributes.Columns.Add("Name", 0, HorizontalAlignment.Left)
            For Each listIter As Object In list.Properties.PropertyNames
                For Each Iter As Object In list.Properties(listIter.ToString())
                    Dim item As ListViewItem = New ListViewItem(listIter.ToString(), 0)
                    'Needed in order to convert these types to a type that is viewable.
                    Select Case Iter.GetType.ToString
                        Case "System.__ComObject"
                            'Iter = AdsDateValue(Iter) 'removed 28-APR-2010 ALD - not sure if I need it.
                        Case "System.Byte[]"
                            Iter = Convert.ToBase64String(Iter)
                    End Select
                    Dim strAttribute As String = listIter.ToString()
                    Dim strValue As String = Iter.ToString()
                    Dim strTemp As String = ""
                    If strAttribute = "memberOf" Or strAttribute = "member" Then
                        Dim strItems() As String = Split(strValue, ",") 'load into an array
                        strTemp = strItems(0).Replace("CN=", "")
                    End If
                    item.SubItems.Add(strTemp)
                    item.SubItems.Add(strValue)
                    item.SubItems.Add(exportFileName)
                    Me.lvwAttributes.Items.AddRange(New ListViewItem() {item})
                Next
            Next
            'update form title
            Me.Text = Application.ProductName & " (" & exportFileName & ")"
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

#End Region

#Region "  Menubar  "

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub

    Private Sub mnuFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFind.Click
        FindItem()
    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        ExportListViewToCSV(exportFileName)
    End Sub

#End Region

#Region "  Functions  "

    Public Function AdsDateValue(ByVal oV As Object) As Date
        '----------------------------------------------
        ' Purpose: converting object to date
        '----------------------------------------------
        Try
            'Dim v As IADsLargeInteger = DirectCast(oV, IADsLargeInteger)
            'Dim dV As Long = Convert.ToInt64(v.HighPart) * 4294967296 + Convert.ToInt64(v.LowPart)
            'Return Date.FromFileTime(dV)
        Catch
            Return Date.MinValue
        End Try
    End Function

    Public Function DomainName() As String
        Dim objRootDSE As New DirectoryEntry("LDAP://RootDSE")
        DomainName = objRootDSE.Properties("defaultNamingContext")(0).ToString
    End Function

    Public Function ExportListViewToCSV(ByVal pstrName As String) As Boolean
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose:              Export the contents of a listview to a .csv file
        ' Notes:                http://dotnetref.blogspot.com/2008/03/exporting-listview-to-csv-file-in-vbnet.html
        '
        ' Ver.  Date            Author              Details
        ' 1.00  23-MAR-2008     srego               Initial version.
        ' 1.01  12-APR-2010     Anthony Duguid      formatting changes
        '--------------------------------------------------------------------------------------------------------------------
        Dim strPath As String = My.Computer.FileSystem.SpecialDirectories.Desktop.ToString
        Dim strFileName As String = strPath & "\ActiveDirectory_" & pstrName & ".csv"
        Dim lv As ListView = Me.lvwAttributes
        Try
            ' Open output file  
            Dim os As New System.IO.StreamWriter(strFileName)

            ' Write Headers  
            For i As Integer = 0 To lv.Columns.Count - 1
                ' replace quotes with double quotes if necessary  
                os.Write("""" & lv.Columns(i).Text.Replace("""", """""") & """,")
            Next
            os.WriteLine() 'create headers

            ' Write records 
            For i As Integer = 0 To lv.Items.Count - 1
                'If lv.Items(i).SubItems(0).Text = "member" Or lv.Items(i).SubItems(0).Text = "memberOf" Then
                For j As Integer = 0 To lv.Columns.Count - 1
                    os.Write(lv.Items(i).SubItems(j).Text.Replace(",", ";") & ",")
                Next
                os.WriteLine()
                'End If
            Next
            os.Close()
            Process.Start(strFileName) 'open file

        Catch ex As Exception
            ' catch any errors  
            Return (False)
        End Try

        Return (True)

    End Function

#End Region

#Region "  Subroutines  "

    Public Sub subRootNode()
        Try
            deBase = New DirectoryEntry("LDAP://" & DomainName())
            Dim oRootNode As TreeNode = Me.tvwGroups.Nodes.Add(deBase.Path)
            oRootNode.ImageIndex = 0
            oRootNode.SelectedImageIndex = 0
            oRootNode.Tag = deBase
            oRootNode.Text = deBase.Name.Substring(3)
            oRootNode.Nodes.Add("")
        Catch ex As Exception
            MessageBox.Show("Cannot create initial node: " & ex.ToString)
            End
        End Try
    End Sub

    Public Sub subEnumerateChildren(ByVal oRootNode As TreeNode)
        Dim deParent As DirectoryEntry = DirectCast(oRootNode.Tag, DirectoryEntry)
        For Each deChild As DirectoryEntry In deParent.Children
            Dim oChildNode As TreeNode = oRootNode.Nodes.Add(deChild.Path)
            Select Case deChild.SchemaClassName
                Case "computer"
                    oChildNode.ImageIndex = 4
                    oChildNode.SelectedImageIndex = 4
                Case "user"
                    oChildNode.ImageIndex = 5
                    oChildNode.SelectedImageIndex = 5
                Case "group"
                    oChildNode.ImageIndex = 6
                    oChildNode.SelectedImageIndex = 6
                Case "organizationalUnit"
                    oChildNode.ImageIndex = 0
                    oChildNode.SelectedImageIndex = 0
                    oChildNode.Nodes.Add("")
                Case "container"
                    oChildNode.ImageIndex = 0
                    oChildNode.SelectedImageIndex = 0
                    oChildNode.Nodes.Add("")
                Case "publicFolder"
                    oChildNode.ImageIndex = 3
                    oChildNode.SelectedImageIndex = 3
                    oChildNode.Nodes.Add("")
                Case Else
                    oChildNode.ImageIndex = 2
                    oChildNode.SelectedImageIndex = 2
                    oChildNode.Nodes.Add("")
            End Select
            oChildNode.Tag = deChild
            oChildNode.Text = deChild.Name.Substring(3).Replace("\", "")
        Next
    End Sub

    Public Sub ExpandNode(ByVal objTreeView As TreeView, ByVal FullNodePath As String)
        Dim i As Int32 = 0
        Dim buf As String = ""
        Dim pathSep As String = objTreeView.PathSeparator
        With objTreeView
            For i = 0 To .Nodes.Count - 1
                buf = FullNodePath.Split(pathSep.ToCharArray)(0)
                If .Nodes(i).Text.Substring(0, buf.Length) = FullNodePath.Split(pathSep.ToCharArray)(0) Then
                    OpenNodes(.Nodes(i), FullNodePath, pathSep, 1)
                    Exit Sub
                End If
            Next
        End With

    End Sub

    Public Sub FindItem(Optional ByVal pstrFind As String = "")
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose:              search for anything in active directory
        '
        ' Ver.  Date            Author              Details
        ' 1.00  18-AUG-2006     Joseph Beeler       Initial version.
        ' 1.01  12-APR-2010     Anthony Duguid      added optional variable and formatting
        '--------------------------------------------------------------------------------------------------------------------
        Dim MyFind As String
        If pstrFind = "" Then
            MyFind = InputBox("Enter the Active Directory object you wish to find?", "Find")
            If MyFind = "" Then Exit Sub
        Else
            MyFind = pstrFind
        End If

        Dim objSearch As New DirectorySearcher
        Dim QueryResults As SearchResult
        objSearch.SearchRoot = New DirectoryEntry("LDAP://" & DomainName())
        objSearch.SearchScope = SearchScope.Subtree
        objSearch.PropertiesToLoad.Add("cn")
        objSearch.Filter = "(sAMAccountName=" & MyFind & ")"
        QueryResults = objSearch.FindOne()
        If QueryResults Is Nothing Then
            objSearch.Filter = "(sAMAccountName=" & MyFind & "$)"
            QueryResults = objSearch.FindOne()
        End If
        If QueryResults Is Nothing Then
            objSearch.Filter = "(SN=" & MyFind & ")"
            QueryResults = objSearch.FindOne()
        End If
        If QueryResults Is Nothing Then
            objSearch.Filter = "(cn=" & MyFind & ")"
            QueryResults = objSearch.FindOne()
        End If
        If QueryResults Is Nothing Then
            MessageBox.Show("Nothing found matching your search criteria!", "Nothing Found!", MessageBoxButtons.OK)
            Exit Sub
        End If
        'Format path to format needed by the ExpandNode procedure.
        Dim NodeToFind() As String = Split(QueryResults.Path, ",")
        Array.Reverse(NodeToFind)
        Dim FindIt As String = objSearch.SearchRoot.Name.Substring(3)
        For X As Integer = 0 To UBound(NodeToFind)
            Dim test1 As String = NodeToFind(X).Substring(0, 2)
            If test1 = "OU" Then
                FindIt = FindIt + "\" + NodeToFind(X).Substring(3)
            End If
        Next
        Dim strCN As String = QueryResults.Properties.Item("CN")(0).ToString
        If strCN Is Nothing = False Then
            FindIt = FindIt + "\" + strCN
        End If
        ExpandNode(Me.tvwGroups, FindIt)
    End Sub

    Public Sub FindAllObjectClass(Optional ByVal pstrSearch As String = "User")
        Dim strCn As String
        Dim strMail As String
        Dim strSamAccountName As String
        Dim strOffice As String
        Dim strTitle As String
        Dim strPhone As String
        Dim strLastUpdated As String
        Dim lvwListView As ListView = Me.lvwAttributes
        Dim itmListItem As ListViewItem
        Dim results As SearchResultCollection = Nothing
        Try
            'update form title
            exportFileName = pstrSearch & " List"
            Me.Text = Application.ProductName & " (" & exportFileName & ")"
            Me.Cursor = Cursors.WaitCursor
            ' Bind to the users container.
            Dim entry As New DirectoryEntry("LDAP://" & DomainName())

            ' Create a DirectorySearcher object.
            Dim mySearcher As New DirectorySearcher(entry)

            ' Set a filter for users with the countryCode 
            'mySearcher.Filter = "(&(objectClass=" & pstrSearch & ")(countryCode=36))"
            'mySearcher.Filter = "(&(objectClass=user))"

            ' Use the FindAll method to return objects to a SearchResultCollection.
            results = mySearcher.FindAll()

            'create new columns for list
            lvwListView.Clear()
            lvwListView.Columns.Add("Full Name", 175, HorizontalAlignment.Left)
            lvwListView.Columns.Add("User Id", 100, HorizontalAlignment.Left)
            lvwListView.Columns.Add("Email", 225, HorizontalAlignment.Left)
            lvwListView.Columns.Add("Office", 150, HorizontalAlignment.Left)
            lvwListView.Columns.Add("Title", 200, HorizontalAlignment.Left)
            lvwListView.Columns.Add("Phone Number", 150, HorizontalAlignment.Left)
            lvwListView.Columns.Add("Last Updated", 150, HorizontalAlignment.Left)

            ' Iterate through each SearchResult in the SearchResultCollection.
            Dim searchResult As SearchResult
            For Each searchResult In results
                'get values from active directory search
                strCn = searchResult.GetDirectoryEntry().Properties("cn").Value
                strSamAccountName = searchResult.GetDirectoryEntry().Properties("sAMAccountName").Value
                strMail = searchResult.GetDirectoryEntry().Properties("mail").Value
                strOffice = searchResult.GetDirectoryEntry().Properties("physicaldeliveryofficename").Value
                strTitle = searchResult.GetDirectoryEntry().Properties("title").Value
                strPhone = searchResult.GetDirectoryEntry().Properties("telephoneNumber").Value
                strLastUpdated = searchResult.GetDirectoryEntry().Properties("whenchanged").Value
                'create line in listview
                itmListItem = New ListViewItem()
                itmListItem.Text = strCn
                itmListItem.SubItems.Add(strSamAccountName)
                itmListItem.SubItems.Add(strMail)
                itmListItem.SubItems.Add(strOffice)
                itmListItem.SubItems.Add(strTitle)
                itmListItem.SubItems.Add(strPhone)
                itmListItem.SubItems.Add(strLastUpdated)
                lvwListView.Items.Add(itmListItem)
                lvwListView.Refresh()
                itmListItem = Nothing
            Next searchResult

        Finally
            ' To prevent memory leaks, always call 
            'SearchResultCollection.Dispose() 'manually.
            If Not results Is Nothing Then
                results.Dispose()
                results = Nothing
            End If
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Public Sub OpenNodes(ByVal iNode As TreeNode, ByVal nPath As String,
           ByVal nPathSep As String, ByVal dirIndex As Int32)
        Dim i As Int32 = 0
        Dim pLen As Int32 = 0
        Dim buf As String = ""
        iNode.Expand()
        iNode.EnsureVisible()
        pLen = nPath.Split(nPathSep.ToCharArray).Length - 1
        If pLen >= dirIndex Then
            For i = 0 To dirIndex
                buf &= nPath.Split(nPathSep.ToCharArray)(i) & nPathSep
                Application.DoEvents()
            Next
            Try
                For i = 0 To iNode.Nodes.Count - 1
                    If iNode.Nodes(i).FullPath & nPathSep = buf Then
                        OpenNodes(iNode.Nodes(i), nPath, nPathSep, dirIndex + 1)
                        Exit Sub
                    End If
                    Application.DoEvents()
                Next
            Catch
            End Try
        Else
            iNode.TreeView.SelectedNode = iNode
        End If
    End Sub

#End Region

End Class
