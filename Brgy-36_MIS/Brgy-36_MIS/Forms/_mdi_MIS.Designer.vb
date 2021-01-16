<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _mdi_MIS
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_mdi_MIS))
        Me.mnuMIS = New System.Windows.Forms.MenuStrip()
        Me.mnuResidents = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewResidentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddResidentToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOfficials = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewOfficialsRecordsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UpdateOfficialsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCases = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComplaintsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewComplaintsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileComplaintsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncidentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewIncidentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportIncidentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BlotterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewBlotterrsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileBlotterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCertificates = New System.Windows.Forms.ToolStripMenuItem()
        Me.BarangayClearanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BarangayIndigencyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BarangayCertificatesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FireIncidentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.COVID19RelatedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInventories = New System.Windows.Forms.ToolStripMenuItem()
        Me.EquipmentsPropertiesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuReports = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMapping = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMaintenance = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLogout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMIS.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMIS
        '
        Me.mnuMIS.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuResidents, Me.mnuOfficials, Me.mnuCases, Me.mnuCertificates, Me.mnuInventories, Me.mnuReports, Me.mnuMapping, Me.mnuMaintenance, Me.mnuLogout})
        Me.mnuMIS.Location = New System.Drawing.Point(0, 0)
        Me.mnuMIS.Name = "mnuMIS"
        Me.mnuMIS.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.mnuMIS.Size = New System.Drawing.Size(737, 24)
        Me.mnuMIS.TabIndex = 5
        Me.mnuMIS.Text = "mnuMIS"
        '
        'mnuResidents
        '
        Me.mnuResidents.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewResidentToolStripMenuItem, Me.AddResidentToolStripMenuItem1})
        Me.mnuResidents.ImageTransparentColor = System.Drawing.SystemColors.ActiveBorder
        Me.mnuResidents.Name = "mnuResidents"
        Me.mnuResidents.Size = New System.Drawing.Size(80, 20)
        Me.mnuResidents.Text = "&RESIDENTS"
        '
        'ViewResidentToolStripMenuItem
        '
        Me.ViewResidentToolStripMenuItem.Name = "ViewResidentToolStripMenuItem"
        Me.ViewResidentToolStripMenuItem.Size = New System.Drawing.Size(192, 22)
        Me.ViewResidentToolStripMenuItem.Text = "View Resident's Records"
        '
        'AddResidentToolStripMenuItem1
        '
        Me.AddResidentToolStripMenuItem1.Name = "AddResidentToolStripMenuItem1"
        Me.AddResidentToolStripMenuItem1.Size = New System.Drawing.Size(192, 22)
        Me.AddResidentToolStripMenuItem1.Text = "Add Resident"
        '
        'mnuOfficials
        '
        Me.mnuOfficials.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewOfficialsRecordsToolStripMenuItem, Me.UpdateOfficialsToolStripMenuItem})
        Me.mnuOfficials.Name = "mnuOfficials"
        Me.mnuOfficials.Size = New System.Drawing.Size(74, 20)
        Me.mnuOfficials.Text = "&OFFICIALS"
        '
        'ViewOfficialsRecordsToolStripMenuItem
        '
        Me.ViewOfficialsRecordsToolStripMenuItem.Name = "ViewOfficialsRecordsToolStripMenuItem"
        Me.ViewOfficialsRecordsToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.ViewOfficialsRecordsToolStripMenuItem.Text = "View Official's Records"
        '
        'UpdateOfficialsToolStripMenuItem
        '
        Me.UpdateOfficialsToolStripMenuItem.Name = "UpdateOfficialsToolStripMenuItem"
        Me.UpdateOfficialsToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.UpdateOfficialsToolStripMenuItem.Text = "Update Officials"
        '
        'mnuCases
        '
        Me.mnuCases.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ComplaintsToolStripMenuItem, Me.IncidentsToolStripMenuItem, Me.BlotterToolStripMenuItem})
        Me.mnuCases.Name = "mnuCases"
        Me.mnuCases.Size = New System.Drawing.Size(55, 20)
        Me.mnuCases.Text = "&CASES"
        '
        'ComplaintsToolStripMenuItem
        '
        Me.ComplaintsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewComplaintsToolStripMenuItem, Me.FileComplaintsToolStripMenuItem})
        Me.ComplaintsToolStripMenuItem.Name = "ComplaintsToolStripMenuItem"
        Me.ComplaintsToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.ComplaintsToolStripMenuItem.Text = "Complaints"
        '
        'ViewComplaintsToolStripMenuItem
        '
        Me.ViewComplaintsToolStripMenuItem.Name = "ViewComplaintsToolStripMenuItem"
        Me.ViewComplaintsToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.ViewComplaintsToolStripMenuItem.Text = "View Complaints"
        '
        'FileComplaintsToolStripMenuItem
        '
        Me.FileComplaintsToolStripMenuItem.Name = "FileComplaintsToolStripMenuItem"
        Me.FileComplaintsToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.FileComplaintsToolStripMenuItem.Text = "File Complaint"
        '
        'IncidentsToolStripMenuItem
        '
        Me.IncidentsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewIncidentsToolStripMenuItem, Me.ReportIncidentToolStripMenuItem})
        Me.IncidentsToolStripMenuItem.Name = "IncidentsToolStripMenuItem"
        Me.IncidentsToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.IncidentsToolStripMenuItem.Text = "Incidents"
        '
        'ViewIncidentsToolStripMenuItem
        '
        Me.ViewIncidentsToolStripMenuItem.Name = "ViewIncidentsToolStripMenuItem"
        Me.ViewIncidentsToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ViewIncidentsToolStripMenuItem.Text = "View Incidents"
        '
        'ReportIncidentToolStripMenuItem
        '
        Me.ReportIncidentToolStripMenuItem.Name = "ReportIncidentToolStripMenuItem"
        Me.ReportIncidentToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ReportIncidentToolStripMenuItem.Text = "Report Incident"
        '
        'BlotterToolStripMenuItem
        '
        Me.BlotterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewBlotterrsToolStripMenuItem, Me.FileBlotterToolStripMenuItem})
        Me.BlotterToolStripMenuItem.Name = "BlotterToolStripMenuItem"
        Me.BlotterToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.BlotterToolStripMenuItem.Text = "Blotters"
        '
        'ViewBlotterrsToolStripMenuItem
        '
        Me.ViewBlotterrsToolStripMenuItem.Name = "ViewBlotterrsToolStripMenuItem"
        Me.ViewBlotterrsToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.ViewBlotterrsToolStripMenuItem.Text = "View Blotterrs"
        '
        'FileBlotterToolStripMenuItem
        '
        Me.FileBlotterToolStripMenuItem.Name = "FileBlotterToolStripMenuItem"
        Me.FileBlotterToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.FileBlotterToolStripMenuItem.Text = "File Blotter"
        '
        'mnuCertificates
        '
        Me.mnuCertificates.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BarangayClearanceToolStripMenuItem, Me.BarangayIndigencyToolStripMenuItem, Me.BarangayCertificatesToolStripMenuItem})
        Me.mnuCertificates.Name = "mnuCertificates"
        Me.mnuCertificates.Size = New System.Drawing.Size(94, 20)
        Me.mnuCertificates.Text = "C&ERTIFICATES"
        '
        'BarangayClearanceToolStripMenuItem
        '
        Me.BarangayClearanceToolStripMenuItem.Name = "BarangayClearanceToolStripMenuItem"
        Me.BarangayClearanceToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.BarangayClearanceToolStripMenuItem.Text = "Barangay Clearance"
        '
        'BarangayIndigencyToolStripMenuItem
        '
        Me.BarangayIndigencyToolStripMenuItem.Name = "BarangayIndigencyToolStripMenuItem"
        Me.BarangayIndigencyToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.BarangayIndigencyToolStripMenuItem.Text = "Barangay Indigency"
        '
        'BarangayCertificatesToolStripMenuItem
        '
        Me.BarangayCertificatesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FireIncidentToolStripMenuItem, Me.COVID19RelatedToolStripMenuItem})
        Me.BarangayCertificatesToolStripMenuItem.Name = "BarangayCertificatesToolStripMenuItem"
        Me.BarangayCertificatesToolStripMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.BarangayCertificatesToolStripMenuItem.Text = "Barangay Certificates"
        '
        'FireIncidentToolStripMenuItem
        '
        Me.FireIncidentToolStripMenuItem.Name = "FireIncidentToolStripMenuItem"
        Me.FireIncidentToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.FireIncidentToolStripMenuItem.Text = "Fire Incident"
        '
        'COVID19RelatedToolStripMenuItem
        '
        Me.COVID19RelatedToolStripMenuItem.Name = "COVID19RelatedToolStripMenuItem"
        Me.COVID19RelatedToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.COVID19RelatedToolStripMenuItem.Text = "COVID-19 Related"
        '
        'mnuInventories
        '
        Me.mnuInventories.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EquipmentsPropertiesToolStripMenuItem})
        Me.mnuInventories.Name = "mnuInventories"
        Me.mnuInventories.Size = New System.Drawing.Size(92, 20)
        Me.mnuInventories.Text = "&INVENTORIES"
        '
        'EquipmentsPropertiesToolStripMenuItem
        '
        Me.EquipmentsPropertiesToolStripMenuItem.Name = "EquipmentsPropertiesToolStripMenuItem"
        Me.EquipmentsPropertiesToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.EquipmentsPropertiesToolStripMenuItem.Text = "Equipments/Properties"
        '
        'mnuReports
        '
        Me.mnuReports.Name = "mnuReports"
        Me.mnuReports.Size = New System.Drawing.Size(70, 20)
        Me.mnuReports.Text = "&REPORTS"
        '
        'mnuMapping
        '
        Me.mnuMapping.Name = "mnuMapping"
        Me.mnuMapping.Size = New System.Drawing.Size(70, 20)
        Me.mnuMapping.Text = "&MAPPING"
        '
        'mnuMaintenance
        '
        Me.mnuMaintenance.Name = "mnuMaintenance"
        Me.mnuMaintenance.Size = New System.Drawing.Size(100, 20)
        Me.mnuMaintenance.Text = "M&AINTENANCE"
        '
        'mnuLogout
        '
        Me.mnuLogout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.mnuLogout.Name = "mnuLogout"
        Me.mnuLogout.Size = New System.Drawing.Size(66, 20)
        Me.mnuLogout.Text = "&LOGOUT"
        '
        '_mdi_MIS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.BackgroundImage = Global.Brgy_36_MIS.My.Resources.Resources.Brgy_36
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(737, 558)
        Me.Controls.Add(Me.mnuMIS)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Century Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.mnuMIS
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "_mdi_MIS"
        Me.Text = "MANAGEMENT INFORMATION SYSTEM  of BARANGAY 36-A  (SAP-SAP)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.mnuMIS.ResumeLayout(False)
        Me.mnuMIS.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuResidents As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMIS As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuOfficials As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCases As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCertificates As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuInventories As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuReports As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMapping As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMaintenance As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLogout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewResidentToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddResidentToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewOfficialsRecordsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateOfficialsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ComplaintsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewComplaintsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileComplaintsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IncidentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewIncidentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportIncidentToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BlotterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewBlotterrsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileBlotterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BarangayClearanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BarangayIndigencyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BarangayCertificatesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FireIncidentToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents COVID19RelatedToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EquipmentsPropertiesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
