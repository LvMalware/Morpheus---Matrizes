<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ArquivoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NovaMatrizToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator
        Me.SalvarMatrizToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OpçõesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EscalonarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator
        Me.InversaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripSeparator
        Me.TranspostaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripSeparator
        Me.CofatoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripSeparator
        Me.DeterminanteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripSeparator
        Me.AdjuntaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripSeparator
        Me.SairToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MatrizesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.OrdemLabel = New System.Windows.Forms.ToolStripLabel
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ElemLabel = New System.Windows.Forms.ToolStripLabel
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.DetLabel = New System.Windows.Forms.ToolStripLabel
        Me.SobreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ArquivoToolStripMenuItem, Me.OpçõesToolStripMenuItem, Me.MatrizesToolStripMenuItem, Me.SobreToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(706, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ArquivoToolStripMenuItem
        '
        Me.ArquivoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NovaMatrizToolStripMenuItem, Me.ToolStripMenuItem1, Me.SalvarMatrizToolStripMenuItem, Me.ToolStripMenuItem7, Me.SairToolStripMenuItem})
        Me.ArquivoToolStripMenuItem.Name = "ArquivoToolStripMenuItem"
        Me.ArquivoToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.ArquivoToolStripMenuItem.Text = "Arquivo"
        '
        'NovaMatrizToolStripMenuItem
        '
        Me.NovaMatrizToolStripMenuItem.Name = "NovaMatrizToolStripMenuItem"
        Me.NovaMatrizToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.NovaMatrizToolStripMenuItem.Text = "Nova matriz"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(149, 6)
        '
        'SalvarMatrizToolStripMenuItem
        '
        Me.SalvarMatrizToolStripMenuItem.Name = "SalvarMatrizToolStripMenuItem"
        Me.SalvarMatrizToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SalvarMatrizToolStripMenuItem.Text = "Salvar Matriz"
        '
        'OpçõesToolStripMenuItem
        '
        Me.OpçõesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EscalonarToolStripMenuItem, Me.ToolStripMenuItem2, Me.InversaToolStripMenuItem, Me.ToolStripMenuItem3, Me.TranspostaToolStripMenuItem, Me.ToolStripMenuItem4, Me.CofatoresToolStripMenuItem, Me.ToolStripMenuItem5, Me.DeterminanteToolStripMenuItem, Me.ToolStripMenuItem6, Me.AdjuntaToolStripMenuItem})
        Me.OpçõesToolStripMenuItem.Name = "OpçõesToolStripMenuItem"
        Me.OpçõesToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.OpçõesToolStripMenuItem.Text = "Opções"
        '
        'EscalonarToolStripMenuItem
        '
        Me.EscalonarToolStripMenuItem.Name = "EscalonarToolStripMenuItem"
        Me.EscalonarToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.EscalonarToolStripMenuItem.Text = "Escalonar"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(149, 6)
        '
        'InversaToolStripMenuItem
        '
        Me.InversaToolStripMenuItem.Name = "InversaToolStripMenuItem"
        Me.InversaToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.InversaToolStripMenuItem.Text = "Inversa"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(149, 6)
        '
        'TranspostaToolStripMenuItem
        '
        Me.TranspostaToolStripMenuItem.Name = "TranspostaToolStripMenuItem"
        Me.TranspostaToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.TranspostaToolStripMenuItem.Text = "Transposta"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(149, 6)
        '
        'CofatoresToolStripMenuItem
        '
        Me.CofatoresToolStripMenuItem.Name = "CofatoresToolStripMenuItem"
        Me.CofatoresToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.CofatoresToolStripMenuItem.Text = "Cofatores"
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(149, 6)
        '
        'DeterminanteToolStripMenuItem
        '
        Me.DeterminanteToolStripMenuItem.Name = "DeterminanteToolStripMenuItem"
        Me.DeterminanteToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DeterminanteToolStripMenuItem.Text = "Determinante"
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(149, 6)
        '
        'AdjuntaToolStripMenuItem
        '
        Me.AdjuntaToolStripMenuItem.Name = "AdjuntaToolStripMenuItem"
        Me.AdjuntaToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AdjuntaToolStripMenuItem.Text = "Adjunta"
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(149, 6)
        '
        'SairToolStripMenuItem
        '
        Me.SairToolStripMenuItem.Name = "SairToolStripMenuItem"
        Me.SairToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SairToolStripMenuItem.Text = "Sair"
        '
        'MatrizesToolStripMenuItem
        '
        Me.MatrizesToolStripMenuItem.CheckOnClick = True
        Me.MatrizesToolStripMenuItem.Name = "MatrizesToolStripMenuItem"
        Me.MatrizesToolStripMenuItem.Size = New System.Drawing.Size(63, 20)
        Me.MatrizesToolStripMenuItem.Text = "Matrizes"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.PictureBox1.Location = New System.Drawing.Point(114, 75)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(459, 323)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OrdemLabel, Me.ToolStripSeparator1, Me.ElemLabel, Me.ToolStripSeparator2, Me.DetLabel})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 419)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(706, 25)
        Me.ToolStrip1.TabIndex = 3
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'OrdemLabel
        '
        Me.OrdemLabel.Name = "OrdemLabel"
        Me.OrdemLabel.Size = New System.Drawing.Size(56, 22)
        Me.OrdemLabel.Text = "Ordem: 0"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ElemLabel
        '
        Me.ElemLabel.Name = "ElemLabel"
        Me.ElemLabel.Size = New System.Drawing.Size(74, 22)
        Me.ElemLabel.Text = "Elementos: 0"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'DetLabel
        '
        Me.DetLabel.Name = "DetLabel"
        Me.DetLabel.Size = New System.Drawing.Size(90, 22)
        Me.DetLabel.Text = "Determinante: ?"
        '
        'SobreToolStripMenuItem
        '
        Me.SobreToolStripMenuItem.Name = "SobreToolStripMenuItem"
        Me.SobreToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.SobreToolStripMenuItem.Text = "Sobre"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(706, 444)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Matrizes"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ArquivoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NovaMatrizToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SalvarMatrizToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SairToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpçõesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EscalonarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents InversaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TranspostaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CofatoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents DeterminanteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AdjuntaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MatrizesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents OrdemLabel As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ElemLabel As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents DetLabel As System.Windows.Forms.ToolStripLabel
    Friend WithEvents SobreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
