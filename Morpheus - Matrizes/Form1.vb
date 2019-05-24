Public Class Form1
    Dim MatrizAtual As Matriz
    Public Matrizes As New Collections.Generic.Dictionary(Of String, Matriz)
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub NovaMatrizToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NovaMatrizToolStripMenuItem.Click
        With New OrdemDialogo
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim M As New Matriz(.Linhas.Value, .Colunas.Value)
                For i = 1 To M.Linhas
                    For j = 1 To M.Colunas
                        M.DefineItem(i, j, InputBox("Informe o valor para " & .Nome.Text & i & j))
                    Next
                Next
                MatrizAtual = M
                Matrizes.Add(.Nome.Text, M)
                AddHandler MatrizesToolStripMenuItem.DropDownItems.Add(.Nome.Text).Click, AddressOf MATRIZCLICK
                PictureBox1.Image = M.Desenhar(PictureBox1.Height, PictureBox1.Width)
                OrdemLabel.Text = "Ordem: " & M.Ordem
                ElemLabel.Text = "Elementos: " & M.Linhas * M.Colunas
                If M.Linhas = M.Colunas Then
                    DetLabel.Text = "Determinante: " & M.Determinante
                Else
                    DetLabel.Text = "Determinante: Desconhecida"
                End If
            End If
        End With
    End Sub
    Sub MatrizClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MatrizAtual = Matrizes(sender.ToString)
        PictureBox1.Image = MatrizAtual.Desenhar(PictureBox1.Height, PictureBox1.Width)
        OrdemLabel.Text = "Ordem: " & MatrizAtual.Ordem
        ElemLabel.Text = "Elementos: " & MatrizAtual.Linhas * MatrizAtual.Colunas
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            DetLabel.Text = "Determinante: " & MatrizAtual.Determinante
        Else
            DetLabel.Text = "Determinante: Desconhecida"
        End If
    End Sub

    Private Sub EscalonarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EscalonarToolStripMenuItem.Click
        If MatrizAtual.Colunas <= MatrizAtual.Linhas Then
            MsgBox("Apenas matrizes aplicadas (sistemas de equações lineares) são suportadas.", MsgBoxStyle.Information, "Ops!")
            Exit Sub
        End If
        PictureBox1.Image = MatrizAtual.Escalonada.Desenhar(PictureBox1.Height, PictureBox1.Width)
    End Sub

    Private Sub InversaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InversaToolStripMenuItem.Click
        If MatrizAtual.Linhas <> MatrizAtual.Colunas Then
            MsgBox("Apenas matrizes quadradas são suportadas.", MsgBoxStyle.Information, "Ops!")
            Exit Sub
        End If


        PictureBox1.Image = MatrizAtual.Inversa.Desenhar(PictureBox1.Height, PictureBox1.Width)
        OrdemLabel.Text = "Ordem: " & MatrizAtual.Ordem
        ElemLabel.Text = "Elementos: " & MatrizAtual.Linhas * MatrizAtual.Colunas
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            DetLabel.Text = "Determinante: " & MatrizAtual.Determinante
        Else
            DetLabel.Text = "Determinante: Desconhecida"
        End If
    End Sub

    Private Sub TranspostaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TranspostaToolStripMenuItem.Click
        PictureBox1.Image = MatrizAtual.Transposta.Desenhar(PictureBox1.Height, PictureBox1.Width)
        OrdemLabel.Text = "Ordem: " & MatrizAtual.Ordem
        ElemLabel.Text = "Elementos: " & MatrizAtual.Linhas * MatrizAtual.Colunas
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            DetLabel.Text = "Determinante: " & MatrizAtual.Determinante
        Else
            DetLabel.Text = "Determinante: Desconhecida"
        End If
    End Sub

    Private Sub CofatoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CofatoresToolStripMenuItem.Click
        PictureBox1.Image = MatrizAtual.Cofatores.Desenhar(PictureBox1.Height, PictureBox1.Width)
        OrdemLabel.Text = "Ordem: " & MatrizAtual.Ordem
        ElemLabel.Text = "Elementos: " & MatrizAtual.Linhas * MatrizAtual.Colunas
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            DetLabel.Text = "Determinante: " & MatrizAtual.Determinante
        Else
            DetLabel.Text = "Determinante: Desconhecida"
        End If
    End Sub

    Private Sub DeterminanteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeterminanteToolStripMenuItem.Click
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            MsgBox("O determinante é: " & MatrizAtual.Determinante, MsgBoxStyle.Information, "Determinante")
        Else
            MsgBox("O determinante é desconhecida.", MsgBoxStyle.Information, "Determinante")
        End If

    End Sub

    Private Sub AdjuntaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdjuntaToolStripMenuItem.Click
        PictureBox1.Image = MatrizAtual.Adjunta.Desenhar(PictureBox1.Height, PictureBox1.Width)
        OrdemLabel.Text = "Ordem: " & MatrizAtual.Ordem
        ElemLabel.Text = "Elementos: " & MatrizAtual.Linhas * MatrizAtual.Colunas
        If MatrizAtual.Linhas = MatrizAtual.Colunas Then
            DetLabel.Text = "Determinante: " & MatrizAtual.Determinante
        Else
            DetLabel.Text = "Determinante: Desconhecida"
        End If
    End Sub

    Private Sub SalvarMatrizToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalvarMatrizToolStripMenuItem.Click
        With New SaveFileDialog
            .Filter = "Imagens (*.jpg)|*.jpg"
            .Title = "Salvar matriz"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                PictureBox1.Image.Save(.FileName, Imaging.ImageFormat.Png)
            End If
        End With
    End Sub

    Private Sub SairToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SairToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub SobreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SobreToolStripMenuItem.Click
        AboutBox1.ShowDialog()

    End Sub
End Class
