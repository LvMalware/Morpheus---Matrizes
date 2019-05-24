'Criado por Lucas V. Araujo
'Bacharelando em Ciência da Computação pela UESPI - Parnaíba - PI

Imports System.Runtime.Serialization

<Serializable()> Public Class Matriz
    Private mLC(1, 1)
    Private mLinhas, mColunas As Integer

    Public Sub New(ByVal Linhas As Integer, ByVal Colunas As Integer)
        mLinhas = Linhas
        mColunas = Colunas
        Dim M(Linhas, Colunas)
        mLC = M
    End Sub

    Sub DefineItem(ByVal Linha As Integer, ByVal Coluna As Integer, ByVal valor As Object)
        On Error Resume Next
        mLC(Linha, Coluna) = valor
    End Sub

    Function PegaItem(ByVal Linha As Integer, ByVal Coluna As Integer) As Object
        On Error Resume Next
        Return mLC(Linha, Coluna)
    End Function

    Public ReadOnly Property Ordem() As String
        Get
            If mLinhas = mColunas Then
                Return mLinhas
            End If
            Return mLinhas & "x" & mColunas
        End Get
    End Property

    Public ReadOnly Property Linhas() As Integer
        Get
            Return mLinhas
        End Get
    End Property

    Public ReadOnly Property Colunas() As Integer
        Get
            Return mColunas
        End Get
    End Property

    Function Transposta() As Matriz
        Dim t As New Matriz(mColunas, mLinhas)
        For j = 1 To mColunas
            For i = 1 To mLinhas
                t.DefineItem(j, i, mLC(i, j))
            Next
        Next
        Return t
    End Function

    Overrides Function ToString() As String
        Dim conteudo As String = ""
        For i = 1 To mLinhas
            For j = 1 To mColunas
                conteudo &= mLC(i, j)
                If j < mColunas Then
                    conteudo &= "   " 'Chr(9)
                End If
            Next
            conteudo &= vbCrLf
        Next
        Return conteudo
    End Function

    Function Desenhar(ByVal Altura As Integer, ByVal Largura As Integer, Optional ByVal Fonte As Font = Nothing) As Bitmap 'Em fase de testes...
        If Fonte Is Nothing Then
            Fonte = SystemFonts.CaptionFont
        End If
        Dim IMG As New Drawing.Bitmap(Largura, Altura, 0, Imaging.PixelFormat.Format16bppRgb555, 0)
        Dim g As Graphics = Graphics.FromImage(IMG)
        g.FillRectangle(Brushes.White, New Drawing.Rectangle(0, 0, IMG.Width, IMG.Height))
        Dim Sz = g.MeasureString(Me.ToString, Fonte, New Size(Largura, Altura))
        Dim X, Y As Single
        X = (Largura - Sz.Width) / 2
        Y = (Altura - Sz.Height) / 2
        Dim iLinhaY, iLinhaX, fLinhaY, fLinhaX As Single
        iLinhaX = X
        iLinhaY = Y - 2
        fLinhaX = X + Sz.Width + iLinhaX
        fLinhaY = Y + Sz.Height
        g.DrawLine(Pens.Black, iLinhaX, iLinhaY, iLinhaX, fLinhaY)
        g.DrawLine(Pens.Black, iLinhaX + Sz.Width, iLinhaY, iLinhaX + Sz.Width, fLinhaY)
        g.DrawString(Me.ToString, Fonte, Brushes.Black, X, Y)
        Return IMG
    End Function

    Function Somar(ByVal Com As Matriz) As Matriz
        If Com.Linhas = Me.Linhas And Com.Colunas = Me.Colunas Then
            Dim m As New Matriz(Me.Linhas, Me.Colunas)
            For i = 1 To Me.Linhas
                For j = 1 To Me.Colunas
                    Dim mit = Me.PegaItem(i, j)
                    Dim cit = Com.PegaItem(i, j)
                    Dim result = Nothing
                    If IsNumeric(mit) And IsNumeric(cit) Then
                        result = TSingle(mit) + TSingle(cit)
                    Else
                        result = mit & "+" & cit
                        If Not (IsNumeric(mit) And IsNumeric(cit)) And (mit = cit) Then
                            result = "2" & mit
                        End If
                    End If
                    m.DefineItem(i, j, result)
                Next
            Next
            Return m
        Else
            Throw New Exception("Não é possível somar matrizes com tamanhos diferentes!")
        End If
    End Function

    Function Subtrair(ByVal V As Matriz) As Matriz
        If Me.Linhas = V.Linhas And Me.Colunas = V.Colunas Then
            Dim m As New Matriz(Me.Linhas, Me.Colunas)
            For i = 1 To Me.Linhas
                For j = 1 To Me.Colunas
                    Dim result = Nothing
                    Dim mit = Me.PegaItem(i, j)
                    Dim cit = V.PegaItem(i, j)
                    If mit = cit Then
                        result = 0
                    Else
                        If IsNumeric(mit) And IsNumeric(cit) Then
                            result = TSingle(mit) - TSingle(cit)
                        Else
                            result = mit & "-" & cit
                        End If
                    End If
                    m.DefineItem(i, j, result)
                Next
            Next
            Return m
        Else
            Throw New Exception("Não é possível subtrair matrizes com tamanhos diferentes!")
        End If
    End Function

    Sub AddItem(ByVal Item As Object)
        For i = 1 To Me.Linhas
            For j = 1 To Me.Colunas
                If Me.PegaItem(i, j) Is Nothing Then
                    Me.DefineItem(i, j, Item)
                    Exit Sub
                End If
            Next
        Next
    End Sub

    Function PegaLinha(ByVal Linha As Integer) As Object()
        Dim LN(mColunas) As String
        For i = 1 To mColunas
            LN(i) = Me.PegaItem(Linha, i)
        Next
        Return LN
    End Function

    Function PegaColuna(ByVal Coluna As Integer) As Object()
        Dim CL(mLinhas) As String
        For i = 1 To mLinhas
            CL(i) = Me.PegaItem(i, Coluna)
        Next
        Return CL
    End Function

    Public ReadOnly Property Completa() As Boolean
        Get
            For i = 1 To mLinhas
                For j = 1 To mColunas
                    If "" & Me.PegaItem(i, j) = "" Then
                        Return False
                    End If
                Next
            Next
            Return True

        End Get
    End Property

    Function Multiplicar(ByVal Por As Matriz) As Matriz
        If Me.Colunas = Por.Linhas Then
            Dim m As New Matriz(Me.Linhas, Por.Colunas)
            For i = 1 To m.Linhas
                For j = 1 To m.Colunas
                    m.DefineItem(i, j, LinhaXLinha(Me.PegaLinha(i), Por.PegaColuna(j)))
                    If m.Completa Then
                        Return m
                    End If
                Next
            Next
            Return m
        Else
            Throw New Exception("Só é possível multiplicar matrizes se o número de linhas de uma for igual ao número de colunas da outra!")
        End If
    End Function

    Private Function LinhaXLinha(ByVal Linha1() As Object, ByVal Linha2() As Object) As Integer
        On Error Resume Next
        For i = 1 To Linha1.Count - 1
            LinhaXLinha += TSingle(Linha1(i)) * TSingle(Linha2(i))
        Next
    End Function

    Public ReadOnly Property Escalonada() As Matriz
        Get
            Return FullEscalona(EscalonarSuperior(EscalonarInferior(Me)))
        End Get
    End Property

    Function FullEscalona(ByVal Ma As Matriz) As Matriz
        Dim RetMat As Matriz = Ma
        For i = 1 To Ma.Linhas
            Dim Ln() = Ma.PegaLinha(i)
            Dim El = TSingle(Ln(i))
            Dim NovaLinha() = DivideLinha(Ln, El)
            For k = 1 To NovaLinha.Count - 1
                RetMat.DefineItem(i, k, NovaLinha(k))
            Next
        Next
        Return RetMat
    End Function

    Private Function EscalonarSuperior(ByVal M As Matriz, Optional ByVal CDone As Integer = 1) As Matriz
        Dim Matriz_Escalonada As Matriz = M
        For i = 1 To CDone
            Dim NovaLinha() = M.EliminaPor(i, CDone)
            For k = 1 To NovaLinha.Count - 1
                Matriz_Escalonada.DefineItem(i, k, NovaLinha(k))
            Next
        Next
        If FullSuperior(Matriz_Escalonada) Then
            Return Matriz_Escalonada
        End If
        Return EscalonarSuperior(Matriz_Escalonada, CDone + 1)
    End Function

    Private Function FullSuperior(ByVal Ma As Matriz) As Boolean
        For i = 1 To Ma.Linhas
            For j = 1 To Ma.Colunas - 1
                If j > i Then
                    If Ma.PegaItem(i, j) <> "0" Then
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function

    Private Function EscalonarInferior(ByVal M As Matriz, Optional ByVal LDone As Integer = 1, Optional ByVal CDone As Integer = 1) As Matriz
        On Error Resume Next
        Dim Matriz_Escalonada As Matriz = M
        For i = LDone + 1 To M.Linhas
            Dim NovaLinha() = M.EliminaPor(i, CDone)
            For k = 1 To NovaLinha.Count - 1
                Matriz_Escalonada.DefineItem(i, k, NovaLinha(k))
            Next
        Next
        If FullInferior(Matriz_Escalonada) Then
            Return Matriz_Escalonada
        End If
        Return EscalonarInferior(Matriz_Escalonada, LDone + 1, CDone + 1)
    End Function

    Private Function FullInferior(ByVal M As Matriz) As Boolean
        For i = 1 To M.Linhas
            For j = 1 To M.Colunas
                If i > j Then
                    If M.PegaItem(i, j) <> "0" Then
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function

    Public ReadOnly Property Inversa() As Matriz
        Get
            Return Inv(Me.Determinante, Me.Adjunta)
        End Get
    End Property

    Private Function Inv(ByVal det As Single, ByVal Adj As Matriz) As Matriz
        Dim Ret_ As New Matriz(Adj.Linhas, Adj.Colunas)
        For i = 1 To Adj.Linhas
            For j = 1 To Adj.Colunas
                Dim n = Adj.PegaItem(i, j)
                If IsNumeric(n) Then
                    If (TSingle(n) Mod det) = 0 Then
                        Ret_.DefineItem(i, j, TSingle(n) / det)
                    Else
                        Ret_.DefineItem(i, j, Str(n) & "/" & Str(det))
                    End If
                Else
                    Ret_.DefineItem(i, j, Str(n) & "/" & Str(det))
                End If
            Next
        Next
        Return Ret_
    End Function

    Public ReadOnly Property Adjunta() As Matriz
        Get
            Return Me.Cofatores.Transposta
        End Get
    End Property

    Public ReadOnly Property Cofatores() As Matriz
        Get
            Return Cofats(Me)
        End Get
    End Property

    Private Function Cofats(ByVal M As Matriz) As Matriz
        Dim Coft As New Matriz(M.Linhas, M.Colunas)
        For i = 1 To M.Linhas
            For j = 1 To M.Colunas
                Dim cof As Single = Math.Pow(-1, i + j) * Det(M.Elimina(i, j))
                Coft.AddItem(cof)
            Next
        Next
        Return Coft
    End Function

    Public ReadOnly Property Determinante() As Single
        Get
            If Me.Linhas = Me.Colunas Then
                Return Det(Me)
            End If
            Return Nothing
        End Get
    End Property

    Private Function Det(ByVal M As Matriz) As Single
        If M.Ordem = "1" Then
            Return TSingle(M.PegaItem(1, 1))
        ElseIf M.Ordem = "2" Then
            Return (TSingle(M.PegaItem(1, 1)) * TSingle(M.PegaItem(2, 2))) - (TSingle(M.PegaItem(1, 2)) * TSingle(M.PegaItem(2, 1)))
        Else
            For i = 1 To M.Colunas
                Det += TSingle(M.PegaItem(1, i)) * Math.Pow(-1, 1 + i) * Det(M.Elimina(1, i))
            Next
        End If
    End Function

    Private Function TSingle(ByVal O As Object) As Single
        On Error Resume Next
        If O.ToString.Contains("/") Then
            Dim k = O.ToString.Split("/")
            Return Single.Parse(k(0)) / Single.Parse(k(1))
        End If
        Return Single.Parse(O)
    End Function

    Private Function Elimina(ByVal L As Integer, ByVal C As Integer) As Matriz
        Dim Ret As New Matriz(Me.Linhas - 1, Me.Colunas - 1)
        For i = 1 To Me.Linhas
            For j = 1 To Me.Colunas
                If i = L Or j = C Then

                Else
                    Ret.AddItem(Me.PegaItem(i, j))
                End If
            Next
        Next
        Return Ret
    End Function

    Private Function EliminaPor(ByVal iL As Integer, ByVal iC As Integer) As Object()
        Dim It As Single = TSingle(Me.PegaItem(iL, iC))
        If iL > iC Then
            For i = iL - 1 To 1 Step -1
                Dim cIT As Single = TSingle(Me.PegaItem(i, iC))
                If (It = cIT) Then
                    'MsgBox("L" & iL & " -> L" & i & " - L" & iL)
                    Return Me.SubtraiLinhas(Me.PegaLinha(i), Me.PegaLinha(iL))
                ElseIf (It = (-1 * cIT)) Or (cIT = (-1 * It)) Then
                    'MsgBox("L" & iL & " -> L" & iL & " + L" & i)
                    Return Me.SomaLinhas(Me.PegaLinha(i), Me.PegaLinha(iL))
                ElseIf cIT > 0 And It > 0 Then
                    'MsgBox("L" & iL & " -> " & It & "L" & i & "-" & cIT & "L" & iL)
                    Return Me.SubtraiLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                ElseIf cIT < 0 And It > 0 Then
                    'MsgBox("L" & iL & " -> L" & iL & "+" & cIT & "L" & i)
                    Return Me.SomaLinhas(MultiplicaLinha(Me.PegaLinha(iL), It), MultiplicaLinha(Me.PegaLinha(i), cIT))
                ElseIf It < 0 And cIT > 0 Then
                    'MsgBox("L" & iL & " -> L" & iL & "+" & cIT & "L" & i)
                    Return Me.SomaLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                ElseIf It < 0 And cIT < 0 Then
                    'MsgBox("L" & iL & " -> " & It & "L" & i & "-" & cIT & "L" & iL)
                    Return Me.SubtraiLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                End If
            Next
        Else
            For i = Me.Linhas To iL + 1 Step -1
                Dim cIT As Single = TSingle(Me.PegaItem(i, iC))
                If (It = cIT) Then
                    'MsgBox("L" & iL & " -> L" & i & " - L" & iL)
                    Return Me.SubtraiLinhas(Me.PegaLinha(i), Me.PegaLinha(iL))
                ElseIf (It = (-1 * cIT)) Or (cIT = (-1 * It)) Then
                    'MsgBox("L" & iL & " -> L" & iL & " + L" & i)
                    Return Me.SomaLinhas(Me.PegaLinha(i), Me.PegaLinha(iL))
                ElseIf cIT > 0 And It > 0 Then
                    'MsgBox("L" & iL & " -> " & cIT & "L" & iL & "-" & It & "L" & i)
                    Return Me.SubtraiLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                ElseIf cIT < 0 And It > 0 Then
                    'MsgBox("L" & iL & " -> " & It & "L" & i & "+" & cIT & "L" & iL)
                    Return Me.SomaLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), -1 * cIT))
                ElseIf It < 0 And cIT > 0 Then
                    'MsgBox("L" & iL & " -> " & It & "L" & i & "-" & cIT & "L" & iL)
                    Return Me.SubtraiLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                ElseIf It < 0 And cIT < 0 Then
                    'MsgBox("L" & iL & " -> " & It & "L" & i & "-" & cIT & "L" & iL)
                    Return Me.SubtraiLinhas(MultiplicaLinha(Me.PegaLinha(i), It), MultiplicaLinha(Me.PegaLinha(iL), cIT))
                End If
            Next
        End If
        Return Me.PegaLinha(iL)
    End Function

    Function SomaLinhas(ByVal Linha1() As Object, ByVal Linha2() As Object) As Object()
        On Error Resume Next
        Dim NovaLinha(Linha1.Count - 1) As String
        For i = 1 To Linha1.Count - 1
            NovaLinha(i) = TSingle(Linha1(i)) + TSingle(Linha2(i))
        Next
        Return NovaLinha
    End Function

    Function SubtraiLinhas(ByVal Linha1() As Object, ByVal Linha2() As Object) As Object()
        On Error Resume Next
        Dim NovaLinha(Linha1.Count - 1) As String
        For i = 1 To Linha1.Count - 1
            NovaLinha(i) = TSingle(Linha1(i)) - TSingle(Linha2(i))
        Next
        Return NovaLinha
    End Function

    Function DivideLinha(ByVal _Linha() As Object, ByVal N As Single) As Object()
        Dim Linha_Retorno(_Linha.Count - 1) As String
        For i = 1 To _Linha.Count - 1
            Linha_Retorno(i) = TSingle(_Linha(i)) / N
        Next
        Return Linha_Retorno
    End Function

    Function MultiplicaLinha(ByVal _Linha() As Object, ByVal N As Single) As Object()
        Dim Linha_Retorno(_Linha.Count - 1) As String
        For i = 1 To _Linha.Count - 1
            Linha_Retorno(i) = TSingle(_Linha(i)) * N
        Next
        Return Linha_Retorno
    End Function

    Function MultiplicacaoEscalar(ByVal Numero As Single) As Matriz
        Dim m As New Matriz(Me.Linhas, Me.Colunas)
        For i = 1 To Me.Linhas
            For j = 1 To Me.Colunas
                Dim result = Nothing
                Dim mit = Me.PegaItem(i, j)
                If IsNumeric(mit) Then
                    result = TSingle(mit) * Numero
                Else
                    result = Numero & mit
                End If
                m.DefineItem(i, j, result)
            Next
        Next
        Return m
    End Function
End Class
