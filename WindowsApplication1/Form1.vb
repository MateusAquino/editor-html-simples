Imports System.Text.RegularExpressions
Imports System.IO

Public Class Form1

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        painel.VerticalScroll.Visible = False
        painel.HorizontalScroll.Visible = False
        codigo.AcceptsTab = True
        codigo.Focus()
        abrir(Nothing)
        atualizarCores()
        undo.Visible = False
        redo.Visible = False
        urc.AddItem(codigo.Text)
        webPanel.VerticalScroll.Visible = False

        'Textos de ajuda para ícones
        Dim tt As New ToolTip()
        tt.SetToolTip(undo, "Desfazer")
        tt.SetToolTip(redo, "Refazer")
        tt.SetToolTip(fecha, "Fechar guia")
        tt.SetToolTip(navPadrao, "Abrir no Navegador Padrão")
        tt.SetToolTip(web, "Pré-visualizar")
    End Sub

    ' ----- Expressões Regulares ----- '
    Private expTags As New Regex("(</?\w+((\s+\w+(\s*=\s*(?:\x22.*?\x22|'.*?'|[\^'\x22>\s]+))?)+\s*|\s*)/?>)")
    Private expAspas As New Regex("(\x22[^\x22]+\x22)")
    Private expComentario As New Regex("(\<![ \r\n\t]*(--([^\-]|[\r\n]|-[^\-])*--[ \r\n\t]*)\>)")

    ' ----- BOTÕES DE ICONES ----- '
    'Un-do & Re-do
    Private urc As New UndoRedoClass(Of String)()
    Private NoAdd As Boolean = False

    Private Sub undo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles undo.Click
        NoAdd = True
        urc.Undo()
        codigo.Text = urc.CurrentItem
        undo.Visible = urc.CanUndo
        redo.Visible = urc.CanRedo
        NoAdd = False
    End Sub

    Private Sub redo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles redo.Click
        Dim sel = codigo.SelectionStart
        NoAdd = True
        urc.Redo()
        codigo.Text = urc.CurrentItem
        undo.Visible = urc.CanUndo
        redo.Visible = urc.CanRedo
        NoAdd = False
        codigo.SelectionStart = sel
        codigo.SelectionLength = 0
    End Sub

    Private Sub fecha_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fecha.Click
        fechar()
        TabControl1_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub navPadrao_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles navPadrao.Click
        salvar(False)
        If Not TabControl1.SelectedTab.Text.Contains("*") Then
            Process.Start(arquivos.Items.Item(TabControl1.SelectedIndex).ToString())
        End If
    End Sub

    Private Sub web_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles web.Click
        If webPanel.Visible Then
            WebBrowser1.Stop()
            webPanel.Visible = False
        Else
            salvar(False)
            If Not TabControl1.SelectedTab.Text.Contains("*") Then
                WebBrowser1.Navigate(arquivos.Items.Item(TabControl1.SelectedIndex).ToString())
                webPanel.Visible = True
            End If
        End If
    End Sub

    Private Sub webUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles webUndo.Click
        WebBrowser1.GoBack()
    End Sub

    Private Sub webRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles webRedo.Click
        WebBrowser1.GoForward()
    End Sub

    Private Sub webAtualiza_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles webAtualiza.Click
        WebBrowser1.Refresh()
    End Sub

    ' ----- EVENTOS ----- '
    Private Sub RichTextBox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles codigo.TextChanged
        If Not NoAdd And urc.CurrentItem <> codigo.Text Then
            urc.AddItem(codigo.Text)
            undo.Visible = urc.CanUndo
            redo.Visible = urc.CanRedo
        End If

        'quantidade de linhas
        Dim str As String = ""
        For x As Integer = 1 To codigo.Text.Split(vbLf).Length
            str &= x & vbCrLf
        Next
        linhas.Text = str

        Dim OldSel As Integer = codigo.SelectionStart
        codigo.SelectAll()
        codigo.SelectionColor = Configuracoes.corPadrao
        codigo.DeselectAll()
        Dim MC1 As MatchCollection = expTags.Matches(codigo.Text)
        Dim MC2 As MatchCollection = expAspas.Matches(codigo.Text)
        Dim MC3 As MatchCollection = expComentario.Matches(codigo.Text)
        'Tags
        For Each M As Match In MC1
            codigo.SelectionStart = M.Index
            codigo.SelectionLength = M.Value.Length
            codigo.SelectionColor = Configuracoes.corTags
            codigo.DeselectAll()
        Next
        'Aspas
        For Each M As Match In MC2
            codigo.SelectionStart = M.Index
            codigo.SelectionLength = M.Value.Length
            codigo.SelectionColor = Configuracoes.corAspas
            codigo.DeselectAll()
        Next
        'Comentários
        For Each M As Match In MC3
            codigo.SelectionStart = M.Index
            codigo.SelectionLength = M.Value.Length
            codigo.SelectionColor = Configuracoes.corComentario
            codigo.DeselectAll()
        Next
        codigos.Items.RemoveAt(TabControl1.SelectedIndex)
        codigos.Items.Insert(TabControl1.SelectedIndex, codigo.Text)
        codigo.SelectionStart = OldSel
        If Not TabControl1.SelectedTab.Text.Contains("*") And Not travarAsterisco Then
            TabControl1.SelectedTab.Text += " *"
        End If
    End Sub

    Private Sub codigo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles codigo.KeyDown
        If e.KeyCode = Keys.Z And e.Control And undo.Visible Then
            undo_Click(Nothing, Nothing)
        ElseIf e.KeyCode = Keys.Y And e.Control And redo.Visible Then
            redo_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub codigo_VScroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles codigo.VScroll
        Dim x As ScrollEventArgs = e
        painel.VerticalScroll.Value = painel.VerticalScroll.Value + x.NewValue - x.OldValue
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If webPanel.Visible Then
            web_Click(Nothing, Nothing)
        End If
        travarAsterisco = True
        codigo.Text = codigos.Items.Item(TabControl1.SelectedIndex)
        travarAsterisco = False
    End Sub

    Private Sub WebBrowser1_Navigated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
        webURL.Text = WebBrowser1.Url.ToString()
    End Sub

    Private Sub webURL_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles webURL.KeyDown
        If e.KeyCode = Keys.Enter Then
            WebBrowser1.Navigate(webURL.Text)
        End If
    End Sub

    ' ----- FUNÇÕES ----- '
    Private Sub inserir(ByVal antes As String, ByVal depois As String)
        Dim x = codigo.SelectionStart
        Dim l = codigo.SelectionLength
        NoAdd = True
        codigo.Text = codigo.Text.Insert(x + l, depois)
        NoAdd = False
        codigo.Text = codigo.Text.Insert(x, antes)
        codigo.SelectionStart = x + antes.Length + l
    End Sub

    Private Sub substituir(ByVal novo As String)
        codigo.SelectedText = novo
    End Sub
    
    Private Function nome(ByVal arquivo As String)
        Return arquivo.Substring(arquivo.LastIndexOf("\") + 1, arquivo.Length - arquivo.LastIndexOf("\") - 1)
    End Function

    Private Sub abrir(ByVal arquivo As String)
        If arquivo Is Nothing Then
            arquivos.Items.Add("")
            TabControl1.TabPages.Add(New TabPage("Novo Arquivo *"))
            codigos.Items.Add("<html>" + vbCrLf + vbCrLf +
                              "<!-- Cabeça -->" + vbCrLf +
                              "<head>" + vbCrLf + "</head>" + vbCrLf + vbCrLf +
                              "<!-- Corpo -->" + vbCrLf +
                              "<body>" + vbCrLf + "</body>" + vbCrLf + vbCrLf +
                              "</html>")
        Else
            arquivos.Items.Add(arquivo)
            TabControl1.TabPages.Add(New TabPage(nome(arquivo)))
            codigos.Items.Add(File.ReadAllText(arquivo))
        End If
        TabControl1.SelectTab(TabControl1.TabPages.Count - 1)
    End Sub

    Private Sub fechar()
        If TabControl1.TabCount > 1 Then
            arquivos.Items.RemoveAt(TabControl1.SelectedIndex)
            codigos.Items.RemoveAt(TabControl1.SelectedIndex)
            TabControl1.TabPages.RemoveAt(TabControl1.SelectedIndex)
        End If
    End Sub

    Private travarAsterisco As Boolean = False
    Private Sub salvar(ByVal como As Boolean)
        If arquivos.Items.Item(TabControl1.SelectedIndex) = "" Or como Then
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim arquivo As String = SaveFileDialog1.FileName
                File.WriteAllText(arquivo, codigo.Text)
                TabControl1.SelectedTab.Text = nome(arquivo)
                arquivos.Items.Item(TabControl1.SelectedIndex) = arquivo
            End If
        ElseIf TabControl1.SelectedTab.Text.Contains("*") Then
            File.WriteAllText(arquivos.Items.Item(TabControl1.SelectedIndex), codigo.Text)
            TabControl1.SelectedTab.Text = nome(arquivos.Items.Item(TabControl1.SelectedIndex))
        End If
    End Sub

    Public Sub atualizarCores()
        RichTextBox1_TextChanged(Nothing, Nothing)
        painel.BackColor = Configuracoes.corFundo
        codigo.BackColor = Configuracoes.corFundo
        linhas.ForeColor = Configuracoes.corNumeros
    End Sub

    ' ----- MENU ----- '
    Private Sub NovoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NovoToolStripMenuItem.Click
        abrir(Nothing)
    End Sub

    Private Sub SalvarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalvarToolStripMenuItem.Click
        salvar(False)
    End Sub

    Private Sub SalvarComoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalvarComoToolStripMenuItem.Click
        salvar(True)
    End Sub

    Private Sub AbrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AbrirToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK And File.Exists(OpenFileDialog1.FileName) Then
            Dim idx = TabControl1.SelectedIndex
            abrir(OpenFileDialog1.FileName)
            If arquivos.Items.Item(idx) = Nothing Then
                arquivos.Items.RemoveAt(idx)
                codigos.Items.RemoveAt(idx)
                TabControl1.TabPages.RemoveAt(idx)
            End If
        End If
    End Sub

    Private Sub FecharToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FecharToolStripMenuItem.Click
        fechar()
        TabControl1_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub CorDeFundoToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CorDeFundoToolStripMenuItem1.Click
        Paleta.Show()
    End Sub

    Private Sub ParágrafoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ParágrafoToolStripMenuItem.Click
        inserir("<p>", "</p>")
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        inserir("<br/>", "")
    End Sub

    Private Sub InserirToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InserirToolStripMenuItem1.Click
        Dim txt = InputBox("Endereço:", "Digite o endereço do link")
        If Not txt = "" Then
            inserir("<a href=""" + txt + """>", "</a>")
        End If
    End Sub

    Private Sub FigurasEImagensToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FigurasEImagensToolStripMenuItem.Click
        Dim txt = InputBox("Endereço da imagem:", "Digite a fonte (source) da imagem")
        If Not txt = "" Then
            inserir("<img src=""" + txt + """ alt=""", """>")
        End If
    End Sub

    Private Sub OrdenadosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrdenadosToolStripMenuItem.Click
        Dim str = codigo.SelectedText
        Dim linhas() = str.Split(vbLf)
        Dim novo As String = "<ol type=""1"">" + vbCrLf
        If linhas.Length = 0 Then
            novo += "<li></li>" + vbCrLf
        End If
        For x = 0 To linhas.Length - 1
            While linhas(x).StartsWith(vbTab) Or linhas(x).StartsWith(" ")
                novo += linhas(x)(0)
                linhas(x) = linhas(x).Substring(1)
            End While
            novo += "<li>" + linhas(x).Trim() + "</li>" + vbCrLf
        Next
        substituir(novo + "</ol>")
    End Sub

    Private Sub NãoOrdenadosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NãoOrdenadosToolStripMenuItem.Click
        Dim str = codigo.SelectedText
        Dim linhas() = str.Split(vbLf)
        Dim novo As String = "<ul>" + vbCrLf
        If linhas.Length = 0 Then
            novo += "<li></li>" + vbCrLf
        End If
        For x = 0 To linhas.Length - 1
            While linhas(x).StartsWith(vbTab) Or linhas(x).StartsWith(" ")
                novo += linhas(x)(0)
                linhas(x) = linhas(x).Substring(1)
            End While
            novo += "<li>" + linhas(x).Trim() + "</li>" + vbCrLf
        Next
        substituir(novo + "</ul>")
    End Sub

    Private Sub CentralizarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CentralizarToolStripMenuItem.Click
        inserir("<span style=""text-align: center;"">", "</span>")
    End Sub

    Private Sub NegritoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NegritoToolStripMenuItem.Click
        inserir("<b>", "</b>")
    End Sub

    Private Sub SublinhadoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SublinhadoToolStripMenuItem.Click
        inserir("<u>", "</u>")
    End Sub

    Private Sub ItálicoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItálicoToolStripMenuItem.Click
        inserir("<i>", "</i>")
    End Sub

    Private Sub FonteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FonteToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            Dim font = "font:"
            If FontDialog1.Font.Italic Then font += " italic"
            If FontDialog1.Font.Bold Then font += " bold"
            font += " " + Math.Round(FontDialog1.Font.Size).ToString() + "px"
            font += " '" + FontDialog1.Font.FontFamily.Name + "'"
            font += ";"
            If FontDialog1.Font.Underline Or FontDialog1.Font.Strikeout Then
                font += " text-decoration:"
                If FontDialog1.Font.Underline Then font += " underline"
                If FontDialog1.Font.Strikeout Then font += " line-through"
                font += ";"
            End If
            Dim color = FontDialog1.Color
            font += " color: " + String.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B) + ";"
            inserir("<span style=""" + font + """>", "</span>")
        End If
    End Sub

    Private Sub CorDeFundoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorDeFundoToolStripMenuItem.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            Dim color = ColorDialog1.Color
            inserir("<span style=""background-color: " + String.Format("#{0:X2}{1:X2}{2:X2}", Color.R, Color.G, Color.B) + ";"">", "</span>")
        End If
    End Sub
End Class
