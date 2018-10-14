Public Class Paleta

    Private Sub atualizar()
        Panel1.BackColor = Configuracoes.corFundo
        Panel2.BackColor = Configuracoes.corPadrao
        Panel3.BackColor = Configuracoes.corNumeros
        Panel4.BackColor = Configuracoes.corTags
        Panel5.BackColor = Configuracoes.corAspas
        Panel6.BackColor = Configuracoes.corComentario

        Form1.atualizarCores()
        ColorDialog1.Color = Color.Black
    End Sub

    Private Sub Paleta_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        atualizar()
    End Sub

    Private Sub Panel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corFundo = ColorDialog1.Color
        atualizar()
    End Sub

    Private Sub Panel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel2.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corPadrao = ColorDialog1.Color
        atualizar()
    End Sub

    Private Sub Panel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel3.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corNumeros = ColorDialog1.Color
        atualizar()
    End Sub

    Private Sub Panel4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel4.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corTags = ColorDialog1.Color
        atualizar()
    End Sub

    Private Sub Panel5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel5.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corAspas = ColorDialog1.Color
        atualizar()
    End Sub

    Private Sub Panel6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel6.Click
        ColorDialog1.ShowDialog()
        Configuracoes.corComentario = ColorDialog1.Color
        atualizar()
    End Sub
End Class