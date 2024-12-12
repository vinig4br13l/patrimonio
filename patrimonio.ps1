# Importando o assembly do Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#=======================================================================================


while($true){

# Criando um formulário
$form = New-Object System.Windows.Forms.Form
$form.Text = "Patrimônios"
$form.Size = New-Object System.Drawing.Size(1100, 1000)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
#$form.BackColor = '#ffffff'

#===========================================================================
$fonteGroupBox = New-Object System.Drawing.Font("Arial", 12)

$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Text = "Cadastro de Patrimônio"
$groupBox1.Font = $fonteGroupBox
#$groupBox1.ForeColor = [System.Drawing.Color]::White
$groupBox1.Size = New-Object System.Drawing.Size(800, 625)
$groupBox1.Location = New-Object System.Drawing.Point(200, 100)
#$groupBox1.BackColor = '#000000'

#===========================================================================


# Criar o objeto PictureBox para exibir a imagem
$logo = New-Object System.Windows.Forms.PictureBox
$logo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage

# Carregar a imagem (substitua o caminho pela sua imagem)
$caminhoLogo = "C:\Users\gabriel.vinicius\Desktop\censipam-logo.png"
$logo.Image = [System.Drawing.Image]::FromFile($caminhoLogo)

# Definir a posição da logo na parte superior (y = 0)
$logo.Location = New-Object System.Drawing.Point(425, 0)
$logo.Size = New-Object System.Drawing.Size(250,100)
$form.Controls.Add($logo)



#===========================================================================
$fonteTexto = New-Object System.Drawing.Font("Arial", 16)

$labelPatrimonio = New-Object System.Windows.Forms.Label
$labelPatrimonio.Text = "Patrimônio:"
$labelPatrimonio.AutoSize = $true
$labelPatrimonio.Font = $fonteTexto
#$labelPatrimonio.ForeColor = [System.Drawing.Color]::White
$labelPatrimonio.Location = New-Object System.Drawing.Point(60, 50)
$labelPatrimonio.Size = New-Object System.Drawing.Size(62,20)
$form.Controls.Add($labelPatrimonio)

<<<<<<< HEAD
$obrigatorioPatri = New-Object System.Windows.Forms.Label
$obrigatorioPatri.Text = "*"
$obrigatorioPatri.AutoSize = $true
$obrigatorioPatri.Font = $fonteTexto
$obrigatorioPatri.ForeColor = [System.Drawing.Color]::Red
$obrigatorioPatri.Location = New-Object System.Drawing.Point(175, 50)
$obrigatorioPatri.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioPatri)


$txtPatrimonio = New-Object System.Windows.Forms.TextBox
$txtPatrimonio.Font = $fonteTexto
$txtPatrimonio.MaxLength = 7
$txtPatrimonio.Location = New-Object System.Drawing.Point(200, 50)
=======
$txtPatrimonio = New-Object System.Windows.Forms.TextBox
$txtPatrimonio.Font = $fonteTexto
$txtPatrimonio.MaxLength = 7
$txtPatrimonio.Location = New-Object System.Drawing.Point(190, 50)
>>>>>>> 7e3ab440133d3ff345759ca1e15ef67a7284eaa5
$txtPatrimonio.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtPatrimonio)
# Evento KeyDown no primeiro TextBox para mover o foco após a leitura
$txtPatrimonio.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $marca.Focus()
    }
})

#==================================================================================

$labelMarca = New-Object System.Windows.Forms.Label
$labelMarca.Text = "Marca:"
$labelMarca.AutoSize = $true
$labelMarca.Font = $fonteTexto
#$labelMarca.ForeColor = [System.Drawing.Color]::White
$labelMarca.Location = New-Object System.Drawing.Point(60, 100)
$labelMarca.Size = New-Object System.Drawing.Size(60,20)
$form.Controls.Add($labelMarca)

$obrigatorioMarca = New-Object System.Windows.Forms.Label
$obrigatorioMarca.Text = "*"
$obrigatorioMarca.AutoSize = $true
$obrigatorioMarca.Font = $fonteTexto
$obrigatorioMarca.ForeColor = [System.Drawing.Color]::Red
$obrigatorioMarca.Location = New-Object System.Drawing.Point(130, 100)
$obrigatorioMarca.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioMarca)
 7e3ab440133d3ff345759ca1e15ef67a7284eaa5
$marca = New-Object Windows.Forms.ComboBox
$marca.Location = New-Object Drawing.Point(190, 100)
$marca.Size = New-Object Drawing.Size(550, 300)
$marca.Font = $fonteTexto

# Adicionar itens ao ComboBox
$marca.Items.Add("Dell")
$marca.Items.Add("HP")
$marca.Items.Add("Lenovo")
$marca.Items.Add("Samsung")

# Definir uma opção selecionada por padrão (opcional)
$marca.SelectedIndex = 0

$form.Controls.Add($marca)

#==================================================================================

$labelModelo = New-Object System.Windows.Forms.Label
$labelModelo.Text = "Modelo:"
$labelModelo.AutoSize = $true
$labelModelo.Font = $fonteTexto
#$labelModelo.ForeColor = [System.Drawing.Color]::White
$labelModelo.Location = New-Object System.Drawing.Point(60, 150)
$labelModelo.Size = New-Object System.Drawing.Size(70,30)
$form.Controls.Add($labelModelo)

$obrigatorioModelo = New-Object System.Windows.Forms.Label
$obrigatorioModelo.Text = "*"
$obrigatorioModelo.AutoSize = $true
$obrigatorioModelo.Font = $fonteTexto
$obrigatorioModelo.ForeColor = [System.Drawing.Color]::Red
$obrigatorioModelo.Location = New-Object System.Drawing.Point(142, 150)
$obrigatorioModelo.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioModelo)


$modelo = New-Object Windows.Forms.ComboBox
$modelo.Location = New-Object Drawing.Point(190, 150)
$modelo.Size = New-Object Drawing.Size(550, 300)
$modelo.Font = $fonteTexto

$modelo.Items.Add("OptiPlex 7050")
$modelo.Items.Add("Modelo B")
$modelo.Items.Add("Modelo C")
$modelo.Items.Add("Modelo D")

$modelo.SelectedIndex = 0
$form.Controls.Add($modelo)

#==================================================================================

$labelNumeroSerie = New-Object System.Windows.Forms.Label
$labelNumeroSerie.Text = "Nº de Série:"
$labelNumeroSerie.AutoSize = $true
$labelNumeroSerie.Font = $fonteTexto
#$labelNumeroSerie.ForeColor = [System.Drawing.Color]::White
$labelNumeroSerie.Location = New-Object System.Drawing.Point(60, 200)
$labelNumeroSerie.Size = New-Object System.Drawing.Size(80,40)
$form.Controls.Add($labelNumeroSerie)

$txtNumeroSerie = New-Object System.Windows.Forms.TextBox
$txtNumeroSerie.Font = $fonteTexto
$txtNumeroSerie.Location = New-Object System.Drawing.Point(190, 200)
$txtNumeroSerie.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtNumeroSerie)

$txtNumeroSerie.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $situacao.Focus()
        $txtSituacao.Focus()
    }
})


#==================================================================================

$labelSituacao = New-Object System.Windows.Forms.Label
$labelSituacao.Text = "Situação:"
$labelSituacao.AutoSize = $true
$labelSituacao.Font = $fonteTexto
#$labelSituacao.ForeColor = [System.Drawing.Color]::White
$labelSituacao.Location = New-Object System.Drawing.Point(60, 250)
$labelSituacao.Size = New-Object System.Drawing.Size(90,50)
$form.Controls.Add($labelSituacao)

$obrigatorioSituacao = New-Object System.Windows.Forms.Label
$obrigatorioSituacao.Text = "*"
$obrigatorioSituacao.AutoSize = $true
$obrigatorioSituacao.Font = $fonteTexto
$obrigatorioSituacao.ForeColor = [System.Drawing.Color]::Red
$obrigatorioSituacao.Location = New-Object System.Drawing.Point(155, 250)
$obrigatorioSituacao.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioSituacao)

$situacao = New-Object System.Windows.Forms.ComboBox
$situacao.Location = New-Object System.Drawing.Point(190, 250)
$situacao.Size = New-Object System.Drawing.Size(550,300)
$situacao.Font = $fonteTexto

$situacao.Items.Add("Em uso")
$situacao.Items.Add("Disponível")
$situacao.Items.Add("Com defeito")
$situacao.Items.Add("Com defeito/Quebrado")
$situacao.Items.Add("Não localizado")

$situacao.SelectedIndex = 0
$form.Controls.Add($situacao)

#==================================================================================

$labelDescricao = New-Object System.Windows.Forms.Label
$labelDescricao.Text = "Descrição:"
$labelDescricao.AutoSize = $true
$labelDescricao.Font = $fonteTexto
#$labelDescricao.ForeColor = [System.Drawing.Color]::White
$labelDescricao.Location = New-Object System.Drawing.Point(60, 300)
$labelDescricao.Size = New-Object System.Drawing.Size(100,60)
$form.Controls.Add($labelDescricao)

$obrigatorioDescricao = New-Object System.Windows.Forms.Label
$obrigatorioDescricao.Text = "*"
$obrigatorioDescricao.AutoSize = $true
$obrigatorioDescricao.Font = $fonteTexto
$obrigatorioDescricao.ForeColor = [System.Drawing.Color]::Red
$obrigatorioDescricao.Location = New-Object System.Drawing.Point(170, 300)
$obrigatorioDescricao.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioDescricao)

$txtDescricao = New-Object System.Windows.Forms.TextBox
$txtDescricao.Font = $fonteTexto
$txtDescricao.Location = New-Object System.Drawing.Point(190, 300)
$txtDescricao.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtDescricao)

#==================================================================================

$labelUsuario = New-Object System.Windows.Forms.Label
$labelUsuario.Text = "Usuário:"
$labelUsuario.AutoSize = $true
$labelUsuario.Font = $fonteTexto
#$labelUsuario.ForeColor = [System.Drawing.Color]::White
$labelUsuario.Location = New-Object System.Drawing.Point(60, 350)
$labelUsuario.Size = New-Object System.Drawing.Size(110,70)
$form.Controls.Add($labelUsuario)
$obrigatorioUsuario = New-Object System.Windows.Forms.Label
$obrigatorioUsuario.Text = "*"
$obrigatorioUsuario.AutoSize = $true
$obrigatorioUsuario.Font = $fonteTexto
$obrigatorioUsuario.ForeColor = [System.Drawing.Color]::Red
$obrigatorioUsuario.Location = New-Object System.Drawing.Point(145, 350)
$obrigatorioUsuario.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioUsuario)
$txtUsuario = New-Object System.Windows.Forms.TextBox
$txtUsuario.Font = $fonteTexto
$txtUsuario.Location = New-Object System.Drawing.Point(190, 350)
$txtUsuario.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtUsuario)

#==================================================================================

$labelSetor = New-Object System.Windows.Forms.Label
$labelSetor.Text = "Setor:"
$labelSetor.AutoSize = $true
$labelSetor.Font = $fonteTexto
#$labelSetor.ForeColor = [System.Drawing.Color]::White
$labelSetor.Location = New-Object System.Drawing.Point(60,400)
$labelSetor.Size = New-Object System.Drawing.Size(120,80)
$form.Controls.Add($labelSetor)

$obrigatorioSetor = New-Object System.Windows.Forms.Label
$obrigatorioSetor.Text = "*"
$obrigatorioSetor.AutoSize = $true
$obrigatorioSetor.Font = $fonteTexto
$obrigatorioSetor.ForeColor = [System.Drawing.Color]::Red
$obrigatorioSetor.Location = New-Object System.Drawing.Point(125, 400)
$obrigatorioSetor.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioSetor)

$txtSetor = New-Object System.Windows.Forms.TextBox
$txtSetor.Font = $fonteTexto
$txtSetor.Location = New-Object System.Drawing.Point(190, 400)
$txtSetor.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtSetor)

#==================================================================================

$labelSala = New-Object System.Windows.Forms.Label
$labelSala.Text = "Sala:"
$labelSala.AutoSize = $true
$labelSala.Font = $fonteTexto
#$labelSala.ForeColor = [System.Drawing.Color]::White
$labelSala.Location = New-Object System.Drawing.Point(60,450)
$labelSala.Size = New-Object System.Drawing.Size(130,90)
$form.Controls.Add($labelSala)

$obrigatorioSala = New-Object System.Windows.Forms.Label
$obrigatorioSala.Text = "*"
$obrigatorioSala.AutoSize = $true
$obrigatorioSala.Font = $fonteTexto
$obrigatorioSala.ForeColor = [System.Drawing.Color]::Red
$obrigatorioSala.Location = New-Object System.Drawing.Point(115, 450)
$obrigatorioSala.Size = New-Object System.Drawing.Size(40,20)
$form.Controls.Add($obrigatorioSala)


$txtSala = New-Object System.Windows.Forms.TextBox
$txtSala.Font = $fonteTexto
$txtSala.Location = New-Object System.Drawing.Point(190, 450)
$txtSala.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtSala)

#==================================================================================

$labelRamal = New-Object System.Windows.Forms.Label
$labelRamal.Text = "Ramal:"
$labelRamal.AutoSize = $true
$labelRamal.Font = $fonteTexto
#$labelRamal.ForeColor = [System.Drawing.Color]::White
$labelRamal.Location = New-Object System.Drawing.Point(60,500)
$labelRamal.Size = New-Object System.Drawing.Size(140,100)
$form.Controls.Add($labelRamal)

$txtRamal = New-Object System.Windows.Forms.TextBox
$txtRamal.Font = $fonteTexto
$txtRamal.Location = New-Object System.Drawing.Point(190, 500)
$txtRamal.Size = New-Object System.Drawing.Size(550,300)
$form.Controls.Add($txtRamal)

#==================================================================================

# Botão para salvar
$btnSalvar = New-Object System.Windows.Forms.Button
$btnSalvar.Text = "Salvar"
$btnSalvar.Text = "Salvar em CSV"
$btnSalvar.AutoSize = $true
$btnSalvar.Font = $fonteTexto
#$btnSalvar.BackColor = [System.Drawing.Color]::White
#$btnSalvar.ForeColor = [System.Drawing.Color]::Black
$btnSalvar.Location = New-Object System.Drawing.Point(300, 550)
$btnSalvar.Size = New-Object System.Drawing.Size(200,40)
$btnSalvar.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnSalvar)

# Botão para salvar imagem
$btnSalvarImagem = New-Object System.Windows.Forms.Button
$btnSalvarImagem.Text = "Enviar Imagem"
$btnSalvarImagem.AutoSize = $true
$btnSalvarImagem.Font = $fonteTexto
#$btnSalvar.BackColor = [System.Drawing.Color]::White
#$btnSalvar.ForeColor = [System.Drawing.Color]::Black
$btnSalvarImagem.Location = New-Object System.Drawing.Point(525, 550)
$btnSalvarImagem.Size = New-Object System.Drawing.Size(200,40)
$btnSalvarImagem.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnSalvarImagem)

#==================================================================================

$groupBox1.Controls.Add($labelPatrimonio)
$groupBox1.Controls.Add($obrigatorioPatri)
$groupBox1.Controls.Add($txtPatrimonio)
$groupBox1.Controls.Add($labelMarca)
$groupBox1.Controls.Add($obrigatorioMarca)
$groupBox1.Controls.Add($marca)
$groupBox1.Controls.Add($labelModelo)
$groupBox1.Controls.Add($obrigatorioModelo)
#==================================================================================

$groupBox1.Controls.Add($labelPatrimonio)
$groupBox1.Controls.Add($txtPatrimonio)
$groupBox1.Controls.Add($labelMarca)
$groupBox1.Controls.Add($marca)
$groupBox1.Controls.Add($labelModelo)
$groupBox1.Controls.Add($modelo)
$groupBox1.Controls.Add($labelNumeroSerie)
$groupBox1.Controls.Add($txtNumeroSerie)
$groupBox1.Controls.Add($labelSituacao)
$groupBox1.Controls.Add($obrigatorioSituacao)
$groupBox1.Controls.Add($situacao)
$groupBox1.Controls.Add($labelDescricao)
$groupBox1.Controls.Add($obrigatorioDescricao)
$groupBox1.Controls.Add($txtDescricao)
$groupBox1.Controls.Add($labelUsuario)
$groupBox1.Controls.Add($obrigatorioUsuario)
$groupBox1.Controls.Add($txtUsuario)
$groupBox1.Controls.Add($labelSetor)
$groupBox1.Controls.Add($obrigatorioSetor)
$groupBox1.Controls.Add($txtSetor)
$groupBox1.Controls.Add($labelSala)
$groupBox1.Controls.Add($obrigatorioSala)
$groupBox1.Controls.Add($situacao)
$groupBox1.Controls.Add($labelDescricao)
$groupBox1.Controls.Add($txtDescricao)
$groupBox1.Controls.Add($labelUsuario)
$groupBox1.Controls.Add($txtUsuario)
$groupBox1.Controls.Add($labelSetor)
$groupBox1.Controls.Add($txtSetor)
$groupBox1.Controls.Add($labelSala)
$groupBox1.Controls.Add($txtSala)
$groupBox1.Controls.Add($labelRamal)
$groupBox1.Controls.Add($txtRamal)
$groupBox1.Controls.Add($btnSalvar)
$groupBox1.Controls.Add($btnSalvarImagem)
$form.Controls.Add($groupBox1)


#==================================================================================

$btnSalvar.Add_Click({
    $validacao = $true
    $mensagem = @()

    if ([string]::IsNullOrWhiteSpace($txtPatrimonio.Text)) {
        $validacao = $false
        $mensagem += "O campo Patrimônio é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($marca.Text)) {
        $validacao = $false
        $mensagem += "O campo Marca é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($modelo.Text)) {
        $validacao = $false
        $mensagem += "O campo Modelo é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($situacao.Text)) {
        $validacao = $false
        $mensagem += "O campo Situação é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($txtDescricao.Text)) {
        $validacao = $false
        $mensagem += "O campo Descrição é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($txtUsuario.Text)) {
        $validacao = $false
        $mensagem += "O campo Usuário é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($txtSetor.Text)) {
        $validacao = $false
        $mensagem += "O campo Setor é obrigatório."
    }

    if ([string]::IsNullOrWhiteSpace($txtSala.Text)) {
        $validacao = $false
        $mensagem += "O campo Sala é obrigatório."
    }

    # Exibe mensagens de erro ou sucesso
    if (-not $validacao) {
        [System.Windows.Forms.MessageBox]::Show($mensagem -join "`n", "Campos Obrigatórios", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        if($resultPatrimonio -eq [System.Windows.Forms.DialogResult]::OK){

            # PSCustomObject - Criando objeto de chave-valor (melhor para manipulação)
            $novoPatrimonio = [PSCustomObject]@{
                Patrimônio             = $txtPatrimonio.Text.Trim()
                Marca                  = $marca.Text.Trim()
                Modelo                 = $modelo.Text.Trim()
                Número_de_Série        = $txtNumeroSerie.Text.Trim()
                Situação               = $situacao.Text.Trim()
                Descrição              = $txtDescricao.Text.Trim()
                Usuário                = $txtUsuario.Text.Trim()
                Setor                  = $txtSetor.Text.Trim()
                Sala                   = $txtSala.Text.Trim()
                Ramal                  = $txtRamal.Text.Trim()
                Data_de_Cadastro       = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        
            }

            # Obtém o caminho para a área de trabalho do usuário
            $caminhoDesktop = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop')

            # Nome do arquivo CSV
            $arquivoCsv = "patrimonio.csv"

            # Caminho completo do arquivo CSV na área de trabalho
            $caminhoArquivoCsv = [System.IO.Path]::Combine($caminhoDesktop, $arquivoCsv)

            # Lendo os dados existentes do CSV
            #$arquivoPatrimonio = "C:\Users\gabriel.vinicius\Desktop\patrimonio.csv"
            $dadosExistentes = @()

            if (Test-Path $caminhoArquivoCsv) {
                $dadosExistentes = Import-Csv -Path $caminhoArquivoCsv -Delimiter ";" -Encoding UTF8
            }

            # Verifica se o patrimônio já existe
            $patrimonioIgual = $dadosExistentes | Where-Object { $_.Patrimônio.Trim() -ieq $novoPatrimonio.Patrimônio }

            if ($patrimonioIgual) {
                [System.Windows.Forms.MessageBox]::Show("O patrimônio '$($novoPatrimonio.Patrimônio)' já está cadastrado!", "Notificação", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # Se não existir, salva o novo patrimônio
                $novoPatrimonio | Export-Csv -Append -Path $caminhoArquivoCsv -NoTypeInformation -Delimiter ";" -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Patrimônio cadastrado com sucesso!", "Notificação", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }

        #[System.Windows.Forms.MessageBox]::Show("Todos os campos foram preenchidos com sucesso!", "Notificação", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        #$form.Close()  # Fecha o formulário se todos os campos estiverem preenchidos
    }
})

#==================================================================================


$groupBox2 = New-Object System.Windows.Forms.GroupBox
$groupBox2.Text = "Pesquisar Patrimônio"
$groupBox2.Text = "Busca de Patrimônio"
$groupBox2.Font = $fonteGroupBox
#$groupBox2.BackColor = '#000000'
#$groupBox2.ForeColor = [System.Drawing.Color]::White
$groupBox2.Location = New-Object System.Drawing.Point(200, 750)
$groupBox2.Size = New-Object System.Drawing.Size(800, 200)

# Criando a caixa de texto para a busca
$txtBusca = New-Object System.Windows.Forms.TextBox
$txtBusca.Font = $fonteTexto
$txtBusca.MaxLength = 7
$txtBusca.Location = New-Object System.Drawing.Point(150, 75)
$txtBusca.Size = New-Object System.Drawing.Size(550, 300)
$form.Controls.Add($txtBusca)

$txtBusca.Add_KeyDown({
    param($sender, $e)
    
    # Verifica se a tecla pressionada foi Enter (código da tecla Enter é 13)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        # Simula o clique no botão
        $btnBuscar.PerformClick()
    }
})

# Criando o botão para realizar a busca
$btnBuscar = New-Object System.Windows.Forms.Button
$btnBuscar.Text = 'Pesquisar'
$btnBuscar.Text = 'Buscar'
$btnBuscar.Font = $fonteTexto
#$btnBuscar.ForeColor = [System.Drawing.Color]::Black
#$btnBuscar.BackColor = [System.Drawing.Color]::White
$btnBuscar.Location = New-Object System.Drawing.Point(300, 125)
$btnBuscar.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($btnBuscar)

$groupBox2.Controls.Add($txtBusca)
$groupBox2.Controls.Add($btnBuscar)
$form.Controls.Add($groupBox2)


#=========================================================================

$btnSalvarImagem.Add_Click({

        # Criar o formulário para salvar a imagem
        $formSalvarImagem = New-Object System.Windows.Forms.Form
        $formSalvarImagem.Text = "Salvar Imagem"
        $formSalvarImagem.Size = New-Object System.Drawing.Size(900, 700)

        $logoFormImagem = New-Object System.Windows.Forms.PictureBox
        $logoFormImagem.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
        $caminhoLogoFormImagem = "C:\Users\gabriel.vinicius\Desktop\censipam-logo.png"
        $logoFormImagem.Image = [System.Drawing.Image]::FromFile($caminhoLogoFormImagem)
        $logoFormImagem.Location = New-Object System.Drawing.Point(300, 10)
        $logoFormImagem.Size = New-Object System.Drawing.Size(250,100)
        $formSalvarImagem.Controls.Add($logoFormImagem)

        $fonteTextoFormImagem = New-Object System.Drawing.Font("Arial", 16)

        # Criar o Label para instrução
        $labelImagem = New-Object System.Windows.Forms.Label
        $labelImagem.Text = "Nome da Imagem:"
        $labelImagem.Font = $fonteTextoFormImagem
        $labelImagem.Location = New-Object System.Drawing.Point(50, 152)
        $labelImagem.Size = New-Object System.Drawing.Size(200, 40)
        $formSalvarImagem.Controls.Add($labelImagem)

        # Criar o TextBox para o nome da imagem
        $textBoxImagem = New-Object System.Windows.Forms.TextBox
        $textBoxImagem.Font = $fonteTextoFormImagem
        $textBoxImagem.Location = New-Object System.Drawing.Point(250, 150)
        $textBoxImagem.Size = New-Object System.Drawing.Size(350, 150)
        $formSalvarImagem.Controls.Add($textBoxImagem)

        # Criar o PictureBox para mostrar a imagem
        $pictureBox = New-Object System.Windows.Forms.PictureBox
        $pictureBox.Size = New-Object System.Drawing.Size(400, 200)
        $pictureBox.Location = New-Object System.Drawing.Point(250, 300)
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
        $formSalvarImagem.Controls.Add($pictureBox)

        # Criar o botão para escolher a imagem
        $btnEscolherImagem = New-Object System.Windows.Forms.Button
        $btnEscolherImagem.Text = "Escolher Imagem"
        $btnEscolherImagem.Font = $fonteTextoFormImagem
        $btnEscolherImagem.Location = New-Object System.Drawing.Point(620, 145)
        $btnEscolherImagem.Size = New-Object System.Drawing.Size(200, 40)
        $formSalvarImagem.Controls.Add($btnEscolherImagem)

        # Função para escolher a imagem
        $btnEscolherImagem.Add_Click({
            # Criar o diálogo de abrir arquivo
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "Imagens|*.jpg;*.jpeg;*.png;*.bmp"
            $openFileDialog.Title = "Selecione uma Imagem"
        
            # Se o usuário escolher um arquivo, exibe no PictureBox
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $pictureBox.Image = [System.Drawing.Image]::FromFile($openFileDialog.FileName)
                $selectedImagePath = $openFileDialog.FileName
            }
        })

        # Criar o botão para salvar a imagem
        $btnSalvarImg = New-Object System.Windows.Forms.Button
        $btnSalvarImg.Text = "Salvar Imagem"
        $btnSalvarImg.Font = $fonteTextoFormImagem
        $btnSalvarImg.Location = New-Object System.Drawing.Point(350, 580)
        $btnSalvarImg.Size = New-Object System.Drawing.Size(200, 40)
        $formSalvarImagem.Controls.Add($btnSalvarImg)

        # Função ao clicar no botão salvar
        $btnSalvarImg.Add_Click({
            $imageName = $textBoxImagem.Text
            if ($imageName -and $pictureBox.Image) {
                # Caminho para a área de trabalho e a nova pasta "imagens-equipamentos"
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $folderPath = Join-Path -Path $desktopPath -ChildPath "imagens-equipamentos"
            
                # Verificar se a pasta "imagens-equipamentos" existe, caso contrário, criar
                if (-not (Test-Path $folderPath)) {
                    New-Item -ItemType Directory -Path $folderPath
                }

                # Definir o caminho da imagem a ser salva
                $imagePath = Join-Path -Path $folderPath -ChildPath "$imageName.jpg"

                # Salvar a imagem na pasta especificada
                $pictureBox.Image.Save($imagePath)

                # Mostrar mensagem de sucesso
                [System.Windows.Forms.MessageBox]::Show("Imagem salva com sucesso!", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK)
                $formSalvarImagem.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Por favor, insira o nome da imagem e escolha uma imagem.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK)
            }
        })

        # Exibir o formulário para salvar a imagem
        $formSalvarImagem.ShowDialog()
    
})
#==================================================================================

# Função para buscar no arquivo CSV
$btnBuscar.Add_Click({
    $busca = $txtBusca.Text
    if ($busca) {
    
        # Obtém o caminho para a área de trabalho do usuário
        $caminhoDesktop = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop')

        # Nome do arquivo CSV
        $arquivoCsv = "patrimonio.csv"

        # Caminho completo do arquivo CSV na área de trabalho
        $caminhoArquivoCsv = [System.IO.Path]::Combine($caminhoDesktop, $arquivoCsv)
                
        # Verificando se o arquivo CSV existe
        if (Test-Path $caminhoArquivoCsv) {
            # Importando o CSV como objetos
            $dadosCsv = Import-Csv -Path $caminhoArquivoCsv -Delimiter ";" -Encoding UTF8
            $dadosCsv[0] | Format-List *

            # Buscando a linha com base no termo de busca (por exemplo, o patrimônio)
            $linha = $dadosCsv | Where-Object { $_.Patrimônio.Trim() -eq $busca }

            
            if ($linha) {
                    # Criar o formulário para salvar a imagem
                    $formRecuperarDados = New-Object System.Windows.Forms.Form
                    $formRecuperarDados.Text = "Recuperando Dados do Patrimônio"
                    $formRecuperarDados.Size = New-Object System.Drawing.Size(1100, 900)

                    $logoFormDados = New-Object System.Windows.Forms.PictureBox
                    $logoFormDados.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
                    $caminhoLogoFormDados = "C:\Users\gabriel.vinicius\Desktop\censipam-logo.png"
                    $logoFormDados.Image = [System.Drawing.Image]::FromFile($caminhoLogoFormDados)
                    $logoFormDados.Location = New-Object System.Drawing.Point(420, 10)
                    $logoFormDados.Size = New-Object System.Drawing.Size(250,100)
                    $formRecuperarDados.Controls.Add($logoFormDados)

                    $fonteTextoDados = New-Object System.Drawing.Font("Arial", 16)

                    $labelPat = New-Object System.Windows.Forms.Label
                    $labelPat.Text = "Patrimônio: ", $($linha.Patrimônio)
                    $labelPat.Font = $fonteTextoDados
                    $labelPat.Location = New-Object System.Drawing.Point(60, 200)
                    $labelPat.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelPat)

                    $labelMar = New-Object System.Windows.Forms.Label
                    $labelMar.Text = "Marca: ", $($linha.Marca)
                    $labelMar.Font = $fonteTextoDados
                    $labelMar.Location = New-Object System.Drawing.Point(60, 230)
                    $labelMar.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelMar)

                    $labelMod = New-Object System.Windows.Forms.Label
                    $labelMod.Text = "Modelo: ", $($linha.Modelo)
                    $labelMod.Font = $fonteTextoDados
                    $labelMod.Location = New-Object System.Drawing.Point(60, 260)
                    $labelMod.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelMod)

                    $labelNum = New-Object System.Windows.Forms.Label
                    $labelNum.Text = "Número de Série: ", $($linha.Número_de_Série)
                    $labelNum.Font = $fonteTextoDados
                    $labelNum.Location = New-Object System.Drawing.Point(60, 290)
                    $labelNum.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelNum)

                    $labelSit = New-Object System.Windows.Forms.Label
                    $labelSit.Text = "Situação: ", $($linha.Situação)
                    $labelSit.Font = $fonteTextoDados
                    $labelSit.Location = New-Object System.Drawing.Point(60, 320)
                    $labelSit.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelSit)

                    $labelDes = New-Object System.Windows.Forms.Label
                    $labelDes.Text = "Descrição: ", $($linha.Descrição)
                    $labelDes.Font = $fonteTextoDados
                    $labelDes.Location = New-Object System.Drawing.Point(60, 350)
                    $labelDes.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelDes)

                    $labelUsu = New-Object System.Windows.Forms.Label
                    $labelUsu.Text = "Usuário: ", $($linha.Usuário)
                    $labelUsu.Font = $fonteTextoDados
                    $labelUsu.Location = New-Object System.Drawing.Point(60, 380)
                    $labelUsu.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelUsu)

                    $labelSet = New-Object System.Windows.Forms.Label
                    $labelSet.Text = "Setor: ", $($linha.Setor)
                    $labelSet.Font = $fonteTextoDados
                    $labelSet.Location = New-Object System.Drawing.Point(60, 410)
                    $labelSet.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelSet)

                    $labelSal = New-Object System.Windows.Forms.Label
                    $labelSal.Text = "Sala: ", $($linha.Sala)
                    $labelSal.Font = $fonteTextoDados
                    $labelSal.Location = New-Object System.Drawing.Point(60, 440)
                    $labelSal.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelSal)

                    $labelRam = New-Object System.Windows.Forms.Label
                    $labelRam.Text = "Ramal: ", $($linha.Ramal)
                    $labelRam.Font = $fonteTextoDados
                    $labelRam.Location = New-Object System.Drawing.Point(60, 470)
                    $labelRam.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelRam)

                    $labelDat = New-Object System.Windows.Forms.Label
                    $labelDat.Text = "Data de Cadastro: ", $($linha.Data_de_Cadastro)
                    $labelDat.Font = $fonteTextoDados
                    $labelDat.Location = New-Object System.Drawing.Point(60, 500)
                    $labelDat.Size = New-Object System.Drawing.Size(500, 30)
                    $formRecuperarDados.Controls.Add($labelDat)

                    $imagemEquip = New-Object System.Windows.Forms.PictureBox
                    $imagemEquip.Size = New-Object System.Drawing.Size(400, 200)
                    $imagemEquip.Location = New-Object System.Drawing.Point(350, 600)
                    #$imagemEquip.Image = [System.Drawing.Image]::FromFile($caminhoImagem.FileName)
                    $imagemEquip.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
                    $formRecuperarDados.Controls.Add($imagemEquip)

                    

                    # Caminho da pasta onde as imagens estão localizadas
                    $caminhoPastaImagens = "C:\Users\gabriel.vinicius\Desktop\imagens-equipamentos"

                    $modeloTexto = $($linha.Modelo)

                    # Verificar se o campo de texto não está vazio
                    if (-not [string]::IsNullOrWhiteSpace($modeloTexto)) {
                        # Procurar pelo arquivo de imagem com o nome correspondente ao texto do TextBox
                        $caminhoImagem = Get-ChildItem -Path $caminhoPastaImagens -Filter "$modeloTexto*.*" | Where-Object { $_.Extension -match "\.jpg|\.jpeg|\.png|\.bmp|\.gif" }

                        # Se a imagem foi encontrada, exibir no PictureBox
                        if ($caminhoImagem) {
                            $imagemEquip.Image = [System.Drawing.Image]::FromFile($caminhoImagem.FullName)
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Imagem não encontrada!", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                    }


                    # Exibir o formulário para salvar a imagem
                    $formRecuperarDados.ShowDialog()
                
                #$msg = "Patrimônio: '$($linha.Patrimônio)',`nMarca: '$($linha.Marca)',`nModelo: '$($linha.Modelo)',`nNúmero de Série: '$($linha.Número_de_Série)',`nSituação: '$($linha.Situação)',`nDescrição: '$($linha.Descrição)',`nUsuário: '$($linha.Usuário)',`nSetor: '$($linha.Setor)',`nSala: '$($linha.Sala)',`nRamal: '$($linha.Ramal)',`nData de Cadastro: '$($linha.Data_de_Cadastro)'"
                #[System.Windows.Forms.MessageBox]::Show($msg, "Patrimônio encontrado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $msg = "Patrimônio: '$($linha.Patrimônio)',`nMarca: '$($linha.Marca)',`nModelo: '$($linha.Modelo)',`nNúmero de Série: '$($linha.Número_de_Série)',`nSituação: '$($linha.Situação)',`nDescrição: '$($linha.Descrição)',`nUsuário: '$($linha.Usuário)',`nSetor: '$($linha.Setor)',`nSala: '$($linha.Sala)',`nRamal: '$($linha.Ramal)',`nData de Cadastro: '$($linha.Data_de_Cadastro)'"
                [System.Windows.Forms.MessageBox]::Show($msg, "Patrimônio encontrado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Nenhuma linha encontrada com o patrimônio: $busca")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Arquivo CSV não encontrado.")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Por favor, insira um patrimônio para buscar.")
    }
})

$linha | Format-List *

#==================================================================================
# Exibindo o formulário
$form.Add_Shown({$form.Activate()})
$resultPatrimonio = $form.ShowDialog()

}