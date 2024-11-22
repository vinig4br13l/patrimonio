# Importando o assembly do Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

$txtPatrimonio = New-Object System.Windows.Forms.TextBox
$txtPatrimonio.Font = $fonteTexto
$txtPatrimonio.MaxLength = 7
$txtPatrimonio.Location = New-Object System.Drawing.Point(190, 50)
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

$situacao = New-Object System.Windows.Forms.ComboBox
$situacao.Location = New-Object System.Drawing.Point(190, 250)
$situacao.Size = New-Object System.Drawing.Size(550,300)
$situacao.Font = $fonteTexto

$situacao.Items.Add("Em uso")
$situacao.Items.Add("Disponível")
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
$btnSalvar.Text = "Salvar em CSV"
$btnSalvar.AutoSize = $true
$btnSalvar.Font = $fonteTexto
#$btnSalvar.BackColor = [System.Drawing.Color]::White
#$btnSalvar.ForeColor = [System.Drawing.Color]::Black
$btnSalvar.Location = New-Object System.Drawing.Point(300, 550)
$btnSalvar.Size = New-Object System.Drawing.Size(200,40)
$btnSalvar.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnSalvar)

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
$groupBox2.Text = "Busca de Patrimônio"
$groupBox2.Font = $fonteGroupBox
#$groupBox2.BackColor = '#000000'
#$groupBox2.ForeColor = [System.Drawing.Color]::White
$groupBox2.Location = New-Object System.Drawing.Point(200, 750)
$groupBox2.Size = New-Object System.Drawing.Size(800, 200)

# Criando a caixa de texto para a busca
$txtBusca = New-Object System.Windows.Forms.TextBox
$txtBusca.Font = $fonteTexto
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