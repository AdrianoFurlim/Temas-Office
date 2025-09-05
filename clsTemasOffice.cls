VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemasOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------
' CLASSE: clsTemasOffice - Sistema de Temas e Efeitos Visuais
' DESCRIÇÃO:
'   Classe para aplicação de temas visuais personalizados em UserForms do VBA,
'   incluindo efeitos de hover (sobreposição) e placeholders (texto de dica).
'
' COMO USAR:
' 1. DECLARAÇÃO:
'    ' Para temas principais
'    Dim GTemas As New clsTemasOffice
'
'    ' Para efeitos individuais (hover/placeholder)
'    Dim Efeitos(1 To N) As New clsTemasOffice  ' Onde N é o número de controles com efeitos
'
' 2. INICIALIZAÇÃO (no UserForm_Initialize):
'    GTemas.AplicarTema Me  ' Aplica o tema automático ao formulário
'
' 3. CONFIGURAÇÃO DE CONTROLES:
'    ' Definir tags de estilo para cada controle
'    Me.Label1.Tag = "{tmpr:estilo1}:{tmbr:estilo2}:{tmci:estilo3}:{tmco:estilo4}"
'
'    ' Aplicar efeitos hover/placeholder
'    Efeitos(1).EfeitoHover Me, Me.Label1, "tag_estilo_hover", Me.FrameFundo
'    Efeitos(2).EfeitoPlaceHolder Me, Me.TextBox1, "Digite seu nome"
'
' AUTOR: Adriano Furtado Lima
'--------------------------------------------------------------
'--------------------------------------------------------------
' SISTEMA DE TAGS - FORMATO E SINTAXE
'
' ESTRUTURA BÁSICA:
'   "{tema:estilo}:{tema:estilo}:..."
'   (Um estilo para cada tema suportado)
'
' TEMAS DISPONÍVEIS:
'   tmpr = Tema Preto
'   tmbr = Tema Branco
'   tmci = Tema Cinza Escuro
'   tmco = Tema Colorido
'
' TIPOS DE ESTILO:
'
' 1. ESTILOS PRÉ-DEFINIDOS (nomes curtos):
'   Ex: "lblfcd01", "lblbtd01", etc.
'   Lista:
'   - lblfcd01-05: Label com fundo colorido (destaque D1-D5)
'   - lblbtd01: Label com borda de destaque
'   - lblbts01: Label com borda padrão
'   - lbltds01: Label transparente com texto grande
'   - lbltsp01: Label transparente com texto pequeno
'
' 2. ESTILOS MODULARES (propriedades individuais):
'   Formato: "(prop1=valor1|prop2=valor2)"
'   Propriedades disponíveis:
'   - bc: BackColor (cor de fundo - F1-F5, D1-D5 ou cor hexadecimal)
'   - fc: ForeColor (cor do texto - L1-L5 ou cor hexadecimal)
'   - bd: BorderColor (cor da borda - B1-B5 ou cor hexadecimal)
'   - bs: BackStyle (0=transparente, 1=opaco)
'   - by: BorderStyle (0=none, 1=single)
'   - fs: Font.Size (tamanho da fonte)
'
' EXEMPLOS DE TAGS:
'
' 1. Estilo pré-definido para todos os temas:
'   "{tmpr:lblbtd01}:{tmbr:lblbtd01}:{tmci:lblbtd01}:{tmco:lblbtd01}"
'
' 2. Estilo modular personalizado:
'   "{tmpr(bc=d2|bs=1|bd=d1|by=1|fc=l1)}:{tmbr(bc=f3|bs=0|fc=l2)}"
'
' 3. Combinação de estilos:
'   "{tmpr:lblfcd01}:{tmbr(bc=f3|fc=l1)}:{tmci:lbltds01}"
'
' 4. Cores hexadecimais:
'   "{tmpr(bc=#FF0000|fc=&H00FF00)}" (vermelho fundo, verde texto)
'--------------------------------------------------------------
'--------------------------------------------------------------
' EXEMPLO COMPLETO DE IMPLEMENTAÇÃO
'
' 1. DECLARAÇÕES:
'    Option Explicit
'
'    ' Instância para o tema principal
'    Dim GTemas As New clsTemasOffice
'
'    ' Instâncias para efeitos individuais
'    Dim Efeitos(1 To 5) As New clsTemasOffice
'
' 2. INICIALIZAÇÃO (UserForm_Initialize):
'    Private Sub UserForm_Initialize()
'        ' Aplicar tema automático ao formulário
'        GTemas.AplicarTema Me
'
'        ' Configurar tags para controles
'        Me.Label1.Tag = "{tmpr:lblfcd01}:{tmbr:lblfcd01}:{tmci:lblfcd01}:{tmco:lblfcd01}"
'        Me.TextBox1.Tag = "{tmpr(bc=d2|bs=1|bd=d1|by=1|fc=l1)}:{tmbr:lblbtd01}"
'
'        ' Configurar efeitos hover
'        Efeitos(1).EfeitoHover Me, Me.Label1, "{tmpr:lblbtd11}:{tmbr:lblbtd11}", Me.Frame1
'
'        ' Configurar placeholder
'        Efeitos(2).EfeitoPlaceHolder Me, Me.TextBox1, "Digite seu texto"
'    End Sub
'
' 3. TROCA DE TEMA DINÂMICA:
'    ' Para alterar o tema em tempo de execução:
'    Private Sub AlterarParaTemaPreto()
'        GTemas.AplicarTema Me, TemaPreto  ' Usa a enumeração Temas
'    End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------
' MELHORES PRÁTICAS:
'
' 1. ORGANIZAÇÃO:
'   - Use GTemas para configurações globais
'   - Use Efeitos() para controles individuais com hover/placeholder
'
' 2. NOMENCLATURA:
'   - Prefixos recomendados para tags:
'     * lbl_ para Labels
'     * btn_ para CommandButtons
'     * txt_ para TextBoxes
'     * cmb_ para ComboBoxes
'
' 3. PERFORMANCE:
'   - Evite muitos controles com efeitos hover simultâneos
'   - Para formulários complexos, considere aplicar temas apenas
'     aos controles visíveis inicialmente
'
' 4. MANUTENÇÃO:
'   - Mantenha uma biblioteca de estilos padrão
'   - Documente os estilos customizados
'--------------------------------------------------------------
'--------------------------------------------------------------
' LIMITAÇÕES CONHECIDAS:
' 1. Suporta apenas os controles padrão do MSForms
' 2. Efeitos hover podem não funcionar perfeitamente em:
'    - Controles dentro de MultiPage
'    - Controles com ZOrder modificado dinamicamente
' 3. Placeholders funcionam apenas em TextBox e ComboBox
'--------------------------------------------------------------








'--------------------------------------------------------------
' CLASSE: clsTemasOffice
' DESCRIÇÃO: Classe para gerenciamento completo de temas visuais em UserForms VBA,
'            incluindo aplicação de temas, efeitos hover e placeholders.
'
' COMO USAR:
' 1. Instancie a classe no UserForm:
'    Private GTemas As New clsTemasOffice
'
' 2. Aplique um tema principal:
'    GTemas.AplicarTema Me, [TemaEscolido]
'
' 3. Para controles com efeito hover:
'    GTemas.EfeitoHover Me, Me.Label1, "{tmpr:estilo}", Me.Frame1
'
' 4. Para placeholders:
'    GTemas.EfeitoPlaceHolder Me, Me.TextBox1, "Digite aqui"
'
' AUTOR: Adriano Furtado Lima
'--------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------
' VARIÁVEIS DE EVENTO
' Controles com eventos monitorados para efeitos visuais
'--------------------------------------------------------------
Private WithEvents Formulario       As MSForms.UserForm
Attribute Formulario.VB_VarHelpID = -1
Private WithEvents CtrlCheckBox     As MSForms.CheckBox
Attribute CtrlCheckBox.VB_VarHelpID = -1
Private WithEvents CtrlComboBox     As MSForms.ComboBox
Attribute CtrlComboBox.VB_VarHelpID = -1
Private WithEvents CtrlCommandBtn   As MSForms.CommandButton
Attribute CtrlCommandBtn.VB_VarHelpID = -1
Private WithEvents CtrlFrame        As MSForms.Frame
Attribute CtrlFrame.VB_VarHelpID = -1
Private WithEvents CtrlImage        As MSForms.Image
Attribute CtrlImage.VB_VarHelpID = -1
Private WithEvents CtrlLabel        As MSForms.Label
Attribute CtrlLabel.VB_VarHelpID = -1
Private WithEvents CtrlListBox      As MSForms.ListBox
Attribute CtrlListBox.VB_VarHelpID = -1
Private WithEvents CtrlMultiPage    As MSForms.MultiPage
Attribute CtrlMultiPage.VB_VarHelpID = -1
Private WithEvents CtrlOptionBtn    As MSForms.OptionButton
Attribute CtrlOptionBtn.VB_VarHelpID = -1
Private WithEvents CtrlTabStrip     As MSForms.TabStrip
Attribute CtrlTabStrip.VB_VarHelpID = -1
Private WithEvents CtrlTextBox      As MSForms.TextBox
Attribute CtrlTextBox.VB_VarHelpID = -1
Private WithEvents CtrlToggleBtn    As MSForms.ToggleButton
Attribute CtrlToggleBtn.VB_VarHelpID = -1

Private WithEvents FndLabel         As MSForms.Label
Attribute FndLabel.VB_VarHelpID = -1
Private WithEvents FndFrame         As MSForms.Frame
Attribute FndFrame.VB_VarHelpID = -1

'--------------------------------------------------------------
' CONSTANTES E VARIÁVEIS DE CLASSE
'--------------------------------------------------------------
Private Const FontePadrao = "Segoe UI" ' Fonte padrão para todos os controles

Private TemaAtual       As TemaCores        ' Armazena o tema atualmente aplicado
Private TagEmExec       As String           ' Tag do controle em execução para efeitos
Private UltContro       As MSForms.Control  ' Último controle com hover ativo
Private CtrlPlaceHolder As MSForms.TextBox  ' Controle de placeholder
Private PlaceHolder     As Boolean          ' Flag indicando se placeholder está ativo

'--------------------------------------------------------------
' ENUMERAÇÃO: Temas
' Descrição: Define os temas visuais disponíveis para aplicação
' Valores:
'   TemaAutomatico   = -1  ' Segue o tema do Office
'   TemaCinzaEscuro  = 3   ' Tema escuro em tons de cinza
'   TemaPreto        = 4   ' Tema preto completo
'   TemaBranco       = 5   ' Tema claro padrão
'   TemaConfWindows  = 6   ' Segue o tema do Windows
'   TemaColorido     = 7   ' Tema colorido padrão
'--------------------------------------------------------------
Public Enum Temas
    TemaAutomatico = -1   ' Tema acompanha o tema atual do Excel
    TemaCinzaEscuro = 3   ' Tema baseado no modo cinza escuro
    TemaPreto = 4         ' Tema em estilo escuro total
    TemaBranco = 5        ' Tema claro padrão
    TemaConfWindows = 6   ' Tema com base nas configurações do Windows
    TemaColorido = 7      ' Tema colorido padrão do Excel
End Enum

'--------------------------------------------------------------
' TIPO: TemaCores
' Descrição: Estrutura que armazena as paletas de cores para cada tema
' Campos:
'   Nome - Identificador do tema
'   F1-F5 - Cores de fundo (F = Fundo)
'   S1-S5 - Cores de seleção (S = Seleção)
'   D1-D5 - Cores de destaque (D = Destaque)
'   L1-L5 - Cores de texto (L = Letra)
'   B1-B5 - Cores de borda (B = Borda)
'--------------------------------------------------------------
Private Type TemaCores
    Nome   As String
    F1  As Long: F2 As Long: F3 As Long: F4 As Long: F5 As Long
    S1  As Long: S2 As Long: S3 As Long: S4 As Long: S5 As Long
    D1  As Long: D2 As Long: D3 As Long: D4 As Long: D5 As Long
    L1  As Long: L2 As Long: L3 As Long: L4 As Long: L5 As Long
    B1  As Long: B2 As Long: B3 As Long: B4 As Long: B5 As Long
End Type

' Variáveis de instância para cada tema
Private TmColorido      As TemaCores
Private TmCinzaEscuro   As TemaCores
Private TmPreto         As TemaCores
Private TmBranco        As TemaCores

'--------------------------------------------------------------
' MÉTODO: Class_Initialize
' Descrição: Inicializa os temas com suas paletas de cores
' Chamado automaticamente quando a classe é instanciada
'--------------------------------------------------------------
Private Sub Class_Initialize()
    ' Configuração do tema colorido
    With TmColorido
        .Nome = "tmco"
        .F1 = &HF0F0F0: .F2 = &HF5F5F5: .F3 = &HFFFFFF: .F4 = &HFAFAFA: .F5 = &HF7F7F7
        .S1 = &HEBEBEB: .S2 = &HEDEDED: .S3 = &HEFEFEF: .S4 = &HF1F1F1: .S5 = &HF3F3F3
        .D1 = &H417C10: .D2 = &H5BAA1F: .D3 = &H76D130: .D4 = &H91E840: .D5 = &HA3FA5A
        .L1 = &H242424: .L2 = &H444444: .L3 = &H666666: .L4 = &H888888: .L5 = &HAAAAAA
        .B1 = &H808080: .B2 = &H999999: .B3 = &HB0B0B0: .B4 = &HC8C8C8: .B5 = &HDDDDDD
    End With

    With TmCinzaEscuro
        .Nome = "tmci"
        .F1 = &H525252: .F2 = &H6A6A6A: .F3 = &H2E2E2E: .F4 = &H3C3C3C: .F5 = &H484848
        .S1 = &H616161: .S2 = &H707070: .S3 = &H7E7E7E: .S4 = &H8D8D8D: .S5 = &H9C9C9C
        .D1 = &H417C10: .D2 = &H5BAA1F: .D3 = &H76D130: .D4 = &H91E840: .D5 = &HA3FA5A
        .L1 = &H242424: .L2 = &HE0E0E0: .L3 = &HFAFAFA: .L4 = &HFCFCFC: .L5 = &HFEFEFE
        .B1 = &H8A8A8A: .B2 = &HA0A0A0: .B3 = &HB5B5B5: .B4 = &HCACACA: .B5 = &HDFDFDF
    End With

    With TmPreto
        .Nome = "tmpr"
        .F1 = &HA0A0A:  .F2 = &H292929: .F3 = &H3D3D3D: .F4 = &H505050: .F5 = &H636363
        .S1 = &H3D3D3D: .S2 = &H4A4A4A: .S3 = &H585858: .S4 = &H666666: .S5 = &H747474
        .D1 = &HFE6B4D: .D2 = &HFF825F: .D3 = &HFF9972: .D4 = &HFFAF85: .D5 = &HFFC598
        .L1 = &HEFEFEF: .L2 = &HD0D0D0: .L3 = &HB0B0B0: .L4 = &H909090: .L5 = &H707070
        .B1 = &H757575: .B2 = &H8A8A8A: .B3 = &HA0A0A0: .B4 = &HB5B5B5: .B5 = &HCACACA
    End With

    With TmBranco
        .Nome = "tmbr"
        .F1 = &HF3F3F3: .F2 = &HF5F5F5: .F3 = &HFFFFFF: .F4 = &HFCFCFC: .F5 = &HF8F8F8
        .S1 = &HEBEBEB: .S2 = &HEDEDED: .S3 = &HEFEFEF: .S4 = &HF1F1F1: .S5 = &HF3F3F3
        .D1 = &H234EF0: .D2 = &H3B66F5: .D3 = &H5380FA: .D4 = &H6C98FD: .D5 = &H85B0FF
        .L1 = &H242424: .L2 = &H444444: .L3 = &H666666: .L4 = &H888888: .L5 = &HAAAAAA
        .B1 = &H808080: .B2 = &H999999: .B3 = &HB0B0B0: .B4 = &HC7C7C7: .B5 = &HDEDEDE
    End With
End Sub

'--------------------------------------------------------------
' MÉTODO: AplicarTema
' Descrição: Aplica um tema visual completo a um UserForm, incluindo todos os seus controles,
'            formatando cores, fontes e estilos conforme o tema selecionado.
'
' Parâmetros:
'   - Form: Referência obrigatória ao UserForm que receberá o tema (MSForms.UserForm)
'   - TemaEscolido: Valor opcional da enumeração Temas que define o tema a ser aplicado.
'     Valores possíveis:
'       - TemaAutomatico (-1): Usa o tema atual do Office
'       - TemaCinzaEscuro (3): Tema escuro em tons de cinza
'       - TemaPreto (4): Tema preto completo
'       - TemaBranco (5): Tema claro padrão
'       - TemaConfWindows (6): Usa o tema do Windows
'       - TemaColorido (7): Tema colorido padrão
'     Se omitido, padrão é TemaAutomatico (-1)
'
' Retorno:
'   - Nenhum (método Sub)
'
' Fluxo:
'   1. Define o tema atual chamando DefinirTema()
'   2. Aplica a formatação visual em todo o formulário com FormatarFormulario()
'
' Observações:
'   - Formata todos os controles contidos no UserForm
'   - Utiliza a propriedade .Tag dos controles para estilos específicos
'   - Define a fonte padrão "Segoe UI" tamanho 10 para todos os controles
'   - Se ocorrer erro na aplicação, continua com próximo controle (On Error Resume Next)
'
' Exemplo de uso:
'   GTemas.AplicarTema Me                  ' Usa tema automático
'   GTemas.AplicarTema Me, TemaPreto       ' Aplica tema preto
'   GTemas.AplicarTema Me, TemaColorido    ' Aplica tema colorido
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Public Sub AplicarTema(ByRef Form As MSForms.UserForm, Optional ByVal TemaEscolido As Temas = -1)
    DefinirTema TemaEscolido
    
    ' Aplica a formatação no formulário com o tema selecionado
    FormatarFormulario Form, TemaAtual
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: DefinirTema
' Descrição: Define o tema visual ativo com base na seleção ou detecção automática,
'            configurando a paleta de cores correspondente para uso posterior.
'
' Parâmetros:
'   - Tema: Valor opcional da enumeração Temas que especifica o tema desejado.
'     Valores possíveis:
'       - TemaAutomatico (-1): Detecta automaticamente (padrão)
'       - TemaCinzaEscuro (3)
'       - TemaPreto (4)
'       - TemaBranco (5)
'       - TemaConfWindows (6)
'       - TemaColorido (7)
'     Se não fornecido, usa TemaAutomatico (-1)
'
' Fluxo de Execução:
'   1. Se Tema = -1 (automático), detecta tema do Office via VerificarTema()
'   2. Atribui a paleta de cores correspondente à variável TemaAtual
'   3. Caso receba valor não mapeado, usa TemaBranco como fallback
'
' Variáveis Afetadas:
'   - TemaAtual: Recebe a estrutura TemaCores correspondente ao tema selecionado
'
' Mapeamento de Temas:
'   0 -> TmColorido     3 -> TmCinzaEscuro
'   4 -> TmPreto        5 -> TmBranco
'   6 -> TmPreto        7 -> TmColorido
'   Outros -> TmBranco (fallback)
'
' Observações:
'   - Método privado chamado apenas por AplicarTema()
'   - Para temas automáticos, usa VerificarTema() que acessa registro do Windows
'   - O fallback para Windows (6) usa TmPreto por decisão de design
'   - Valores não mapeados usam TmBranco como padrão seguro
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub DefinirTema(Optional ByVal Tema As Temas = -1)
    ' Se nenhum tema for informado, tenta detectar automaticamente com base no tema do Excel
    If Tema = -1 Then Tema = VerificarTema

    ' Define qual conjunto de cores usar conforme o tipo de tema
    Select Case Tema
        Case 0: TemaAtual = TmColorido         ' Tema definido como 0 (Colorido)
        Case 3: TemaAtual = TmCinzaEscuro      ' Tema escuro com tons de cinza
        Case 4: TemaAtual = TmPreto            ' Tema com base em preto
        Case 5: TemaAtual = TmBranco           ' Tema claro com branco
        Case 6: TemaAtual = TmPreto            ' Fallback para configuração do Windows usa tema Preto
        Case 7: TemaAtual = TmColorido         ' Tema colorido padrão
        Case Else: TemaAtual = TmBranco        ' Se o tema for desconhecido, usa tema branco
    End Select
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: FormatarFormulario
' Descrição: Aplica a formatação visual completa em um UserForm e todos os seus controles,
'            utilizando a paleta de cores do tema especificado.
'
' Parâmetros:
'   - Form: Referência obrigatória ao UserForm que será formatado (MSForms.UserForm)
'   - Cores: Estrutura TemaCores contendo a paleta de cores do tema atual (ByRef para otimização)
'
' Fluxo de Operação:
'   1. Configura propriedades visuais básicas do UserForm (fundo, borda e fonte)
'   2. Itera por todos os controles contidos no formulário
'   3. Para cada controle, chama FormatarControle() passando sua Tag e o tema
'
' Propriedades Configuradas:
'   - BackColor: Cor de fundo principal (Cores.F1)
'   - BorderColor: Cor da borda (Cores.B1)
'   - Font.Name: Fonte padrão (constante FontePadrao = "Segoe UI")
'   - Font.Size: Tamanho 10 como padrão
'
' Observações:
'   - Varre recursivamente todos os controles do formulário
'   - Utiliza a propriedade .Tag de cada controle para estilos específicos
'   - Ignora erros durante formatação (On Error Resume Next implícito)
'   - Controla apenas propriedades básicas - estilos complexos são aplicados por FormatarControle()
'
' Exemplo de Chamada:
'   (Normalmente chamado apenas por AplicarTema())
'   FormatarFormulario Me, TmPreto
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub FormatarFormulario(ByRef Form As MSForms.UserForm, ByRef Cores As TemaCores)
    Dim Ctrl As MSForms.Control ' Declaração da variável de controle para iterar sobre os controles do formulário

    With Form
        .BackColor = Cores.F1         ' Aplica a cor de fundo principal do tema
        .BorderColor = Cores.B1       ' Aplica a cor da borda do formulário
        .Font.Name = FontePadrao      ' Define a fonte padrão (constante definida no módulo)
        .Font.Size = 10               ' Define o tamanho padrão da fonte

        ' Percorre todos os controles do formulário e aplica estilos individuais
        For Each Ctrl In .Controls
            FormatarControle Ctrl, Ctrl.Tag, Cores ' Aplica estilo com base na tag do controle e no tema atual
        Next Ctrl
    End With
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: FormatarControle
' Descrição: Interpreta e aplica formatação a um controle individual com base
'            na Tag especificada e no tema atual, suportando dois formatos
'            de marcação: estilos modulares e pré-definidos.
'
' Parâmetros:
'   - Ctrl: Referência ao controle MSForms a ser formatado (ByRef para otimização)
'   - Tag: String contendo as definições de estilo no formato "{tema:estilo}"
'   - Tema: Estrutura TemaCores contendo a paleta do tema atual (ByRef)
'
' Fluxo de Processamento:
'   1. Localiza o bloco de formatação correspondente ao tema atual
'   2. Identifica se o estilo é modular ou pré-definido
'   3. Extrai os parâmetros de formatação
'   4. Chama o método especializado conforme o tipo de estilo
'
' Formatos de Tag Suportados:
'   1. Estilo Pré-definido: "{tmnome:estilopredef}"
'      Ex: "{tmpr:lblbtd01}" - Usa estilo lblbtd01 do tema preto
'
'   2. Estilo Modular: "{tmnome(prop1=valor1|prop2=valor2)}"
'      Ex: "{tmpr(bc=f1|fc=l3)}" - Fundo F1 e texto L3 do tema preto
'
' Comportamentos Especiais:
'   - Se não encontrar marcação para o tema atual: sai silenciosamente (Exit Sub)
'   - Se encontrar marcação inválida: ignora o formato não reconhecido
'
' Métodos Chamados:
'   - EstilosModular: Para formatação com propriedades individuais
'   - EstilosPredefinido: Para estilos nomeados pré-configurados
'
' Exemplo de Uso Interno:
'   FormatarControle Me.Label1, "{tmpr:lblfcd01}:{tmba:lblbts01}", TemaAtual
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------

Private Sub FormatarControle(ByRef Ctrl As MSForms.Control, ByRef Tag As String, ByRef Tema As TemaCores)
    ' Declaração de variáveis para auxiliar no processamento da tag
    Dim Inicio      As Long, Fim As Long    ' Posições dos delimitadores
    Dim Bloco       As String               ' Bloco de estilo extraído
    Dim iEst        As Long                 ' Índice inicial do estilo
    Dim cEst        As Long                 ' Comprimento do estilo
    Dim Esti        As String               ' String do estilo processado
        
    ' Encontrar a posição da tag que corresponde ao nome do tema no controle
    Inicio = InStr(Tag, "{" & Tema.Nome)
    
    ' Se a tag não for encontrada, sai da sub-rotina
    If Inicio = 0 Then Exit Sub
    
    ' Localiza o final da tag (posição do fechamento de chaves)
    Fim = InStr(Inicio, Tag, "}")
    
    ' Se o fechamento de chaves não for encontrado, sai da sub-rotina
    If Fim = 0 Then Exit Sub
    
    ' Extrai o bloco de texto que representa o estilo dentro da tag
    Bloco = Mid(Tag, Inicio, Fim - Inicio + 1)
    
    ' Verifica o tipo de estilo presente na tag, distinguindo entre modular e predefinido
    Select Case Mid(Bloco, 6, 1)
        ' Se o estilo for modular (identificado por parênteses), extrai o conteúdo
        Case "("
            ' Extrai o conteúdo entre parênteses
            iEst = InStr(Bloco, "(") + 1
            cEst = InStr(Bloco, ")") - iEst
            Esti = Mid(Bloco, iEst, cEst)
            
            ' Chama a sub-rotina para aplicar o estilo modular
            EstilosModular Ctrl, Esti, Tema
            
        ' Se o estilo for predefinido (identificado por dois-pontos), extrai o conteúdo
        Case ":"
            iEst = InStr(Bloco, ":") + 1
            cEst = InStr(Bloco, "}") - iEst
            Esti = Mid(Bloco, iEst, cEst)

            ' Chama a sub-rotina para aplicar o estilo predefinido
            EstilosPredefinido Ctrl, Esti, Tema
    End Select
End Sub

'--------------------------------------------------------------
' MÉTODO: FormatarControleEmExec
' Descrição: Formata um controle específico utilizando o tema atualmente aplicado,
'            projetado para uso durante operações em tempo real como efeitos hover.
'
' Parâmetros:
'   - Ctrl: Referência ao controle MSForms a ser formatado (obrigatório)
'   - Tag: String contendo as definições de estilo no formato "{tema:estilo}" (obrigatório)
'
' Características:
'   - Interface pública para formatação pontual de controles
'   - Mantém consistência visual utilizando o TemaAtual definido
'   - Otimizado para chamadas durante interações do usuário
'
' Diferença para FormatarControle:
'   - Não recebe parâmetro de tema, utiliza sempre o TemaAtual
'   - Método público (enquanto FormatarControle é privado)
'
' Fluxo:
'   1. Recebe controle e tag de formatação
'   2. Aplica formatação utilizando o TemaAtual da classe
'   3. Delega a lógica principal para FormatarControle()
'
' Casos de Uso Típicos:
'   - Efeitos hover/active em controles
'   - Atualizações visuais dinâmicas
'   - Redefinição de estilo após modificações
'
' Exemplo:
'   ' Durante efeito hover em um Label:
'   FormatarControleEmExec Me.Label1, "{tmpr:lblbtd01}"
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Public Sub FormatarControleEmExec(ByRef Ctrl As MSForms.Control, ByRef Tag As String)
    FormatarControle Ctrl, Tag, TemaAtual
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: EstilosPredefinido
' Descrição: Aplica estilos visuais pré-definidos a um controle utilizando
'            a paleta de cores do tema especificado.
'
' Parâmetros:
'   - Ctrl: Referência ao controle MSForms a ser formatado
'   - Estilo: Nome do estilo pré-definido a ser aplicado (case-sensitive)
'   - Cores: Estrutura TemaCores contendo a paleta do tema ativo
'
' Estilos Suportados:
'   lbltsp01: Label transparente pequeno (fonte 10)
'   lbltds01: Label transparente destaque (fonte 14)
'   lblbts01: Label com borda estilo padrão
'   lblbtd01: Label com borda de destaque (cor D2)
'   lblbtd11: Label com borda de destaque (cor D1)
'   lblfcd01-05: Label com fundo colorido (destaque D1-D5)
'
' Estrutura do Array de Aparência:
'   Índice 0: BackColor     (Cor de fundo)
'   Índice 1: BackStyle     (0-Transparente, 1-Opaque)
'   Índice 2: BorderColor   (Cor da borda)
'   Índice 3: BorderStyle   (Estilo da borda)
'   Índice 4: ForeColor     (Cor do texto)
'   Índice 5: SpecialEffect (Efeitos visuais)
'   Índice 6: Font.Size     (Tamanho da fonte)
'
' Tratamento de Erros:
'   - Usa On Error Resume Next para ignorar erros em propriedades inválidas
'   - Estilo desconhecido aplica padrão (roxo com borda preta)
'   - Considera estado PlaceHolder no BackStyle
'
' Exemplo de Uso Interno:
'   EstilosPredefinido Me.Label1, "lblbtd01", TemaAtual
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub EstilosPredefinido(ByRef Ctrl As MSForms.Control, ByRef Estilo As String, ByRef Cores As TemaCores)
    ' Declaração dos arrays que armazenam os estilos e aparências
    Dim Aparencia() As Variant
    
    With Cores
        ' Seleciona o estilo a ser aplicado com base no nome do estilo
        Select Case Estilo
            ' Para cada estilo definido, carrega as propriedades correspondentes
            'Case "Nome": Aparencia() = Array(backcolor, backstyle, BorderColor, BorderStyle, ForeColor, SpecialEffect, Font.Size)
            Case "lbltsp01": Aparencia() = Array(, 0, , , .L1, 0, 10)  ' Label texto pequeno
            Case "lbltds01": Aparencia() = Array(, 0, , , .D1, 0, 14)  ' Label texto destaque
            
            Case "lblbts01": Aparencia() = Array(.F3, 1, .B2, 1, .L2, , 10)  ' Label borda padrão
            Case "lblbtd01": Aparencia() = Array(.D2, 1, .D5, 1, .L2, , 10)  ' Label borda destaque
            Case "lblbtd11": Aparencia() = Array(.D1, 1, .D5, 1, .L2, , 10)  ' Label borda destaque
            
            Case "lblfcd01": Aparencia() = Array(.D1, 1, , , .L1, 0)  ' Label fundo colorido D1
            Case "lblfcd02": Aparencia() = Array(.D2, 1, , , .L1, 0)  ' Label fundo colorido D2
            Case "lblfcd03": Aparencia() = Array(.D3, 1, , , .L1, 0)  ' Label fundo colorido D3
            Case "lblfcd04": Aparencia() = Array(.D4, 1, , , .L1, 0)  ' Label fundo colorido D4
            Case "lblfcd05": Aparencia() = Array(.D5, 1, , , .L1, 0)  ' Label fundo colorido D5
            
            Case "lblfnd01": Aparencia() = Array(.F1, 1, , , .D1, 0)  ' Label fundo colorido D1
            Case "lblfnd02": Aparencia() = Array(.F2, 1, , , .D2, 0)  ' Label fundo colorido D2
            Case "lblfnd03": Aparencia() = Array(.F3, 1, , , .D3, 0)  ' Label fundo colorido D3
            Case "lblfnd04": Aparencia() = Array(.F4, 1, , , .D4, 0)  ' Label fundo colorido D4
            Case "lblfnd05": Aparencia() = Array(.F5, 1, , , .D5, 0)  ' Label fundo colorido D5
            
            Case "lblSlc01": Aparencia() = Array(.S1, 1, , , .D1, 0)  ' Label fundo colorido D1
            Case "lblSlc02": Aparencia() = Array(.S2, 1, , , .D2, 0)  ' Label fundo colorido D2
            Case "lblSlc03": Aparencia() = Array(.S3, 1, , , .D3, 0)  ' Label fundo colorido D3
            Case "lblSlc04": Aparencia() = Array(.S4, 1, , , .D4, 0)  ' Label fundo colorido D4
            Case "lblSlc05": Aparencia() = Array(.S5, 1, , , .D5, 0)  ' Label fundo colorido D5
            
            ' Caso o estilo não esteja definido, aplica um estilo padrão
            Case Else: Aparencia() = Array(&H1000F0, 1, &H0, 1, , 0, 10) ' Estilo padrão (roxo)
        End Select
    End With
    
    ' Ignora erros temporariamente durante a aplicação das propriedades
    On Error Resume Next
    With Ctrl
        ' Aplica as propriedades extraídas para o controle
        .BackColor = Aparencia(0)      ' Cor de fundo
        .BackStyle = Aparencia(1)      ' Estilo de fundo (transparente ou sólido)
        .BorderColor = Aparencia(2)    ' Cor da borda
        .BorderStyle = Aparencia(3)    ' Estilo da borda (reta, arredondada, etc.)
        .ForeColor = Aparencia(4)      ' Cor da fonte
        .SpecialEffect = Aparencia(5)  ' Efeito especial (como 3D)
        .Font.Size = Aparencia(6)      ' Tamanho da fonte
        .Font.Name = FontePadrao       ' Fonte padrão definida
        If PlaceHolder Then .BackStyle = 0  ' Override para placeholders
    End With
    ' Restaura a configuração de erro padrão após a aplicação
    On Error GoTo 0
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: EstilosModular
' Descrição: Aplica estilos visuais personalizados a um controle através de uma
'            string de formatação modular, permitindo customização avançada.
'
' Parâmetros:
'   - Ctrl: Referência ao controle MSForms a ser formatado (ByRef)
'   - EstiloTexto: String contendo as definições de estilo no formato "prop=valor|prop=valor"
'   - Cores: Estrutura TemaCores contendo a paleta do tema ativo (ByRef)
'
' Formato da String de Estilo:
'   "BC=F1|FC=L3|BS=1" onde:
'   - BC: BackColor (cor de fundo)
'   - FC: ForeColor (cor do texto)
'   - BS: BackStyle (estilo de fundo)
'   Valores podem ser:
'     - Cores temáticas (F1, S2, D3, etc.)
'     - Cores hexadecimais (#FF0000 ou &H0000FF)
'     - Valores numéricos para propriedades específicas
'
' Fluxo de Processamento:
'   1. Divide a string de estilo nos separadores "|"
'   2. Para cada definição, separa propriedade e valor no caractere "="
'   3. Chama InterpretarEstilo para aplicar cada propriedade individual
'   4. Garante fonte padrão ao final
'
' Tratamento Especial:
'   - Considera o estado PlaceHolder (define BackStyle=0 quando ativo)
'   - Remove espaços em branco com Trim()
'   - Converte siglas para maiúsculas (UCase)
'
' Exemplo de Uso:
'   EstilosModular Me.Label1, "BC=D2|FC=L1|BS=1", TemaAtual
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub EstilosModular(ByRef Ctrl As MSForms.Control, ByRef EstiloTexto As String, ByRef Cores As TemaCores)
    ' Declaração de variáveis
    Dim Estilos() As String ' Array para armazenar os fragmentos de estilo
    Dim Item As Variant     ' Variável para iterar sobre os estilos
    Dim Sigla As String     ' Sigla da propriedade (ex: BC=BackColor)
    Dim Valor As String     ' Valor da propriedade correspondente
    
    ' Divide o texto de estilo nos separadores "|"
    Estilos = Split(EstiloTexto, "|")
    
    ' Processa cada definição de estilo individualmente
    For Each Item In Estilos
        ' Extrai a propriedade e valor, removendo espaços e padronizando
        Sigla = UCase(Trim(Split(Item, "=")(0))) ' Padroniza para maiúsculas
        Valor = Trim(Split(Item, "=")(1))         ' Remove espaços extras
        
        ' Aplica o estilo através do método auxiliar
        InterpretarEstilo Sigla, Valor, Cores, Ctrl
        
        ' Override para controles com placeholder ativo
        If PlaceHolder Then Ctrl.BackStyle = 0
    Next Item
    
    ' Garante que não aconteça um erro se o controle não poder receber uma fonte
    On Error Resume Next
    ' Garante o uso da fonte padrão em todos os casos
    Ctrl.Font.Name = FontePadrao
    On Error GoTo 0
End Sub

'--------------------------------------------------------------
' MÉTODO PRIVADO: InterpretarEstilo
' Descrição: Interpreta e aplica uma propriedade de estilo individual a um controle,
'            suportando múltiplos formatos de valores (hexadecimal, numérico e temático).
'
' Parâmetros:
'   - Sigla: Código da propriedade a ser aplicada (BC, FC, BD, BS, BY, FS, SE)
'   - Valor: Valor da propriedade em diversos formatos possíveis
'   - Cores: Paleta de cores do tema ativo (por referência para eficiência)
'   - Ctrl: Controle MSForms que receberá a formatação (por referência)
'
' Formatos de Valor Suportados:
'   1. Hexadecimal: "#RRGGBB" ou "&HBBGGRR"
'      Ex: "BC=#FF0000" (fundo vermelho) ou "FC=&H0000FF" (texto azul)
'
'   2. Numérico Direto: "0" a "5" para níveis ou valores específicos
'      Ex: "BS=1" (fundo opaco), "FS=12" (fonte tamanho 12)
'
'   3. Referência Temática: "LetraNúmero" (F1, S2, D3, etc.)
'      Ex: "BC=F1" (fundo usando cor F1 do tema)
'
' Tabela de Siglas:
'   BC: BackColor       (Cor de fundo)
'   FC: ForeColor       (Cor do texto)
'   BD: BorderColor     (Cor da borda)
'   BS: BackStyle       (0-Transparente, 1-Opaque)
'   BY: BorderStyle     (Estilo da borda)
'   FS: Font.Size       (Tamanho da fonte)
'   SE: SpecialEffect   (Efeitos 3D)
'
' Fluxo de Decisão:
'   1. Verifica se valor é hexadecimal (começa com # ou &H)
'   2. Se não, verifica se é valor numérico puro
'   3. Se não, interpreta como referência temática (Letra+Nível)
'
' Tratamento de Erros:
'   - On Error Resume Next para propriedades inválidas
'   - Conversão segura para Long nos valores numéricos
'   - CLng seguro para valores temáticos (ex: "F2" -> nível 2)
'
' Exemplo de Chamada Interna:
'   InterpretarEstilo "BC", "D3", TemaAtual, Me.Label1
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub InterpretarEstilo(ByVal Sigla As String, ByVal Valor As String, ByRef Cores As TemaCores, ByRef Ctrl As MSForms.Control)
    ' Declaração de variáveis
    Dim Letra As String  ' Primeira letra para referências temáticas
    Dim Nivel As Long    ' Nível da cor temática (1-5)
    
    ' Configura tratamento de erros para propriedades inválidas
    On Error Resume Next

    ' Caso 1: Valor hexadecimal (#RRGGBB ou &HBBGGRR)
    If Left(Valor, 1) = "#" Or Left(Valor, 2) = "&H" Then
        Select Case Sigla
            Case "BC": Ctrl.BackColor = ConverterParaCor(Valor)
            Case "FC": Ctrl.ForeColor = ConverterParaCor(Valor)
            Case "BD": Ctrl.BorderColor = ConverterParaCor(Valor)
        End Select

    ' Caso 2: Valor numérico direto
    ElseIf IsNumeric(Valor) Then
        Nivel = CLng(Valor)
        Select Case Sigla
            Case "BC": Ctrl.BackColor = GetCorPorSiglaNivel(Cores, Sigla, Nivel)
            Case "BS": Ctrl.BackStyle = Nivel
            Case "FC": Ctrl.ForeColor = GetCorPorSiglaNivel(Cores, Sigla, Nivel)
            Case "BD": Ctrl.BorderColor = GetCorPorSiglaNivel(Cores, Sigla, Nivel)
            Case "BY": Ctrl.BorderStyle = Nivel
            Case "FS": Ctrl.Font.Size = Nivel
            Case "SE": Ctrl.SpecialEffect = Nivel
        End Select

    ' Caso 3: Referência temática (Letra+Nível)
    Else
        Letra = Left(Valor, 1)
        Nivel = CLng(Mid(Valor, 2))
        Select Case Sigla
            Case "BC": Ctrl.BackColor = GetCorPorLetraNivel(Cores, Letra, Nivel)
            Case "FC": Ctrl.ForeColor = GetCorPorLetraNivel(Cores, Letra, Nivel)
            Case "BD": Ctrl.BorderColor = GetCorPorLetraNivel(Cores, Letra, Nivel)
        End Select
    End If

    ' Restaura o tratamento de erros padrão
    On Error GoTo 0
End Sub

'--------------------------------------------------------------
' FUNÇÃO PRIVADA: GetCorPorSiglaNivel
' Descrição: Obtém uma cor específica da paleta do tema com base no tipo de propriedade
'            (fundo, texto ou borda) e no nível de intensidade desejado.
'
' Parâmetros:
'   - Cores: Estrutura TemaCores contendo a paleta completa (passada por referência)
'   - Sigla: Tipo de propriedade ("BC", "FC" ou "BD")
'   - Nivel: Intensidade da cor (1 a 5, sendo 1 mais claro/suave e 5 mais intenso)
'
' Valores de Retorno:
'   - Valor Long representando a cor RGB solicitada
'   - Vermelho (RGB(255, 0, 0)) para parâmetros inválidos
'
' Tabela de Siglas:
'   BC: BackColor (cor de fundo) - Usa a escala F1-F5 do tema
'   FC: ForeColor (cor do texto) - Usa a escala L1-L5 do tema
'   BD: BorderColor (cor da borda) - Usa a escala B1-B5 do tema
'
' Níveis de Cor:
'   1: Cor mais clara/suave da escala
'   3: Cor média da escala
'   5: Cor mais intensa da escala
'
' Tratamento de Erros:
'   - Retorna vermelho para níveis fora do intervalo 1-5
'   - Retorna vermelho para siglas não reconhecidas
'   - Conversão case-insensitive das siglas
'
' Exemplo de Uso:
'   CorFundo = GetCorPorSiglaNivel(TemaAtual, "BC", 2) ' Obtém F2 do tema
'   CorTexto = GetCorPorSiglaNivel(TemaAtual, "FC", 3) ' Obtém L3 do tema
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Function GetCorPorSiglaNivel(ByRef Cores As TemaCores, ByVal Sigla As String, Optional ByVal Nivel As Long = 1) As Long
    ' Declaração de variável para armazenar as cores do tipo especificado (fundo, fonte, borda)
    Dim CoresArray() As Variant

    ' Valida se o nível está dentro do intervalo aceitável (1 a 5)
    If Nivel < 1 Or Nivel > 5 Then
        GetCorPorSiglaNivel = RGB(255, 0, 0) ' Retorna a cor de erro (vermelho) se o nível for inválido
        Exit Function
    End If

    ' Seleciona o conjunto de cores correspondente à sigla fornecida
    Select Case UCase(Sigla) ' UCase garante que a sigla seja tratada de forma case-insensitive
        Case "BC" ' Caso a sigla seja "BC", refere-se à cor de fundo
            CoresArray = Array(Cores.F1, Cores.F2, Cores.F3, Cores.F4, Cores.F5)
        Case "FC" ' Caso a sigla seja "FC", refere-se à cor da fonte
            CoresArray = Array(Cores.L1, Cores.L2, Cores.L3, Cores.L4, Cores.L5)
        Case "BD" ' Caso a sigla seja "BD", refere-se à cor da borda
            CoresArray = Array(Cores.B1, Cores.B2, Cores.B3, Cores.B4, Cores.B5)
        Case Else ' Caso a sigla não seja válida, retorna a cor de erro (vermelho)
            GetCorPorSiglaNivel = RGB(255, 0, 0)
            Exit Function
    End Select

    ' Retorna a cor correspondente ao nível especificado (nível é de 1 a 5, então ajusta o índice para 0 a 4)
    GetCorPorSiglaNivel = CoresArray(Nivel - 1)
End Function

'--------------------------------------------------------------
' FUNÇÃO PRIVADA: GetCorPorLetraNivel
' Descrição: Obtém uma cor específica da paleta de temas com base na categoria (F, S, D, L, B)
'            e nível de intensidade (1-5), seguindo o padrão de design do sistema.
'
' Parâmetros:
'   - Cores: Estrutura TemaCores contendo todas as paletas (passada por referência)
'   - Letra: Categoria de cor (F-Fundo, S-Seleção, D-Destaque, L-Letra, B-Borda)
'   - Nivel: Intensidade da cor (1 mais claro/suave a 5 mais intenso) - Opcional (padrão=1)
'
' Tabela de Categorias:
'   F: Cores de Fundo (F1 a F5) - Cores principais de background
'   S: Cores de Seleção (S1 a S5) - Cores para estados selecionados/hover
'   D: Cores de Destaque (D1 a D5) - Cores acentuadas para elementos importantes
'   L: Cores de Letra (L1 a L5) - Cores para texto e tipografia
'   B: Cores de Borda (B1 a B5) - Cores para bordas e divisores
'
' Níveis de Intensidade:
'   1: Tom mais claro/suave da escala
'   3: Tom médio da escala
'   5: Tom mais escuro/intenso da escala
'
' Valores de Retorno:
'   - Valor Long contendo a cor RGB solicitada
'   - Vermelho (RGB(255,0,0)) para parâmetros inválidos
'
' Tratamento de Erros:
'   - Retorna vermelho para níveis fora da faixa 1-5
'   - Retorna vermelho para categorias não reconhecidas
'   - Case-insensitive (aceita 'f' ou 'F')
'
' Exemplo de Uso:
'   corFundo = GetCorPorLetraNivel(TemaAtual, "F", 2)  ' Obtém F2 (segundo tom de fundo)
'   corTexto = GetCorPorLetraNivel(TemaAtual, "L", 1)  ' Obtém L1 (texto mais claro)
'   corDestaque = GetCorPorLetraNivel(TemaAtual, "D", 4) ' Obtém D4 (destaque intenso)
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Function GetCorPorLetraNivel(ByRef Cores As TemaCores, ByVal Letra As String, Optional ByVal Nivel As Long = 1) As Long
    ' Declaração de variável para armazenar as cores do tipo especificado
    Dim CoresArray() As Variant

    ' Valida se o nível está dentro do intervalo aceitável (1 a 5)
    If Nivel < 1 Or Nivel > 5 Then
        GetCorPorLetraNivel = RGB(255, 0, 0) ' Retorna vermelho para nível inválido
        Exit Function
    End If

    ' Seleciona o conjunto de cores correspondente à letra fornecida
    Select Case UCase(Letra) ' Converte para maiúsculas para comparação
        Case "F": CoresArray = Array(Cores.F1, Cores.F2, Cores.F3, Cores.F4, Cores.F5)
        Case "S": CoresArray = Array(Cores.S1, Cores.S2, Cores.S3, Cores.S4, Cores.S5)
        Case "D": CoresArray = Array(Cores.D1, Cores.D2, Cores.D3, Cores.D4, Cores.D5)
        Case "L": CoresArray = Array(Cores.L1, Cores.L2, Cores.L3, Cores.L4, Cores.L5)
        Case "B": CoresArray = Array(Cores.B1, Cores.B2, Cores.B3, Cores.B4, Cores.B5)
        Case Else:
            GetCorPorLetraNivel = RGB(255, 0, 0) ' Retorna vermelho para categoria inválida
            Exit Function
    End Select

    ' Retorna a cor correspondente ao nível (ajustando índice 1-5 para 0-4)
    GetCorPorLetraNivel = CoresArray(Nivel - 1)
End Function

'--------------------------------------------------------------
' FUNÇÃO PRIVADA: ConverterParaCor
' Descrição: Converte uma representação de cor em string para o formato Long
'            utilizado pelo VBA, suportando múltiplos formatos de entrada.
'
' Parâmetros:
'   - Valor: String contendo a cor em um dos formatos suportados
'
' Formatos Suportados:
'   1. Valor numérico direto: "16711680" ou "&HFF0000" (formato BGR do VBA)
'   2. Hexadecimal com #: "#RRGGBB" (ex: "#FF0000" para vermelho)
'   3. Hexadecimal VBA: "&HBBGGRR" (ex: "&H0000FF" para vermelho)
'   4. Hexadecimal sem prefixo: "RRGGBB" (ex: "FF0000" para vermelho)
'
' Valores de Retorno:
'   - Valor Long representando a cor no formato RGB do VBA
'   - Vermelho (RGB(255,0,0)) para formatos inválidos
'
' Fluxo de Conversão:
'   1. Remove espaços e verifica formato numérico direto
'   2. Processa formato #RRGGBB (converte componentes R, G, B)
'   3. Processa formato &HBBGGRR (conversão direta VBA)
'   4. Processa formato RRGGBB sem prefixo
'   5. Retorna vermelho para qualquer formato não reconhecido
'
' Tratamento de Erros:
'   - On Error Resume Next para conversões hexadecimais inválidas
'   - Retorna vermelho para qualquer entrada não reconhecida
'   - Remove espaços extras automaticamente
'
' Exemplos de Uso:
'   cor1 = ConverterParaCor("#FF0000")   ' Vermelho puro
'   cor2 = ConverterParaCor("&H0000FF")  ' Vermelho (formato VBA)
'   cor3 = ConverterParaCor("00FF00")    ' Verde puro
'   cor4 = ConverterParaCor("255")       ' Valor numérico direto
'
' Observações:
'   - Formato VBA (&H) usa ordem inversa de componentes (BBGGRR)
'   - Formato # e sem prefixo usam ordem normal (RRGGBB)
'   - A função não suporta valores com transparência (alpha)
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Function ConverterParaCor(ByVal Valor As String) As Long
    Dim R As Long, G As Long, B As Long
    Dim HexValor As String

    Valor = Trim(Valor) ' Remove espaços em branco ao redor do valor

    ' Caso 1: Valor numérico direto (decimal ou hexadecimal VBA)
    If IsNumeric(Valor) Then
        ConverterParaCor = CLng(Valor)
        Exit Function

    ' Caso 2: Formato #RRGGBB (hexadecimal com #)
    ElseIf Left(Valor, 1) = "#" Then
        HexValor = Replace(Valor, "#", "")
        If Len(HexValor) = 6 Then
            R = CLng("&H" & Mid(HexValor, 1, 2)) ' Componente Vermelho
            G = CLng("&H" & Mid(HexValor, 3, 2)) ' Componente Verde
            B = CLng("&H" & Mid(HexValor, 5, 2)) ' Componente Azul
            ConverterParaCor = RGB(R, G, B)
            Exit Function
        End If

    ' Caso 3: Formato &HBBGGRR (hexadecimal VBA)
    ElseIf UCase(Left(Valor, 2)) = "&H" Then
        On Error Resume Next
        ConverterParaCor = CLng(Valor)
        If Err.Number <> 0 Then ConverterParaCor = RGB(255, 0, 0)
        On Error GoTo 0
        Exit Function

    ' Caso 4: Formato RRGGBB sem prefixo
    ElseIf Len(Valor) = 6 And IsNumeric("&H" & Valor) Then
        R = CLng("&H" & Mid(Valor, 1, 2))
        G = CLng("&H" & Mid(Valor, 3, 2))
        B = CLng("&H" & Mid(Valor, 5, 2))
        ConverterParaCor = RGB(R, G, B)
        Exit Function
    End If

    ' Caso de erro: formato não reconhecido
    ConverterParaCor = RGB(255, 0, 0)
End Function

'--------------------------------------------------------------
' FUNÇÃO: VerificarTema
' Descrição: Detecta o tema de interface do usuário do Microsoft Office
'            consultando o registro do Windows.
'
' Retorno:
'   - Retorna um valor inteiro representando o tema do Office:
'     0 = Tema Colorido (Office 2019+)
'     3 = Tema Cinza Escuro
'     4 = Tema Preto
'     5 = Tema Branco
'     6 = Usar configuração do Windows
'     7 = Tema Colorido (Office 2013-2016)
'     -1 = Não foi possível detectar o tema
'
' Fluxo de Operação:
'   1. Tenta ler a chave do registro para Office 2016+ (versão 16.0)
'   2. Se falhar, tenta a chave para Office 2013 (versão 15.0)
'   3. Se ambas falharem, retorna -1 (indetectável)
'
' Requisitos:
'   - Acesso ao registro do Windows
'   - Permissões de leitura na chave do Office
'   - Execução em plataforma Windows
'
' Limitações:
'   - Funciona apenas no Windows
'   - Depende da estrutura de registro do Office
'   - Pode não funcionar em versões futuras do Office
'
' Exemplo de Uso:
'   temaAtual = VerificarTema()
'   Select Case temaAtual
'       Case 4: AplicarTema Me, TemaPreto
'       Case 5: AplicarTema Me, TemaBranco
'       Case Else: AplicarTema Me, TemaAutomatico
'   End Select
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Public Function VerificarTema() As Integer
    ' Declaração das variáveis
    Dim objShell    As Object    ' Objeto WScript.Shell para acesso ao registro
    Dim Tema        As Variant   ' Valor lido do registro
    Dim ChaveReg    As String    ' Caminho completo da chave de registro
    Dim VersOffc    As String    ' Versão do Office a verificar

    On Error Resume Next  ' Habilita continuidade em caso de erros

    ' Primeira tentativa: Office 2016 ou superior (versão 16.0+)
    VersOffc = "16.0"
    ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & VersOffc & "\Common\UI Theme"
    
    Set objShell = CreateObject("WScript.Shell")
    VerificarTema = objShell.RegRead(ChaveReg)
    
    ' Segunda tentativa (se falhar): Office 2013 (versão 15.0)
    If Err.Number <> 0 Then
        Err.Clear  ' Limpa o erro anterior
        VersOffc = "15.0"
        ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & VersOffc & "\Common\UI Theme"
        VerificarTema = objShell.RegRead(ChaveReg)
        
        ' Se ainda falhar, retorna -1 (tema indetectável)
        If Err.Number <> 0 Then
            VerificarTema = -1
        End If
    End If

    ' Nota: Não é necessário liberar objShell pois é liberado automaticamente
    ' quando a função termina (escopo local)
End Function

'--------------------------------------------------------------
' MÉTODO: EfeitoHover
' Descrição: Configura os eventos de mouse para um controle específico,
'            permitindo efeitos visuais dinâmicos quando o usuário interage,
'            com suporte a aplicação de temas específicos.
'
' Parâmetros:
'   - Form: Referência ao UserForm pai (obrigatório, MSForms.UserForm)
'   - Ctrl: Controle que receberá o efeito hover (obrigatório, MSForms.Control)
'   - Tag: String de formatação com estilos para diferentes estados (String)
'   - Tema: Valor da enumeração Temas que define o tema a ser aplicado (Temas)
'   - Fundo: [Opcional] Controle de fundo para resetar o hover (Label ou Frame)
'
' Controles Suportados:
'   [Básicos]
'   - Label, CommandButton, CheckBox, OptionButton, ToggleButton
'   [Controles de Entrada]
'   - TextBox, ComboBox, ListBox
'   [Contêineres]
'   - Frame, MultiPage, TabStrip
'   [Outros]
'   - Image
'
' Fluxo de Operação:
'   1. Configura o controle de fundo para eventos MouseMove
'   2. Associa o controle principal aos manipuladores de eventos apropriados
'   3. Armazena a Tag de formatação na variável TagEmExec
'   4. Aplica o tema especificado através do parâmetro Tema
'   5. Automaticamente remove o efeito quando o mouse sai da área
'
' Estrutura da Tag:
'   Formato: "{tema1:estilo1}:{tema2:estilo2}:{tema3:estilo3}"
'   Onde:
'   - temaX: Prefixo do tema (tmpr, tmbr, tmci, tmco)
'   - estiloX: Pode ser:
'     * Nome de estilo pré-definido (ex: "lblbtd01")
'     * Estilo modular (ex: "(bc=f1|fc=l1|bs=1)")
'
' Tratamento de Erros:
'   - Exibe mensagem para controles não suportados
'   - Ignora erros durante aplicação do tema (On Error Resume Next)
'   - Mantém funcionamento mesmo se fundo não for especificado
'
' Observações Importantes:
'   1. Variáveis WithEvents necessárias:
'      - Para cada tipo de controle (declaradas no módulo da classe)
'      - Ex: Private WithEvents CtrlLabel As MSForms.Label
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Public Sub EfeitoHover(ByRef Form As MSForms.UserForm, ByRef Ctrl As MSForms.Control, ByRef Tag As String, ByVal Tema As Temas, Optional ByRef Fundo As MSForms.Control)
    ' Configura o controle de fundo para eventos MouseMove
    Select Case TypeName(Fundo)
        Case "Label": Set FndLabel = Fundo
        Case "Frame": Set FndFrame = Fundo
        Case Else: Set Formulario = Form
    End Select
    
    ' Associa o controle principal ao manipulador de eventos apropriado
    Select Case TypeName(Ctrl)
        Case "Label":         Set CtrlLabel = Ctrl
        Case "CommandButton": Set CtrlCommandBtn = Ctrl
        Case "CheckBox":      Set CtrlCheckBox = Ctrl
        Case "OptionButton":  Set CtrlOptionBtn = Ctrl
        Case "ToggleButton":  Set CtrlToggleBtn = Ctrl
        Case "TextBox":       Set CtrlTextBox = Ctrl
        Case "ComboBox":      Set CtrlComboBox = Ctrl
        Case "ListBox":       Set CtrlListBox = Ctrl
        Case "Frame":         Set CtrlFrame = Ctrl
        Case "Image":         Set CtrlImage = Ctrl
        Case "MultiPage":     Set CtrlMultiPage = Ctrl
        Case "TabStrip":      Set CtrlTabStrip = Ctrl
        Case Else
            MsgBox "Tipo de controle não suportado: " & TypeName(Ctrl)
    End Select
    
    ' Armazena a tag de formatação e aplica o tema
    TagEmExec = Tag
    DefinirTema Tema
End Sub

'--------------------------------------------------------------
' MÉTODO: EfeitoPlaceHolder
' Descrição: Cria um efeito de placeholder (texto de dica) para controles de entrada,
'            mostrando um texto indicativo quando o controle está vazio.
'
' Parâmetros:
'   - Form: Referência ao UserForm que contém o controle (obrigatório)
'   - Ctrl: Controle que receberá o placeholder (TextBox ou ComboBox)
'   - Texto: Texto do placeholder a ser exibido (obrigatório)
'
' Controles Suportados:
'   - TextBox (caixa de texto)
'   - ComboBox (caixa de combinação)
'
' Comportamento:
'   1. Verifica se o controle é do tipo suportado
'   2. Cria um controle de placeholder sobreposto (via CriandoPH)
'   3. Configura o controle principal como transparente inicialmente
'   4. Configura o handler para atualização dinâmica (AtualizarPH)
'
' Fluxo de Operação:
'   - O placeholder é exibido quando o controle está vazio
'   - É automaticamente ocultado quando o usuário digita algo
'   - Reaparece se o usuário limpar o conteúdo
'
' Tratamento de Erros:
'   - Exibe mensagem para controles não suportados e sai do método
'   - Ignora erros durante a criação do placeholder
'
' Exemplo de Uso:
'   EfeitoPlaceHolder Me, Me.txtNome, "Digite seu nome completo"
'   EfeitoPlaceHolder Me, Me.cmbEstado, "Selecione um estado"
'
' Observações:
'   - Utiliza um TextBox transparente sobreposto para simular o placeholder
'   - Requer os métodos auxiliares CriandoPH e AtualizarPH
'   - O controle original deve ter BackStyle = 1 (opaco) para funcionar corretamente
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Public Sub EfeitoPlaceHolder(ByRef Form As MSForms.UserForm, ByRef Ctrl As MSForms.Control, ByRef Texto As String, Optional CorLetra As Long = 0)
    ' Verifica o tipo de controle e configura os handlers de eventos
    Select Case TypeName(Ctrl)
        Case "ComboBox":      Set CtrlComboBox = Ctrl
        Case "TextBox":       Set CtrlTextBox = Ctrl
        Case Else
            MsgBox "Tipo de controle não suporta Efeito Place Holder: " & TypeName(Ctrl)
            Exit Sub
    End Select
    
    
    ' Cria o controle de placeholder e configura o estado inicial
    If CorLetra = 0 Then CriandoPH Form, Ctrl, Texto Else CriandoPH Form, Ctrl, Texto, CorLetra
    Ctrl.BackStyle = 0  ' Torna o controle transparente
    AtualizarPH Ctrl    ' Configura o estado inicial do placeholder
End Sub
'--------------------------------------------------------------
' CONJUNTO DE EVENTOS: Manipulação de Hover e PlaceHolder
' Descrição: Implementa efeitos visuais dinâmicos (hover e placeholder) para controles
'            em UserForms do VBA, melhorando a experiência do usuário.
'
' Estrutura:
'   1. Eventos MouseMove para controles específicos:
'      - Ativam efeito hover quando o mouse entra no controle
'      - Chamam AtivarHover passando o controle específico
'
'   2. Eventos MouseMove para áreas neutras (formulário, fundos):
'      - Desativam qualquer hover ativo
'      - Chamam DesativarHover
'
'   3. Eventos Change para controles editáveis:
'      - Gerenciam dinamicamente o efeito placeholder
'      - Chamam AtualizarPH quando o conteúdo é alterado
'
' Controles Suportados:
'   - CheckBox, ComboBox, CommandButton, Frame, Image, Label
'   - ListBox, MultiPage, OptionButton, TabStrip, TextBox, ToggleButton
'
' Fluxo de Operação:
'   Hover:
'   1. Mouse entra no controle -> AtivarHover
'   2. Mouse sai do controle -> DesativarHover
'   3. Mouse move em área neutra -> DesativarHover
'
'   PlaceHolder (apenas ComboBox e TextBox):
'   1. Conteúdo alterado -> AtualizarPH
'   2. Se vazio: mostra placeholder
'   3. Se preenchido: oculta placeholder
'
' Dependências:
'   - Métodos AtivarHover, DesativarHover, AtualizarPH
'   - Variáveis globais: UltContro, CtrlPlaceHolder, PlaceHolder
'
' Padrão de Nomenclatura:
'   Ctrl[TipoControle]_MouseMove -> Para eventos de movimento do mouse
'   Ctrl[TipoControle]_Change -> Para eventos de alteração (editáveis)
'   Formulario_MouseMove -> Para área principal do formulário
'   Fnd[TipoControle]_MouseMove -> Para controles de fundo
'
' Observações:
'   1. Consistência:
'      - Todos os controles interativos implementam hover
'      - Apenas controles editáveis implementam placeholder
'
'   2. Performance:
'      - Os eventos são acionados frequentemente (MouseMove)
'      - Código deve ser otimizado para execução rápida
'
'   3. Manutenção:
'      - Padrão facilita adição de novos controles
'      - Comportamento uniforme em toda a aplicação
'
' Exemplo de Fluxo Completo:
'   1. Usuário move mouse para TextBox -> Ativa hover
'   2. Usuário clica e digita -> Placeholder some
'   3. Usuário remove texto -> Placeholder reaparece
'   4. Usuário move mouse para área vazia -> Hover desativado
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub CtrlCheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlCheckBox
End Sub

Private Sub CtrlComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlComboBox
End Sub
Private Sub CtrlComboBox_Change()
    AtualizarPH CtrlComboBox
End Sub

Private Sub CtrlCommandBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlCommandBtn
End Sub

Private Sub CtrlFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlFrame
End Sub

Private Sub CtrlImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlImage
End Sub

Private Sub CtrlLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlLabel
End Sub

Private Sub CtrlListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlListBox
End Sub

Private Sub CtrlMultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlMultiPage
End Sub

Private Sub CtrlOptionBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlOptionBtn
End Sub

Private Sub CtrlTabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlTabStrip
End Sub

Private Sub CtrlTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlTextBox
End Sub
Private Sub CtrlTextBox_Change()
    AtualizarPH CtrlTextBox
End Sub

Private Sub CtrlToggleBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AtivarHover CtrlToggleBtn
End Sub

Private Sub Formulario_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DesativarHover
End Sub

Private Sub FndLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DesativarHover
End Sub

Private Sub FndFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DesativarHover
End Sub

'--------------------------------------------------------------
' MÉTODO: AtivarHover
' Descrição: Ativa o efeito visual de hover (sobreposição) para um controle específico,
'            aplicando formatação especial quando o mouse está sobre o controle.
'
' Parâmetros:
'   - Ctrol: Controle que receberá o efeito hover (obrigatório)
'
' Comportamento:
'   1. Aplica formatação especial definida em TagEmExec ao controle
'   2. Armazena referência do controle para posterior restauração
'
' Fluxo de Operação:
'   - Chamado quando o mouse entra na área do controle
'   - Mantém controle formatado até chamada de DesativarHover
'
' Exemplo de Uso:
'   AtivarHover Me.cmdBotao
'
' Observações:
'   - Requer a variável global UltContro para armazenar estado
'   - Utiliza o método FormatarControleEmExec para aplicar formatação
'   - Pareado com o método DesativarHover
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub AtivarHover(Ctrol As MSForms.Control)
    FormatarControleEmExec Ctrol, TagEmExec
    Set UltContro = Ctrol
End Sub

'--------------------------------------------------------------
' MÉTODO: DesativarHover
' Descrição: Remove o efeito visual de hover, restaurando a formatação original
'            do controle que estava com o efeito ativo.
'
' Parâmetros: N/A
'
' Comportamento:
'   1. Verifica se há um controle com hover ativo
'   2. Restaura formatação original usando a Tag do controle
'   3. Limpa a referência armazenada
'
' Fluxo de Operação:
'   - Chamado quando o mouse sai da área do controle
'   - Restaura estado anterior do controle
'
' Exemplo de Uso:
'   DesativarHover
'
' Observações:
'   - Funciona em conjunto com AtivarHover
'   - Seguro para chamadas múltiplas (verifica estado)
'   - Utiliza a Tag do controle para armazenar formatação original
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub DesativarHover()
    If Not UltContro Is Nothing Then
        FormatarControleEmExec UltContro, UltContro.Tag
        Set UltContro = Nothing
    End If
End Sub

'--------------------------------------------------------------
' MÉTODO: CriandoPH
' Descrição: Cria um controle de placeholder sobreposto ao controle principal,
'            simulando o efeito de texto de dica (placeholder).
'
' Parâmetros:
'   - Form: UserForm que contém o controle (obrigatório)
'   - Ctrl: Controle que receberá o placeholder (obrigatório)
'   - Texto: Texto do placeholder a ser exibido (obrigatório)
'
' Comportamento:
'   1. Cria um TextBox transparente sobreposto ao controle principal
'   2. Copia propriedades visuais do controle original
'   3. Configura ZOrder para garantir sobreposição correta
'
' Fluxo de Operação:
'   - Chamado uma vez durante inicialização do efeito placeholder
'   - Cria controle que será gerenciado por AtualizarPH
'
' Exemplo de Uso:
'   CriandoPH Me, Me.txtUsuario, "Digite seu usuário"
'
' Observações:
'   - Utiliza a variável global CtrlPlaceHolder
'   - O controle criado é um TextBox padrão do MSForms
'   - Requer chamada subsequente de AtualizarPH
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub CriandoPH(ByRef Form As MSForms.UserForm, ByRef Ctrl As MSForms.Control, ByRef Texto As String, Optional CorLetra As Long)
    Dim Main    As Object
    Set Main = Ctrl.Parent
    Set CtrlPlaceHolder = Main.Controls.Add("Forms.TextBox.1")
    
    With CtrlPlaceHolder
        .Text = Texto
        If CorLetra = 0 Then .ForeColor = Ctrl.ForeColor Else .ForeColor = CorLetra
        .Font.Name = Ctrl.Font.Name
        .Font.Size = Ctrl.Font.Size
        .Left = Ctrl.Left
        .Top = Ctrl.Top
        .Width = Ctrl.Width
        .Height = Ctrl.Height
        .BackColor = Ctrl.BackColor
        .BorderStyle = Ctrl.BorderStyle
        .BorderColor = Ctrl.BorderColor
        .TabStop = False
        .ZOrder 0
    End With
    Ctrl.ZOrder 0
End Sub

'--------------------------------------------------------------
' MÉTODO: AtualizarPH
' Descrição: Gerencia a visibilidade do efeito placeholder com base no estado
'            atual do controle (vazio ou com conteúdo).
'
' Parâmetros:
'   - Ctrl: Controle associado ao placeholder (obrigatório)
'
' Comportamento:
'   1. Verifica se o controle está vazio
'   2. Ajusta BackStyle para mostrar/ocultar placeholder
'   3. Atualiza variável de estado PlaceHolder
'
' Fluxo de Operação:
'   - Chamado durante eventos de mudança no controle
'   - Responde dinamicamente à entrada do usuário
'
' Exemplo de Uso:
'   AtualizarPH Me.txtSenha
'
' Observações:
'   - Trabalha em conjunto com CriandoPH
'   - Utiliza a variável global PlaceHolder para controle de estado
'   - BackStyle 0 = transparente (mostra placeholder)
'   - BackStyle 1 = opaco (esconde placeholder)
'
' Autor: Adriano Furtado Lima
'--------------------------------------------------------------
Private Sub AtualizarPH(ByRef Ctrl As MSForms.Control)
    Select Case Ctrl.Text
        Case ""
            Ctrl.BackStyle = 0
            PlaceHolder = True
        Case Is <> ""
            Ctrl.BackStyle = 1
            PlaceHolder = False
    End Select
End Sub
