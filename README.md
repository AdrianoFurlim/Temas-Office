`clsTemasOffice` é uma classe robusta para Visual Basic for Applications (VBA) projetada para modernizar a aparência dos seus UserForms do Microsoft Office. Ela oferece um sistema completo para aplicar temas visuais personalizados, gerenciar efeitos de _hover_ em controles e adicionar _placeholders_ em caixas de texto de maneira simples e eficiente.

## Características

- **Temas Visuais Personalizados**: Aplique diferentes temas (Preto, Branco, Cinza Escuro, Colorido, ou automático) a um UserForm inteiro.
- **Efeitos de Hover e Placeholders**: Adicione efeitos de destaque ao passar o mouse sobre os controles e inclua textos de dica (_placeholders_) em caixas de texto.
- **Sistema de Tags Flexível**: Personalize a aparência de cada controle usando a propriedade `.Tag`, com suporte para estilos pré-definidos e modulares.

## Temas Disponíveis

A classe inclui a enumeração `Temas` para facilitar a aplicação dos temas:

- `TemaAutomatico`: Acompanha o tema atual do Office.
- `TemaCinzaEscuro`: Tema escuro em tons de cinza.
- `TemaPreto`: Tema preto completo.
- `TemaBranco`: Tema claro padrão.
- `TemaConfWindows`: Segue o tema do Windows (usa o Tema Preto como fallback).
- `TemaColorido`: Tema colorido padrão.

## Sistema de Tags - Formato e Sintaxe

A personalização de cada controle é feita através da sua propriedade `.Tag`, seguindo a sintaxe `"{tema:estilo}:{tema:estilo}:..."`.

**Tipos de Estilo:**

1. **Estilos Pré-definidos**: Nomes curtos para estilos comuns.
    
    - **Exemplo**: `{tmpr:lblbtd01}:{tmbr:lblbtd01}` (Aplica o estilo "label com borda de destaque" para o tema preto e branco).
    
2. **Estilos Modulares**: Permite customizar propriedades individuais usando o formato `"(prop=valor|prop=valor)"`.
    
    - **Propriedades disponíveis**: `bc` (BackColor), `fc` (ForeColor), `bd` (BorderColor), `bs` (BackStyle), `by` (BorderStyle), `fs` (Font.Size).
        
    - **Exemplo**: `{tmpr(bc=d2|bs=1|bd=d1|by=1|fc=l1)}` (Personaliza o fundo, estilo, borda e cor da fonte para o tema preto).

## Como Usar

O processo de implementação da classe em um UserForm é simples.

1. **Adicione a Classe ao seu Projeto VBA**: Importe o arquivo `clsTemasOffice.txt` para o seu projeto VBA.
2. **Declaração da Classe**: No seu UserForm, declare uma instância da classe para gerenciar os temas. Você pode declarar instâncias adicionais para cada controle que terá um efeito individual (como hover).
    
    VBA
    ```
    ' Para temas principais
    Dim GTemas As New clsTemasOffice
    
    ' Para efeitos individuais
    Dim Efeitos(1 To 5) As New clsTemasOffice
    ```
    
3. **Inicialização no `UserForm_Initialize`**: Aplique o tema principal e configure a propriedade `.Tag` dos seus controles.
    
    VBA
    ```
    Private Sub UserForm_Initialize()
        ' Aplica o tema automático ao formulário
        GTemas.AplicarTema Me
    
        ' Configura a tag do controle para estilos específicos
        Me.Label1.Tag = "{tmpr:lblfcd01}:{tmbr:lblfcd01}:{tmci:lblfcd01}:{tmco:lblfcd01}"
        Me.TextBox1.Tag = "{tmpr(bc=d2|bs=1|bd=d1|by=1|fc=l1)}:{tmbr:lblbtd01}"
    
        ' Configura efeitos hover
        Efeitos(1).EfeitoHover Me, Me.Label1, "{tmpr:lblbtd11}:{tmbr:lblbtd11}", Me.Frame1
    
        ' Configura placeholder
        Efeitos(2).EfeitoPlaceHolder Me, Me.TextBox1, "Digite seu texto"
    End Sub
    ```
    
1. **Troca de Tema Dinâmica**: Mude o tema em tempo de execução usando o método `AplicarTema` com a enumeração `Temas`.
2. 
    VBA
    ```
    Private Sub AlterarParaTemaPreto()
        GTemas.AplicarTema Me, TemaPreto
    End Sub
    ```

## Melhores Práticas e Limitações

**Melhores Práticas:**

- Use uma instância principal (`GTemas`) para configurações globais.
- Use um _array_ de instâncias (`Efeitos()`) para controles individuais com _hover/placeholder_.
- Evite muitos efeitos de _hover_ simultâneos para manter a performance.

**Limitações Conhecidas:**
- Suporta apenas os controles padrão do MSForms.
- Efeitos de _hover_ podem não funcionar perfeitamente em controles dentro de _MultiPage_ ou com a propriedade `ZOrder` modificada dinamicamente.
- _Placeholders_ funcionam apenas em `TextBox` e `ComboBox`.
---

**Autor**: Adriano Furtado Lima
