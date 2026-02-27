# MemorialCAD
# Plugin Memorial Descritivo v2.0
## AutoCAD Civil 3D 2024 — Guia Completo

---

## O que este plugin entrega

### Comando: `MEMORIAL`

1. Você seleciona um ou mais **Parcels** no Civil 3D
2. Abre uma **janela WPF** com 3 etapas:
   - **① Dados do documento** — nome do loteamento, matrícula, proprietário, responsável técnico, CREA, município
   - **② Revisão das confrontações** — tabela editável com Classificação / Vértices / Comprimento / Azimute / Confrontante. Você escolhe manualmente qual face é a **Frente** e as demais se ajustam automaticamente
   - **③ Exportar** — gera `.docx` e/ou `.xlsx` na pasta de sua escolha

### Arquivos gerados

**DOCX (.docx)** — segue exatamente o layout do seu template:
- Times New Roman 12pt
- Cabeçalho: MEMORIAL DESCRITIVO / nome do loteamento / situação titulada
- Por lote: título estilo Heading1 + tabela de confrontações (2 colunas) com azimute em cada linha
- Texto descritivo *"Daí segue com azimute de..., confrontando com..., numa extensão de... m"*
- Tabela de coordenadas UTM dos vértices (V1, V2, V3...)

**XLSX (.xlsx)** — planilha Excel:
- Aba **Resumo**: todos os lotes em uma tabela (quadra, frente, fundo, dir, esq, área)
- **Uma aba por lote**: tabela de confrontações com azimutes + tabela de vértices UTM

### Exemplo de saída (tabela de confrontações)

```
Frente:   12,000 m com a Rua Ângelo Sichinel. Az: 090°00'00"
Fundo:    12,000 m com o Lote 05. Az: 270°00'00"
Direita:  23,340 m com o Lote 08. Az: 000°00'00"
Esquerda: 23,340 m com o Lote 06. Az: 180°00'00"
Área:     280,120 m² (duzentos e oitenta metros quadrados e doze centímetros quadrados.)
```

### Exemplo de texto descritivo gerado

```
Inicia-se a descrição no Vértice V1, deste ponto segue com azimute de 090°00'00",
confrontando com a Rua Ângelo Sichinel, numa extensão de 12,000 m, até o Vértice V2;
daí segue com azimute de 000°00'00", confrontando com o Lote 08, numa extensão de
23,340 m, até o Vértice V3; daí segue...
```

---

## Estrutura de arquivos

```
MemorialDescritivo/
├── MemorialDescritivo.csproj       ← Projeto Visual Studio
├── LEIAME.md                       ← Este arquivo
├── src/
│   ├── MemorialPlugin.cs           ← Comando MEMORIAL + modelos de dados + geometria
│   ├── GeradorDocx.cs              ← Gera o .docx no layout do seu template
│   └── GeradorXlsx.cs              ← Gera o .xlsx com abas por lote
└── wpf/
    ├── JanelaMemorial.xaml         ← Interface gráfica (layout WPF)
    └── JanelaMemorial.cs           ← Lógica da janela (code-behind)
```

---

## Passo 1 — Pré-requisitos

| Software | Versão | Onde baixar |
|---|---|---|
| AutoCAD Civil 3D | 2024 | Autodesk Account |
| Visual Studio Community | 2022 | https://visualstudio.microsoft.com/pt-br/ |
| .NET Framework | 4.8 | Incluído no Windows 10/11 |

No Visual Studio, durante a instalação ou em **Ferramentas → Obter Ferramentas e Recursos**, marque:
- ✅ **Desenvolvimento para desktop com .NET**
- ✅ **Desenvolvimento de desktop do Windows (WPF/WinForms)**

---

## Passo 2 — Compilar

### 2.1 — Abrir o projeto

```
1. Abra o Visual Studio 2022
2. Arquivo → Abrir → Projeto/Solução
3. Selecione: MemorialDescritivo.csproj
```

### 2.2 — Verificar caminho do Civil 3D

No arquivo `MemorialDescritivo.csproj`, confirme:
```xml
<Civil3DPath>C:\Program Files\Autodesk\AutoCAD 2024</Civil3DPath>
```

Verifique se o caminho existe e contém os arquivos:
- `accoremgd.dll`
- `acdbmgd.dll`
- `acmgd.dll`
- `AecBaseMgd.dll`
- `AeccDbMgd.dll`

### 2.3 — Restaurar pacotes NuGet

```
1. No Solution Explorer, clique com botão direito no projeto
2. "Restaurar Pacotes NuGet"
   (instala automaticamente: DocumentFormat.OpenXml e ClosedXML)
```

### 2.4 — Compilar

```
Menu: Build → Build Solution   (ou Ctrl+Shift+B)
```

O arquivo `MemorialDescritivo.dll` será gerado em:
```
bin\Debug\MemorialDescritivo.dll
```

---

## Passo 3 — Instalar no Civil 3D

### Opção A: NETLOAD (manual, por sessão)

```
1. Abra o AutoCAD Civil 3D 2024
2. Digite na linha de comando: NETLOAD
3. Navegue até: MemorialDescritivo\bin\Debug\
4. Selecione: MemorialDescritivo.dll
5. Clique Abrir
```

### Opção B: Carregamento automático permanente (RECOMENDADO)

Crie ou edite o arquivo `acad.lsp`:
```
Caminho: C:\Users\[SEU USUÁRIO]\AppData\Roaming\Autodesk\AutoCAD 2024\R24.3\ptb\support\acad.lsp
```

Adicione esta linha:
```lisp
(command "NETLOAD" "C:/CAMINHO_COMPLETO/bin/Debug/MemorialDescritivo.dll")
```

Salve e reinicie o Civil 3D.

### Opção C: Pelo CUI (interface customizável)

```
1. Manage → Customization → Edit Program Parameters (acad.pgp)
2. Ou: Manage → User Interface → acad.cui
3. Em "LISP Files to Load" adicione o caminho do .dll
```

---

## Passo 4 — Usar o plugin

### Preparação do desenho

Para funcionar corretamente:

1. **Parcels** devem estar dentro de um **Site** do Civil 3D
2. **Ruas/vias** devem ser **Alignments** (não polylines soltas)
3. Os Parcels precisam **compartilhar arestas** com os vizinhos (tolerância: 1,0 m)
4. Os Alignments devem ter o **nome da rua** (ex: "Rua Ângelo Sichinel")

### Executar

```
1. Abra o desenho com os Parcels
2. Digite: MEMORIAL
3. Selecione os lotes com o mouse (clique ou janela de seleção)
4. Pressione ENTER para confirmar
5. A janela WPF abrirá
```

### Usar a janela WPF

**Etapa ①** — Preencha os dados do loteamento (uma vez para todos os lotes)

**Etapa ②** — Para cada lote na lista à esquerda:
- Use o **ComboBox "Definir qual lado é a FRENTE"** para escolher a face correta
- As demais faces (Fundo, Direita, Esquerda) se ajustam automaticamente pelo ângulo
- Clique na coluna **"Confrontante"** para editar o texto manualmente se necessário
- Confira os azimutes calculados automaticamente

**Etapa ③** — Marque .docx e/ou .xlsx, escolha a pasta e clique **EXPORTAR**

---

## Lógica de classificação das faces

| Face | Critério |
|---|---|
| **Frente** | Escolha manual do usuário no ComboBox |
| **Fundo** | Face com azimute oposto à Frente (±45° de tolerância) |
| **Direita** | Face perpendicular, rotação horária em relação à Frente (+90° a +270°) |
| **Esquerda** | Face perpendicular, rotação anti-horária em relação à Frente (+270° a +360°) |
| **Outro** | Faces de lotes irregulares que não se enquadram nos casos acima |

## Cálculo do azimute

O azimute é calculado a partir do **Norte geográfico** no sentido **horário**, em **Graus-Minutos-Segundos**:

```
Az = atan2(ΔE, ΔN)  →  normalizado para 0°–360°
```

Exemplo: face apontando para Leste = **090°00'00"**

---

## Dependências NuGet (instaladas automaticamente)

| Pacote | Versão | Uso |
|---|---|---|
| `DocumentFormat.OpenXml` | 2.20.0 | Gerar arquivos .docx |
| `ClosedXML` | 0.102.2 | Gerar arquivos .xlsx |

---

## Solução de problemas

**"NETLOAD falhou" ou "DLL não carregou"**
→ Verifique se o Civil 3D 2024 está instalado no caminho configurado
→ Confirme que o .NET Framework 4.8 está instalado
→ Tente desbloquear o arquivo: clique com botão direito → Propriedades → Desbloquear

**"Confrontante aparece como 'Área Pública'"**
→ O parcel não encontrou vizinhos dentro da tolerância de 1,0 m
→ Verifique se os parcels estão no mesmo Site
→ Edite manualmente na coluna Confrontante da janela WPF

**"Azimute parece errado"**
→ Confirme o sistema de coordenadas do desenho (o plugin usa as coordenadas UTM brutas)
→ O azimute é calculado no plano 2D local do desenho

**"Erro ao abrir a janela WPF"**
→ Adicione referências WPF: clique direito no projeto → Adicionar → Referência → Framework → PresentationCore, PresentationFramework, WindowsBase

**"Frente e Fundo invertidos"**
→ Use o ComboBox na janela para trocar manualmente

---

## Personalizar

### Alterar a tolerância de vizinhança

Em `MemorialPlugin.cs`, função `IdentificarConfrontante`:
```csharp
double tolerancia = 1.0; // ← aumente se os parcels não tocam exatamente
```

### Alterar casas decimais dos comprimentos

Em `GeradorDocx.cs` e `GeradorXlsx.cs`, troque `"F3"` por `"F2"`:
```csharp
string comp = lado.Comprimento.ToString("F3", ...);
//                                        ↑ 3 casas decimais
```

### Adicionar logo do escritório no Word

Em `GeradorDocx.cs`, na função `Gerar()`, antes do título, adicione:
```csharp
// Adicionar imagem
var imagePart = main.AddImagePart(ImagePartType.Png);
using (var stream = File.OpenRead("logo.png"))
    imagePart.FeedData(stream);
// (ver documentação OpenXml para inserir DrawingML)
```
