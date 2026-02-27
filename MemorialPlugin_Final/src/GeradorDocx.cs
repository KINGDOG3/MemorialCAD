using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MemorialPlugin.WPF;

namespace MemorialPlugin
{
    /// <summary>
    /// Gera o .docx seguindo fielmente o layout do template "Reserva do Bosque".
    /// Fonte: Times New Roman 12pt. Tabela de confrontações 2 colunas.
    /// Acrescenta tabela de vértices UTM e texto descritivo com azimutes.
    /// </summary>
    public static class GeradorDocx
    {
        // Largura da página A4 com margens 3cm/2cm = ~14cm de área útil = 7938 DXA
        private const int LARGURA_TOTAL = 8510;   // área útil em DXA (~15cm)
        private const int COL1_CONF     = 2268;   // "Frente:", "Fundo:" etc. (~4cm)
        private const int COL2_CONF     = 6242;   // descrição (~11cm)
        private const int COL_VERT_NR   = 756;    // V1, V2...
        private const int COL_VERT_E    = 2646;   // Este
        private const int COL_VERT_N    = 2646;   // Norte
        private const int COL_VERT_AZ   = 2462;   // Azimute

        public static void Gerar(List<LoteInfo> lotes, CabecalhoDoc cab, string caminho)
        {
            using (var docx = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document))
            {
                var main = docx.AddMainDocumentPart();
                main.Document = new Document();
                var body = new Body();

                // Estilos replicando o template
                var stylePart = main.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = CriarEstilos();

                // ── CABEÇALHO DO DOCUMENTO ────────────────────────
                body.Append(Para("MEMORIAL DESCRITIVO - LOTEAMENTO", "Heading1", true, 14, true));
                body.Append(Para(cab.NomeLoteamento.ToUpper(), "Normal", true, 14, true));
                body.Append(ParaVazio());

                body.Append(Para("SITUAÇÃO TITULADA", "Heading1", true, 12, true));
                body.Append(ParaVazio());

                if (!string.IsNullOrEmpty(cab.Matricula))
                    body.Append(ParaMixed("MATRÍCULA N° ", cab.Matricula + ".", bold1: true));

                if (!string.IsNullOrEmpty(cab.Proprietario))
                    body.Append(ParaMixed("PROPRIETÁRIO: ", cab.Proprietario + ".", bold1: true));

                if (!string.IsNullOrEmpty(cab.ResponsavelTecnico))
                    body.Append(ParaMixed("RESPONSÁVEL TÉCNICO: ",
                        $"{cab.ResponsavelTecnico} — CREA n° {cab.Crea}", bold1: true));

                if (!string.IsNullOrEmpty(cab.Municipio))
                    body.Append(ParaMixed("MUNICÍPIO: ", $"{cab.Municipio}/{cab.Estado}", bold1: true));

                body.Append(ParaVazio());
                body.Append(Para("SITUAÇÃO LOTEADA", "Normal", true, 14, true));
                body.Append(ParaVazio());

                // ── QUADRAS E LOTES ───────────────────────────────
                // Agrupar lotes por quadra
                var quadras = lotes.GroupBy(l => l.Quadra ?? "").ToList();

                foreach (var grpQuadra in quadras)
                {
                    string nomeQuadra = string.IsNullOrEmpty(grpQuadra.Key)
                        ? "" : grpQuadra.Key;

                    if (!string.IsNullOrEmpty(nomeQuadra))
                    {
                        body.Append(Para($"{nomeQuadra.ToUpper()} - ({FormatoGrupo(grpQuadra.ToList())}):",
                            "Heading1", true, 12, false));
                        body.Append(Para($"Quantidade de lotes: {grpQuadra.Count()}.",
                            "Normal", false, 12, false));
                        body.Append(ParaVazio());
                    }

                    foreach (var lote in grpQuadra)
                        body.Append(GerarBlocoLote(lote));
                }

                // ── RODAPÉ ────────────────────────────────────────
                body.Append(ParaVazio());
                body.Append(Para(
                    $"{cab.Municipio}/{cab.Estado}, {cab.Data:dd 'de' MMMM 'de' yyyy}",
                    "Normal", false, 12, true));
                body.Append(ParaVazio());
                body.Append(Para("_________________________________", "Normal", false, 12, true));
                body.Append(Para(cab.ResponsavelTecnico, "Normal", false, 12, true));
                body.Append(Para($"CREA n° {cab.Crea}", "Normal", false, 12, true));

                // Configuração de página A4
                body.Append(new SectionProperties(
                    new PageSize { Width = 11906, Height = 16838 },
                    new PageMargin { Top = 1134, Right = 1134, Bottom = 1134, Left = 1701 }
                ));

                main.Document.Append(body);
                main.Document.Save();
            }
        }

        // ================================================================
        // BLOCO DE UM LOTE
        // ================================================================

        private static OpenXmlCompositeElement GerarBlocoLote(LoteInfo lote)
        {
            // Como OpenXml não suporta fragmento, usamos uma Section fake
            // Retornamos apenas a primeira tabela e os parágrafos via lista
            // SOLUÇÃO: devolver Body inteiro — mas como não podemos, retornamos
            // um elemento composto usando um Paragraph que age como wrapper.
            // Na prática, chamamos GerarElementosLote e os adicionamos ao body.
            // Como este método deve retornar um único elemento, usamos um
            // parágrafo marcador e adicionamos o restante após.
            // --> Refatorado: ver GerarElementosLote() abaixo.
            throw new InvalidOperationException("Use GerarElementosLote()");
        }

        public static void AdicionarLoteAoBody(Body body, LoteInfo lote)
        {
            // Título do lote: "LOTE 07 (Formato Regular): situado ao lado..."
            string distanciaEsquina = ""; // Pode ser calculado no futuro
            string ladoPar = "par";       // Lógica futura
            string frente = lote.Frente?.Confrontante ?? "";

            string titulo = $"{lote.Nome.ToUpper()} ({lote.Formato}): " +
                            $"situado com {frente}.";

            body.Append(Para(titulo, "Heading1", true, 12, false));
            body.Append(ParaVazio());

            // ── Tabela de confrontações ──────────────────────────
            body.Append(TabelaConfrontacoes(lote));
            body.Append(ParaVazio());

            // ── Texto descritivo com azimutes ────────────────────
            body.Append(Para("Descrição dos limites:", "Normal", true, 12, false));
            body.Append(GerarTextoDescritivo(lote));
            body.Append(ParaVazio());

            // ── Tabela de vértices UTM ───────────────────────────
            body.Append(Para("Coordenadas UTM dos vértices:", "Normal", true, 12, false));
            body.Append(TabelaVertices(lote));
            body.Append(ParaVazio());
        }

        // ================================================================
        // TABELA DE CONFRONTAÇÕES (replicando o template)
        // ================================================================

        private static Table TabelaConfrontacoes(LoteInfo lote)
        {
            var tbl = new Table();
            tbl.Append(TblProp(LARGURA_TOTAL));

            // Ordem fixa: Frente, Fundo, Direita, Esquerda, Outros, Área
            var ordem = new[] { "Frente", "Fundo", "Direita", "Esquerda" };
            foreach (string face in ordem)
            {
                var lado = lote.Lados.FirstOrDefault(l => l.ClassificacaoFace == face);
                if (lado == null) continue;
                tbl.Append(LinhaConfrontacao(face + ":", FormatarDescricaoLado(lado)));
            }

            // Lados extras (irregulares)
            foreach (var outro in lote.OutrosLados)
                tbl.Append(LinhaConfrontacao("Lado:", FormatarDescricaoLado(outro)));

            // Área
            string areaTexto = $"{lote.Area.ToString("N2", new System.Globalization.CultureInfo("pt-BR"))} m² " +
                                $"({lote.AreaPorExtenso}.)";
            tbl.Append(LinhaConfrontacao("Área:", areaTexto));

            return tbl;
        }

        private static string FormatarDescricaoLado(LadoInfo lado)
        {
            string comp = lado.Comprimento.ToString("F3", new System.Globalization.CultureInfo("pt-BR"));
            string az   = lado.AzimuteFormatado;
            string conf = lado.Confrontante;

            // Formato: "12,000 m com a Rua Ângelo Sichinel. Az: 045°30'12""
            return $"{comp} m com {conf}. Az: {az}";
        }

        private static TableRow LinhaConfrontacao(string col1, string col2)
        {
            var brd = new TableCellBorderType[] { };
            var row = new TableRow();

            // Célula 1 – label em negrito
            var c1 = new TableCell(
                TcProp(COL1_CONF),
                new Paragraph(Run(col1, bold: true, size: 24))
            );

            // Célula 2 – conteúdo normal
            var c2 = new TableCell(
                TcProp(COL2_CONF),
                new Paragraph(Run(col2, bold: false, size: 24))
            );

            row.Append(c1, c2);
            return row;
        }

        // ================================================================
        // TEXTO DESCRITIVO "DAEI SEGUE..."
        // ================================================================

        private static Paragraph GerarTextoDescritivo(LoteInfo lote)
        {
            if (lote.Lados.Count == 0) return ParaVazio();

            // Ordenar lados pela classificação: Frente → Direita → Fundo → Esquerda
            var ordemFace = new[] { "Frente", "Direita", "Fundo", "Esquerda", "Outro" };
            var ladosOrdenados = lote.Lados
                .OrderBy(l => Array.IndexOf(ordemFace, l.ClassificacaoFace ?? "Outro"))
                .ToList();

            var sb = new System.Text.StringBuilder();
            sb.Append($"Inicia-se a descrição no Vértice V1, ");

            for (int i = 0; i < ladosOrdenados.Count; i++)
            {
                var l = ladosOrdenados[i];
                string comp = l.Comprimento.ToString("F3",
                    new System.Globalization.CultureInfo("pt-BR"));

                if (i == 0)
                    sb.Append($"deste ponto segue com azimute de {l.AzimuteFormatado}, " +
                               $"confrontando com {l.Confrontante}, " +
                               $"numa extensão de {comp} m, " +
                               $"até o Vértice V{l.Para}; ");
                else if (i < ladosOrdenados.Count - 1)
                    sb.Append($"daí segue com azimute de {l.AzimuteFormatado}, " +
                               $"confrontando com {l.Confrontante}, " +
                               $"numa extensão de {comp} m, " +
                               $"até o Vértice V{l.Para}; ");
                else
                    sb.Append($"daí segue com azimute de {l.AzimuteFormatado}, " +
                               $"confrontando com {l.Confrontante}, " +
                               $"numa extensão de {comp} m, " +
                               $"retornando ao ponto inicial, Vértice V1.");
            }

            return Para(sb.ToString(), "Normal", false, 24, false);
        }

        // ================================================================
        // TABELA DE VÉRTICES UTM
        // ================================================================

        private static Table TabelaVertices(LoteInfo lote)
        {
            var tbl = new Table();
            tbl.Append(TblProp(LARGURA_TOTAL));

            // Cabeçalho
            tbl.Append(LinhaVertice("Vértice", "E (m)", "N (m)", "Azimute próx. lado", header: true));

            for (int i = 0; i < lote.Vertices.Count; i++)
            {
                var v = lote.Vertices[i];
                // Azimute do lado que SAI deste vértice
                var lado = lote.Lados.FirstOrDefault(l => l.De == v.Numero);
                string az = lado != null ? lado.AzimuteFormatado : "—";

                tbl.Append(LinhaVertice(
                    $"V{v.Numero}",
                    v.E.ToString("F3", new System.Globalization.CultureInfo("pt-BR")),
                    v.N.ToString("F3", new System.Globalization.CultureInfo("pt-BR")),
                    az,
                    header: false
                ));
            }

            return tbl;
        }

        private static TableRow LinhaVertice(string vt, string e, string n, string az, bool header)
        {
            var row = new TableRow();
            string fillHdr = "D6E4F0";
            string fill    = header ? fillHdr : "FFFFFF";

            row.Append(CelulaVertice(vt,  COL_VERT_NR, header, fill));
            row.Append(CelulaVertice(e,   COL_VERT_E,  header, fill));
            row.Append(CelulaVertice(n,   COL_VERT_N,  header, fill));
            row.Append(CelulaVertice(az,  COL_VERT_AZ, header, fill));
            return row;
        }

        private static TableCell CelulaVertice(string txt, int width, bool bold, string fill)
        {
            return new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = fill }
                ),
                new Paragraph(Run(txt, bold: bold, size: 20))
            );
        }

        // ================================================================
        // HELPERS DE CONSTRUÇÃO OpenXml
        // ================================================================

        private static TableProperties TblProp(int width)
        {
            return new TableProperties(
                new TableWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    new TopBorder    { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder   { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder  { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4, Color = "000000" }
                )
            );
        }

        private static TableCellProperties TcProp(int width)
        {
            return new TableCellProperties(
                new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableCellMargin(
                    new TopMargin    { Width = "80",  Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "80",  Type = TableWidthUnitValues.Dxa },
                    new LeftMargin   { Width = "120", Type = TableWidthUnitValues.Dxa },
                    new RightMargin  { Width = "120", Type = TableWidthUnitValues.Dxa }
                )
            );
        }

        private static Run Run(string texto, bool bold, int size)
        {
            return new Run(
                new RunProperties(
                    new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                    new Bold { Val = OnOffValue.FromBoolean(bold) },
                    new FontSize { Val = size.ToString() }
                ),
                new Text(texto) { Space = SpaceProcessingModeValues.Preserve }
            );
        }

        private static Paragraph Para(string txt, string estilo, bool bold, int size, bool centralizar)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = estilo },
                    new Justification
                    {
                        Val = centralizar
                            ? JustificationValues.Center
                            : JustificationValues.Both
                    }
                ),
                Run(txt, bold, size)
            );
        }

        // Parágrafo com dois runs: rótulo negrito + conteúdo normal
        private static Paragraph ParaMixed(string rotulo, string conteudo, bool bold1)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "Normal" },
                    new Justification { Val = JustificationValues.Both }
                ),
                Run(rotulo, bold: true,  size: 24),
                Run(conteudo, bold: false, size: 24)
            );
        }

        private static Paragraph ParaVazio()
        {
            return new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines { Before = "0", After = "0", Line = "240" }
                )
            );
        }

        private static string FormatoGrupo(List<LoteInfo> lotes)
        {
            return lotes.Any(l => l.Formato == "Irregular") ? "Formato Irregular" : "Formato Regular";
        }

        private static Styles CriarEstilos()
        {
            return new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(
                        new RunPropertiesBaseStyle(
                            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                            new FontSize { Val = "24" }  // 12pt
                        )
                    ),
                    new ParagraphPropertiesDefault(
                        new ParagraphPropertiesBaseStyle(
                            new Justification { Val = JustificationValues.Both },
                            new SpacingBetweenLines { Before = "0", After = "120", Line = "276" }
                        )
                    )
                ),
                new Style(
                    new StyleName { Val = "Normal" },
                    new PrimaryStyle()
                ) { Type = StyleValues.Paragraph, StyleId = "Normal" },
                new Style(
                    new StyleName { Val = "heading 1" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(
                        new Bold(),
                        new FontSize { Val = "24" }
                    )
                ) { Type = StyleValues.Paragraph, StyleId = "Heading1" }
            );
        }
    }

    // ================================================================
    // EXTENSÃO: Adicionar bloco de lote ao Body
    // ================================================================
    public static class BodyExtensions
    {
        public static void AppendLote(this Body body, LoteInfo lote)
        {
            GeradorDocx.AdicionarLoteAoBody(body, lote);
        }
    }
}
