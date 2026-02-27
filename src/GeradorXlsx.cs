using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using MemorialPlugin.WPF;

namespace MemorialPlugin
{
    /// <summary>
    /// Gera o .xlsx com abas por lote.
    /// Usa ClosedXML (NuGet: ClosedXML).
    /// </summary>
    public static class GeradorXlsx
    {
        public static void Gerar(List<LoteInfo> lotes, CabecalhoDoc cab, string caminho)
        {
            using (var wb = new XLWorkbook())
            {
                // ── Aba Resumo ──────────────────────────────────
                var wsResumo = wb.Worksheets.Add("Resumo");
                EscreverResumo(wsResumo, lotes, cab);

                // ── Uma aba por lote ─────────────────────────────
                foreach (var lote in lotes)
                {
                    // Nome da aba: máximo 31 chars
                    string nomePlanilha = $"{lote.Nome} {lote.Quadra}";
                    if (nomePlanilha.Length > 31) nomePlanilha = nomePlanilha.Substring(0, 31);
                    // Remover chars inválidos
                    foreach (char c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
                        nomePlanilha = nomePlanilha.Replace(c, '-');

                    var ws = wb.Worksheets.Add(nomePlanilha);
                    EscreverLote(ws, lote, cab);
                }

                wb.SaveAs(caminho);
            }
        }

        // ================================================================
        // ABA RESUMO
        // ================================================================

        private static void EscreverResumo(IXLWorksheet ws, List<LoteInfo> lotes, CabecalhoDoc cab)
        {
            // Cabeçalho
            ws.Cell("A1").Value = "MEMORIAL DESCRITIVO — RESUMO";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Range("A1:G1").Merge();

            ws.Cell("A2").Value = $"Loteamento: {cab.NomeLoteamento}";
            ws.Cell("A3").Value = $"Matrícula: {cab.Matricula}";
            ws.Cell("A4").Value = $"Data: {cab.Data:dd/MM/yyyy}";
            ws.Range("A2:G4").Merge();

            // Cabeçalho da tabela
            int linha = 6;
            string[] headers = { "Quadra", "Lote", "Formato", "Frente (confrontante)",
                                  "Frente (m)", "Fundo (m)", "Dir. (m)", "Esq. (m)",
                                  "Área (m²)", "Área por extenso" };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(linha, i + 1).Value = headers[i];
                ws.Cell(linha, i + 1).Style.Font.Bold = true;
                ws.Cell(linha, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
                ws.Cell(linha, i + 1).Style.Font.FontColor = XLColor.White;
            }

            // Dados
            bool alternar = false;
            foreach (var lote in lotes)
            {
                linha++;
                var fill = alternar ? XLColor.FromHtml("#EBF3FB") : XLColor.White;
                alternar = !alternar;

                ws.Cell(linha, 1).Value = lote.Quadra;
                ws.Cell(linha, 2).Value = lote.Nome;
                ws.Cell(linha, 3).Value = lote.Formato;
                ws.Cell(linha, 4).Value = lote.Frente?.Confrontante ?? "";
                ws.Cell(linha, 5).Value = lote.Frente?.Comprimento ?? 0;
                ws.Cell(linha, 6).Value = lote.Fundo?.Comprimento  ?? 0;
                ws.Cell(linha, 7).Value = lote.Direita?.Comprimento ?? 0;
                ws.Cell(linha, 8).Value = lote.Esquerda?.Comprimento ?? 0;
                ws.Cell(linha, 9).Value = lote.Area;
                ws.Cell(linha, 10).Value = lote.AreaPorExtenso;

                // Formatar números com vírgula (3 casas)
                for (int c = 5; c <= 9; c++)
                    ws.Cell(linha, c).Style.NumberFormat.Format = "#,##0.000";

                // Colorir fundo alternado
                ws.Range(linha, 1, linha, 10).Style.Fill.BackgroundColor = fill;
            }

            // Total
            linha++;
            ws.Cell(linha, 8).Value = "TOTAL LOTES:";
            ws.Cell(linha, 8).Style.Font.Bold = true;
            ws.Cell(linha, 9).Value = lotes.Sum(l => l.Area);
            ws.Cell(linha, 9).Style.Font.Bold = true;
            ws.Cell(linha, 9).Style.NumberFormat.Format = "#,##0.000";

            ws.Columns().AdjustToContents();
        }

        // ================================================================
        // ABA DE CADA LOTE
        // ================================================================

        private static void EscreverLote(IXLWorksheet ws, LoteInfo lote, CabecalhoDoc cab)
        {
            int row = 1;

            // Título
            ws.Cell(row, 1).Value = $"{lote.Nome.ToUpper()} — {lote.Quadra}";
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontSize = 13;
            ws.Range(row, 1, row, 5).Merge();
            row += 2;

            // Informações gerais
            CelulaLabel(ws, row, 1, "Formato:"); ws.Cell(row, 2).Value = lote.Formato; row++;
            CelulaLabel(ws, row, 1, "Área:"); 
            ws.Cell(row, 2).Value = lote.Area;
            ws.Cell(row, 2).Style.NumberFormat.Format = "#,##0.000";
            ws.Cell(row, 3).Value = "m²"; row++;
            CelulaLabel(ws, row, 1, "Área por extenso:"); ws.Cell(row, 2).Value = lote.AreaPorExtenso;
            ws.Range(row, 2, row, 5).Merge(); row += 2;

            // ── Tabela de confrontações com azimutes ─────────────
            CelulaHeader(ws, row, 1, "Face");
            CelulaHeader(ws, row, 2, "Confrontante");
            CelulaHeader(ws, row, 3, "Comprimento (m)");
            CelulaHeader(ws, row, 4, "Azimute");
            CelulaHeader(ws, row, 5, "Vértices");
            row++;

            var ordemFace = new[] { "Frente", "Fundo", "Direita", "Esquerda", "Outro" };
            var ladosOrdenados = lote.Lados
                .OrderBy(l => Array.IndexOf(ordemFace, l.ClassificacaoFace ?? "Outro"))
                .ToList();

            bool alt = false;
            foreach (var lado in ladosOrdenados)
            {
                var fill = alt ? XLColor.FromHtml("#EBF3FB") : XLColor.White;
                ws.Cell(row, 1).Value = lado.ClassificacaoFace;
                ws.Cell(row, 2).Value = lado.Confrontante;
                ws.Cell(row, 3).Value = lado.Comprimento;
                ws.Cell(row, 3).Style.NumberFormat.Format = "#,##0.000";
                ws.Cell(row, 4).Value = lado.AzimuteFormatado;
                ws.Cell(row, 5).Value = $"V{lado.De}→V{lado.Para}";
                ws.Range(row, 1, row, 5).Style.Fill.BackgroundColor = fill;
                alt = !alt;
                row++;
            }
            row++;

            // ── Tabela de vértices UTM ───────────────────────────
            CelulaHeader(ws, row, 1, "Vértice");
            CelulaHeader(ws, row, 2, "E (m) — UTM");
            CelulaHeader(ws, row, 3, "N (m) — UTM");
            CelulaHeader(ws, row, 4, "Azimute (saída)");
            row++;

            alt = false;
            foreach (var v in lote.Vertices)
            {
                var fill = alt ? XLColor.FromHtml("#EBF3FB") : XLColor.White;
                var ladoSaida = lote.Lados.FirstOrDefault(l => l.De == v.Numero);

                ws.Cell(row, 1).Value = $"V{v.Numero}";
                ws.Cell(row, 2).Value = v.E;
                ws.Cell(row, 2).Style.NumberFormat.Format = "#,##0.000";
                ws.Cell(row, 3).Value = v.N;
                ws.Cell(row, 3).Style.NumberFormat.Format = "#,##0.000";
                ws.Cell(row, 4).Value = ladoSaida?.AzimuteFormatado ?? "—";
                ws.Range(row, 1, row, 4).Style.Fill.BackgroundColor = fill;
                alt = !alt;
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        // ── Utilitários de células ──────────────────────────────
        private static void CelulaHeader(IXLWorksheet ws, int row, int col, string texto)
        {
            ws.Cell(row, col).Value = texto;
            ws.Cell(row, col).Style.Font.Bold = true;
            ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
            ws.Cell(row, col).Style.Font.FontColor = XLColor.White;
        }

        private static void CelulaLabel(IXLWorksheet ws, int row, int col, string texto)
        {
            ws.Cell(row, col).Value = texto;
            ws.Cell(row, col).Style.Font.Bold = true;
            ws.Cell(row, col).Style.Font.FontColor = XLColor.FromHtml("#546E7A");
        }
    }
}
