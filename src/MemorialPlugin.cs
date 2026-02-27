using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

[assembly: CommandClass(typeof(MemorialPlugin.MemorialCommands))]

namespace MemorialPlugin
{
    // ================================================================
    // MODELOS DE DADOS
    // ================================================================

    public class VerticeInfo
    {
        public int Numero { get; set; }          // V1, V2, V3...
        public double E { get; set; }            // Coordenada Este (X) UTM
        public double N { get; set; }            // Coordenada Norte (Y) UTM
    }

    public class LadoInfo
    {
        public int De { get; set; }              // Vértice inicial (ex: V1)
        public int Para { get; set; }            // Vértice final   (ex: V2)
        public double Comprimento { get; set; }  // metros, 3 casas
        public double Azimute { get; set; }      // graus decimais  (0–360)
        public string AzimuteFormatado => FormatarAzimute(Azimute);
        public string Confrontante { get; set; } // "Rua X" ou "Lote 02"
        public string ClassificacaoFace { get; set; } // Frente/Fundo/Direita/Esquerda/Outro

        private static string FormatarAzimute(double az)
        {
            // Graus-Minutos-Segundos  ex: 45°30'12"
            az = ((az % 360) + 360) % 360;
            int graus = (int)az;
            double minFrac = (az - graus) * 60.0;
            int min = (int)minFrac;
            double seg = (minFrac - min) * 60.0;
            return $"{graus:D3}°{min:D2}'{seg:05.2f}\"";
        }
    }

    public class LoteInfo
    {
        public string Nome { get; set; }               // "Lote 07"
        public string Quadra { get; set; }             // "Quadra 555"
        public string Formato { get; set; }            // "Regular" / "Irregular"
        public double Area { get; set; }               // m²
        public string AreaPorExtenso => NumeroPorExtenso.AreaCompleta(Area);
        public List<VerticeInfo> Vertices { get; set; } = new List<VerticeInfo>();
        public List<LadoInfo> Lados { get; set; } = new List<LadoInfo>();
        // Lados classificados por posição
        public LadoInfo Frente => Lados.FirstOrDefault(l => l.ClassificacaoFace == "Frente");
        public LadoInfo Fundo  => Lados.FirstOrDefault(l => l.ClassificacaoFace == "Fundo");
        public LadoInfo Direita => Lados.FirstOrDefault(l => l.ClassificacaoFace == "Direita");
        public LadoInfo Esquerda => Lados.FirstOrDefault(l => l.ClassificacaoFace == "Esquerda");
        public List<LadoInfo> OutrosLados => Lados.Where(l => l.ClassificacaoFace == "Outro").ToList();
    }

    // ================================================================
    // COMANDO PRINCIPAL
    // ================================================================

    public class MemorialCommands
    {
        [CommandMethod("MEMORIAL", CommandFlags.Modal)]
        public void AbrirJanelaMemorial()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== Plugin Memorial Descritivo v2.0 ===");
            ed.WriteMessage("\nSelecione os Parcels para o memorial (ENTER para confirmar):\n");

            // Filtro para selecionar apenas AECC_PARCEL
            TypedValue[] filtro = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "AECC_PARCEL")
            };
            SelectionFilter sf = new SelectionFilter(filtro);
            PromptSelectionResult psr = ed.GetSelection(sf);

            if (psr.Status != PromptStatus.OK || psr.Value.Count == 0)
            {
                ed.WriteMessage("\nNenhum Parcel selecionado. Operação cancelada.");
                return;
            }

            // Processar parcels e abrir janela WPF
            List<ObjectId> parcelIds = psr.Value.GetObjectIds().ToList();

            // Abrir janela WPF passando os IDs
            var janela = new MemorialPlugin.WPF.JanelaMemorial(parcelIds, doc);
            Application.ShowModalWindow(janela);
        }
    }

    // ================================================================
    // PROCESSADOR DE PARCELS
    // ================================================================

    public static class ParcelProcessor
    {
        public static List<LoteInfo> ProcessarParcels(List<ObjectId> parcelIds, Document acadDoc)
        {
            var lotes = new List<LoteInfo>();
            Database db = acadDoc.Database;
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Carregar todos os parcels e alignments do desenho
                List<Parcel> todosParcels = CarregarTodosParcels(civilDoc, db, tr);
                List<AlignmentInfo> alignments = CarregarAlignments(civilDoc, db, tr);

                foreach (ObjectId pid in parcelIds)
                {
                    try
                    {
                        Parcel parcel = tr.GetObject(pid, OpenMode.ForRead) as Parcel;
                        if (parcel == null) continue;

                        LoteInfo lote = ProcessarParcel(parcel, todosParcels, alignments, db, tr);
                        if (lote != null)
                            lotes.Add(lote);
                    }
                    catch (Exception ex)
                    {
                        acadDoc.Editor.WriteMessage($"\n[AVISO] Erro em parcel: {ex.Message}");
                    }
                }
                tr.Commit();
            }

            return lotes;
        }

        // ── Processa um parcel individual ──────────────────────────
        private static LoteInfo ProcessarParcel(Parcel parcel,
            List<Parcel> todosParcels,
            List<AlignmentInfo> alignments,
            Database db, Transaction tr)
        {
            LoteInfo lote = new LoteInfo
            {
                Nome   = parcel.Name ?? "Parcel",
                Quadra = ExtrairNomeSite(parcel, tr),
                Area   = parcel.Area
            };

            // Obter vértices da polyline do parcel
            List<Point2d> pontos = ObterVertices(parcel, tr);
            if (pontos.Count < 3) return null;

            // Fechar polígono (remover duplicata final se existir)
            if (pontos.First().GetDistanceTo(pontos.Last()) < 0.001)
                pontos.RemoveAt(pontos.Count - 1);

            // Montar lista de VerticeInfo
            for (int i = 0; i < pontos.Count; i++)
            {
                lote.Vertices.Add(new VerticeInfo
                {
                    Numero = i + 1,
                    E = pontos[i].X,
                    N = pontos[i].Y
                });
            }

            // Montar lados
            for (int i = 0; i < pontos.Count; i++)
            {
                Point2d p1 = pontos[i];
                Point2d p2 = pontos[(i + 1) % pontos.Count];

                double comprimento = p1.GetDistanceTo(p2);
                if (comprimento < 0.001) continue;

                double azimute = CalcularAzimute(p1, p2);
                Point2d meio = new Point2d((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2);

                string confrontante = IdentificarConfrontante(
                    meio, p1, p2, parcel, todosParcels, alignments, tr);

                lote.Lados.Add(new LadoInfo
                {
                    De = i + 1,
                    Para = (i + 1) % pontos.Count + 1,
                    Comprimento = comprimento,
                    Azimute = azimute,
                    Confrontante = confrontante
                });
            }

            // Determinar formato
            lote.Formato = lote.Lados.Count == 4 ? "Regular" : "Irregular";

            return lote;
        }

        // ── Identifica confrontante de um lado ─────────────────────
        private static string IdentificarConfrontante(
            Point2d meio, Point2d p1, Point2d p2,
            Parcel parcelAtual,
            List<Parcel> todosParcels,
            List<AlignmentInfo> alignments,
            Transaction tr)
        {
            double tolerancia = 1.0; // 1 metro

            // 1. Checar Alignments (ruas)
            foreach (var al in alignments)
            {
                if (DistanciaPontoPoliline(meio, al.Pontos) < tolerancia)
                    return al.Nome;
            }

            // 2. Checar outros Parcels
            double distMin = double.MaxValue;
            string vizinhoNome = null;

            foreach (Parcel viz in todosParcels)
            {
                if (viz.ObjectId == parcelAtual.ObjectId) continue;

                List<Point2d> vertsViz = ObterVertices(viz, tr);
                for (int j = 0; j < vertsViz.Count; j++)
                {
                    Point2d v1 = vertsViz[j];
                    Point2d v2 = vertsViz[(j + 1) % vertsViz.Count];
                    double d = DistanciaPontoSegmento(meio, v1, v2);
                    if (d < tolerancia && d < distMin)
                    {
                        distMin = d;
                        vizinhoNome = viz.Name;
                    }
                }
            }

            return vizinhoNome ?? "Área Pública";
        }

        // ── Classificação das faces (chamado após o usuário definir Frente) ──
        public static void ClassificarFaces(LoteInfo lote, int indexFrente)
        {
            if (lote.Lados.Count == 0) return;

            // Resetar classificações
            foreach (var l in lote.Lados) l.ClassificacaoFace = null;

            LadoInfo frente = lote.Lados[indexFrente];
            frente.ClassificacaoFace = "Frente";
            double azFrente = frente.Azimute;

            // Fundo = lado com azimute oposto (±180°, tolerância 45°)
            LadoInfo fundo = null;
            double menorDiffFundo = double.MaxValue;
            foreach (var l in lote.Lados)
            {
                if (l == frente) continue;
                double azOposto = (azFrente + 180.0) % 360.0;
                double diff = DiferencaAzimute(l.Azimute, azOposto);
                if (diff < 45.0 && diff < menorDiffFundo)
                {
                    menorDiffFundo = diff;
                    fundo = l;
                }
            }
            if (fundo != null) fundo.ClassificacaoFace = "Fundo";

            // Direita/Esquerda = lados perpendiculares
            // "Direita" ao olhar da frente para o fundo = +90° do azimute da frente
            foreach (var l in lote.Lados)
            {
                if (l.ClassificacaoFace != null) continue;

                // Ângulo relativo à frente (0–360)
                double rel = ((l.Azimute - azFrente) % 360.0 + 360.0) % 360.0;

                if (rel >= 45.0 && rel < 135.0)
                    l.ClassificacaoFace = "Esquerda";
                else if (rel >= 225.0 && rel < 315.0)
                    l.ClassificacaoFace = "Direita";
                else
                    l.ClassificacaoFace = "Outro";
            }
        }

        // ── Helpers geométricos ─────────────────────────────────────
        public static List<Point2d> ObterVertices(Parcel parcel, Transaction tr)
        {
            var pts = new List<Point2d>();
            try
            {
                foreach (ObjectId eid in parcel.GetEntityIds())
                {
                    var ent = tr.GetObject(eid, OpenMode.ForRead);

                    if (ent is Polyline pl)
                    {
                        for (int i = 0; i < pl.NumberOfVertices; i++)
                            pts.Add(pl.GetPoint2dAt(i));
                        return pts;
                    }
                    if (ent is Polyline2d p2d)
                    {
                        foreach (ObjectId vid in p2d)
                        {
                            var v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d;
                            if (v != null) pts.Add(new Point2d(v.Position.X, v.Position.Y));
                        }
                        return pts;
                    }
                    if (ent is Polyline3d p3d)
                    {
                        foreach (ObjectId vid in p3d)
                        {
                            var v = tr.GetObject(vid, OpenMode.ForRead) as PolylineVertex3d;
                            if (v != null) pts.Add(new Point2d(v.Position.X, v.Position.Y));
                        }
                        return pts;
                    }
                }
            }
            catch { }
            return pts;
        }

        private static double CalcularAzimute(Point2d de, Point2d para)
        {
            double dx = para.X - de.X;
            double dy = para.Y - de.Y;
            // Azimute: medido do Norte, sentido horário
            double az = Math.Atan2(dx, dy) * 180.0 / Math.PI;
            return ((az % 360.0) + 360.0) % 360.0;
        }

        private static double DiferencaAzimute(double a, double b)
        {
            double d = Math.Abs(a - b) % 360.0;
            return d > 180.0 ? 360.0 - d : d;
        }

        private static double DistanciaPontoSegmento(Point2d pt, Point2d a, Point2d b)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y;
            double lenSq = dx * dx + dy * dy;
            if (lenSq < 1e-10) return pt.GetDistanceTo(a);
            double t = Math.Max(0, Math.Min(1,
                ((pt.X - a.X) * dx + (pt.Y - a.Y) * dy) / lenSq));
            return pt.GetDistanceTo(new Point2d(a.X + t * dx, a.Y + t * dy));
        }

        private static double DistanciaPontoPoliline(Point2d pt, List<Point2d> poly)
        {
            double dMin = double.MaxValue;
            for (int i = 0; i < poly.Count - 1; i++)
                dMin = Math.Min(dMin, DistanciaPontoSegmento(pt, poly[i], poly[i + 1]));
            return dMin;
        }

        // ── Carregadores de entidades Civil 3D ─────────────────────
        private static List<Parcel> CarregarTodosParcels(CivilDocument civilDoc, Database db, Transaction tr)
        {
            var lista = new List<Parcel>();
            foreach (ObjectId siteId in civilDoc.GetSiteIds())
            {
                Site site = tr.GetObject(siteId, OpenMode.ForRead) as Site;
                if (site == null) continue;
                foreach (ObjectId pid in site.GetParcelIds())
                {
                    Parcel p = tr.GetObject(pid, OpenMode.ForRead) as Parcel;
                    if (p != null) lista.Add(p);
                }
            }
            return lista;
        }

        private static List<AlignmentInfo> CarregarAlignments(CivilDocument civilDoc, Database db, Transaction tr)
        {
            var lista = new List<AlignmentInfo>();
            try
            {
                foreach (ObjectId aid in civilDoc.GetAlignmentIds())
                {
                    Alignment al = tr.GetObject(aid, OpenMode.ForRead) as Alignment;
                    if (al == null) continue;

                    var pontos = new List<Point2d>();
                    // Samplear o alignment em intervalos de 1 metro
                    double comprimento = al.Length;
                    int amostras = Math.Max(10, (int)(comprimento / 1.0));
                    for (int i = 0; i <= amostras; i++)
                    {
                        double station = al.StartingStation + (comprimento / amostras) * i;
                        double x, y;
                        al.PointLocation(station, 0, out x, out y);
                        pontos.Add(new Point2d(x, y));
                    }

                    lista.Add(new AlignmentInfo { Nome = al.Name, Pontos = pontos });
                }
            }
            catch { }
            return lista;
        }

        private static string ExtrairNomeSite(Parcel parcel, Transaction tr)
        {
            try
            {
                Site site = tr.GetObject(parcel.SiteId, OpenMode.ForRead) as Site;
                return site?.Name ?? "";
            }
            catch { return ""; }
        }
    }

    public class AlignmentInfo
    {
        public string Nome { get; set; }
        public List<Point2d> Pontos { get; set; }
    }

    // ================================================================
    // NÚMERO POR EXTENSO
    // ================================================================

    public static class NumeroPorExtenso
    {
        public static string AreaCompleta(double area)
        {
            int metros = (int)Math.Floor(area);
            int cm2 = (int)Math.Round((area - metros) * 100);

            if (metros > 0 && cm2 > 0)
                return $"{Inteiro(metros)} metros quadrados e {Inteiro(cm2)} centímetros quadrados";
            if (metros > 0)
                return $"{Inteiro(metros)} metros quadrados";
            return $"{Inteiro(cm2)} centímetros quadrados";
        }

        public static string Inteiro(int n)
        {
            if (n == 0) return "zero";
            if (n < 0)  return "menos " + Inteiro(-n);

            string[] un  = { "","um","dois","três","quatro","cinco","seis","sete","oito","nove",
                              "dez","onze","doze","treze","quatorze","quinze","dezesseis",
                              "dezessete","dezoito","dezenove" };
            string[] dez = { "","","vinte","trinta","quarenta","cinquenta",
                              "sessenta","setenta","oitenta","noventa" };
            string[] cen = { "","cento","duzentos","trezentos","quatrocentos","quinhentos",
                              "seiscentos","setecentos","oitocentos","novecentos" };

            if (n == 100) return "cem";
            if (n < 20)   return un[n];
            if (n < 100)  return dez[n/10] + (n%10 != 0 ? " e " + un[n%10] : "");
            if (n < 1000)
            {
                string r = cen[n/100];
                return n%100 != 0 ? r + " e " + Inteiro(n%100) : r;
            }
            if (n < 1000000)
            {
                int mil = n / 1000;
                string r = mil == 1 ? "mil" : Inteiro(mil) + " mil";
                return n%1000 != 0 ? r + " e " + Inteiro(n%1000) : r;
            }
            return n.ToString("N0");
        }
    }
}
