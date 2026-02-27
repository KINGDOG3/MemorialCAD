using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace MemorialPlugin.WPF
{
    /// <summary>
    /// Janela WPF principal do plugin Memorial Descritivo
    /// </summary>
    public partial class JanelaMemorial : Window
    {
        private List<ObjectId> _parcelIds;
        private Document _acadDoc;
        private List<LoteInfo> _lotes;

        // ViewModel para binding
        public ObservableCollection<LoteViewModel> LotesVM { get; set; }
            = new ObservableCollection<LoteViewModel>();

        public JanelaMemorial(List<ObjectId> parcelIds, Document acadDoc)
        {
            _parcelIds = parcelIds;
            _acadDoc   = acadDoc;
            InitializeComponent();
            DataContext = this;
            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            StatusText.Text = "Processando parcels...";
            try
            {
                _lotes = ParcelProcessor.ProcessarParcels(_parcelIds, _acadDoc);

                foreach (var lote in _lotes)
                {
                    // Classificação automática inicial: usar o lado mais próximo
                    // de um Alignment como Frente (index 0 como fallback)
                    int indexFrente = EncontrarIndexFrenteAutomatico(lote);
                    ParcelProcessor.ClassificarFaces(lote, indexFrente);
                    LotesVM.Add(new LoteViewModel(lote, indexFrente));
                }

                ListaLotes.ItemsSource = LotesVM;
                if (LotesVM.Count > 0)
                    ListaLotes.SelectedIndex = 0;

                StatusText.Text = $"{_lotes.Count} lote(s) carregado(s). Revise as confrontações e clique em Exportar.";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Erro: {ex.Message}";
                MessageBox.Show(ex.ToString(), "Erro ao processar parcels", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int EncontrarIndexFrenteAutomatico(LoteInfo lote)
        {
            // Prefere lado cujo confrontante parece uma rua
            for (int i = 0; i < lote.Lados.Count; i++)
            {
                string c = lote.Lados[i].Confrontante?.ToUpper() ?? "";
                if (c.Contains("RUA") || c.Contains("AV") || c.Contains("ESTR") ||
                    c.Contains("ROD") || c.Contains("TRAV") || c.Contains("ALAM"))
                    return i;
            }
            return 0;
        }

        private void ListaLotes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListaLotes.SelectedItem is LoteViewModel vm)
                PainelDetalhe.DataContext = vm;
        }

        private void ComboFrente_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PainelDetalhe.DataContext is LoteViewModel vm)
            {
                vm.AtualizarClassificacao();
                // Forçar refresh da lista de lados
                ListaLados.ItemsSource = null;
                ListaLados.ItemsSource = vm.LadasVM;
            }
        }

        // ── Exportar ──────────────────────────────────────────────

        private void BtnExportar_Click(object sender, RoutedEventArgs e)
        {
            // Validar pasta destino
            string pasta = TxtPastaDestino.Text.Trim();
            if (string.IsNullOrEmpty(pasta) || !System.IO.Directory.Exists(pasta))
            {
                pasta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                TxtPastaDestino.Text = pasta;
            }

            // Coletar dados de cabeçalho
            var cabecalho = new CabecalhoDoc
            {
                NomeLoteamento    = TxtNomeLoteamento.Text.Trim(),
                Matricula         = TxtMatricula.Text.Trim(),
                Proprietario      = TxtProprietario.Text.Trim(),
                ResponsavelTecnico= TxtRespTecnico.Text.Trim(),
                Crea              = TxtCrea.Text.Trim(),
                Municipio         = TxtMunicipio.Text.Trim(),
                Estado            = TxtEstado.Text.Trim(),
                Data              = DateTime.Now
            };

            // Aplicar classificações finais
            foreach (var vm in LotesVM)
            {
                vm.AtualizarClassificacao();
            }
            var lotes = LotesVM.Select(v => v.Lote).ToList();

            try
            {
                StatusText.Text = "Gerando arquivos...";

                bool gerarDocx  = ChkDocx.IsChecked  == true;
                bool gerarXlsx  = ChkXlsx.IsChecked  == true;

                if (gerarDocx)
                {
                    string nomeDocx = $"Memorial_{SanitizarNome(cabecalho.NomeLoteamento)}_{DateTime.Now:yyyyMMdd_HHmm}.docx";
                    string pathDocx = System.IO.Path.Combine(pasta, nomeDocx);
                    GeradorDocx.Gerar(lotes, cabecalho, pathDocx);
                    StatusText.Text += $"\n✓ DOCX: {pathDocx}";
                }

                if (gerarXlsx)
                {
                    string nomeXlsx = $"Memorial_{SanitizarNome(cabecalho.NomeLoteamento)}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
                    string pathXlsx = System.IO.Path.Combine(pasta, nomeXlsx);
                    GeradorXlsx.Gerar(lotes, cabecalho, pathXlsx);
                    StatusText.Text += $"\n✓ XLSX: {pathXlsx}";
                }

                MessageBox.Show(
                    $"Arquivos gerados em:\n{pasta}",
                    "Memorial Descritivo — Exportação concluída",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao exportar:\n{ex.Message}", "Erro",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = $"Erro: {ex.Message}";
            }
        }

        private void BtnPastaDestino_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Selecione a pasta de destino dos arquivos"
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                TxtPastaDestino.Text = dlg.SelectedPath;
        }

        private void BtnFechar_Click(object sender, RoutedEventArgs e) => Close();

        private string SanitizarNome(string n)
        {
            if (string.IsNullOrEmpty(n)) return "Memorial";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                n = n.Replace(c, '_');
            return n.Replace(' ', '_');
        }
    }

    // ================================================================
    // VIEW MODELS
    // ================================================================

    public class LoteViewModel : System.ComponentModel.INotifyPropertyChanged
    {
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        private void OnProp(string p) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(p));

        public LoteInfo Lote { get; }
        public ObservableCollection<LadoViewModel> LadasVM { get; } = new ObservableCollection<LadoViewModel>();

        private int _indexFrente;
        public int IndexFrente
        {
            get => _indexFrente;
            set { _indexFrente = value; OnProp(nameof(IndexFrente)); AtualizarClassificacao(); }
        }

        // Lista de opções para o ComboBox de frente
        public List<string> OpcoesFrente { get; }

        public string NomeLote => Lote.Nome;
        public string Quadra   => Lote.Quadra;

        public LoteViewModel(LoteInfo lote, int indexFrente)
        {
            Lote = lote;
            _indexFrente = indexFrente;

            OpcoesFrente = lote.Lados.Select((l, i) =>
                $"[L{i+1}] {l.Comprimento:F3}m — {l.Confrontante}").ToList();

            foreach (var lado in lote.Lados)
                LadasVM.Add(new LadoViewModel(lado));
        }

        public void AtualizarClassificacao()
        {
            ParcelProcessor.ClassificarFaces(Lote, _indexFrente);
            foreach (var vm in LadasVM)
                vm.RefreshClassificacao();
        }
    }

    public class LadoViewModel : System.ComponentModel.INotifyPropertyChanged
    {
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        private void OnProp(string p) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(p));

        private LadoInfo _lado;
        public LadoViewModel(LadoInfo lado) { _lado = lado; }

        public string Classificacao  => _lado.ClassificacaoFace ?? "—";
        public string Confrontante
        {
            get => _lado.Confrontante;
            set { _lado.Confrontante = value; OnProp(nameof(Confrontante)); }
        }
        public string Comprimento   => $"{_lado.Comprimento:F3} m";
        public string Azimute       => _lado.AzimuteFormatado;
        public string Vertice       => $"V{_lado.De}→V{_lado.Para}";

        public void RefreshClassificacao() => OnProp(nameof(Classificacao));
    }

    // ================================================================
    // CABEÇALHO
    // ================================================================

    public class CabecalhoDoc
    {
        public string NomeLoteamento     { get; set; }
        public string Matricula          { get; set; }
        public string Proprietario       { get; set; }
        public string ResponsavelTecnico { get; set; }
        public string Crea               { get; set; }
        public string Municipio          { get; set; }
        public string Estado             { get; set; }
        public DateTime Data             { get; set; }
    }
}
