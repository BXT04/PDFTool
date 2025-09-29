using Microsoft.Win32;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace PDFTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private enum AppView { Dashboard, Merge, Split, Compress }
        private enum SplitOption { Individual, Range }
        private enum CompressOption { High, Medium, Low }

        public ObservableCollection<string> FilesToMerge { get; set; } = new ObservableCollection<string>();

        private string _fileToSplit;
        private string _fileToCompress;
        private string _splitPageRange = "";

        public string SplitPageRange
        {
            get => _splitPageRange;
            set
            {
                _splitPageRange = value;
                OnPropertyChanged(nameof(SplitPageRange));
            }
        }

        public string FileToCompress
        {
            get => _fileToCompress;
            set
            {
                _fileToCompress = value;
                OnPropertyChanged(nameof(FileToCompress));
                OnPropertyChanged(nameof(FileToCompressDisplayName));
            }
        }

        // Properties to display just the filename in the UI
        public string FileToSplitDisplayName => string.IsNullOrEmpty(FileToSplit) 
            ? "Click or drag and drop a PDF file here" 
            : Path.GetFileName(FileToSplit);
            
        public string FileToCompressDisplayName => string.IsNullOrEmpty(FileToCompress) 
            ? "Click or drag and drop a PDF file here" 
            : Path.GetFileName(FileToCompress);

        public string FileToSplit
        {
            get => _fileToSplit;
            set
            {
                _fileToSplit = value;
                OnPropertyChanged(nameof(FileToSplit));
                OnPropertyChanged(nameof(FileToSplitDisplayName));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void SwitchView(AppView view)
        {
            DashboardView.Visibility = Visibility.Collapsed;
            MergeView.Visibility = Visibility.Collapsed;
            SplitView.Visibility = Visibility.Collapsed;
            CompressView.Visibility = Visibility.Collapsed;

            switch (view)
            {
                case AppView.Dashboard:
                    DashboardView.Visibility = Visibility.Visible;
                    break;
                case AppView.Merge:
                    MergeView.Visibility = Visibility.Visible;
                    break;
                case AppView.Split:
                    SplitView.Visibility = Visibility.Visible;
                    break;
                case AppView.Compress:
                    CompressView.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void GoToDashboard_Click(object sender, RoutedEventArgs e) => SwitchView(AppView.Dashboard);
        private void GoToMerge_Click(object sender, RoutedEventArgs e) => SwitchView(AppView.Merge);
        private void GoToSplit_Click(object sender, RoutedEventArgs e) => SwitchView(AppView.Split);
        private void GoToCompress_Click(object sender, RoutedEventArgs e) => SwitchView(AppView.Compress);

        // ═══════════════════════════════════════════════════════════
        // MERGE FUNCTIONS
        // ═══════════════════════════════════════════════════════════

        private void SelectMergeFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Chọn tệp PDF để ghép"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filename in openFileDialog.FileNames)
                {
                    if (!FilesToMerge.Contains(filename))
                    {
                        FilesToMerge.Add(filename);
                    }
                }
            }
        }

        private void MergeFile_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                var pdfFiles = files.Where(f => Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase));
                foreach (string file in pdfFiles)
                {
                    if (!FilesToMerge.Contains(file))
                    {
                        FilesToMerge.Add(file);
                    }
                }
            }
        }

        private void ClearMergeList_Click(object sender, RoutedEventArgs e)
        {
            FilesToMerge.Clear();
        }

        private void MergeFiles_Click(object sender, RoutedEventArgs e)
        {
            if (FilesToMerge.Count < 2)
            {
                MessageBox.Show("Vui lòng chọn ít nhất 2 tệp PDF để hợp nhất.", 
                    "Không đủ tập tin", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF File (*.pdf)|*.pdf",
                Title = "Lưu tệp PDF đã hợp nhất",
                FileName = "merged.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (PdfDocument outputDocument = new PdfDocument())
                    {
                        foreach (string file in FilesToMerge)
                        {
                            using (PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                            {
                                for (int i = 0; i < inputDocument.PageCount; i++)
                                {
                                    outputDocument.AddPage(inputDocument.Pages[i]);
                                }
                            }
                        }
                        outputDocument.Save(saveFileDialog.FileName);
                    }
                    MessageBox.Show($"Đã hợp nhất thành công {FilesToMerge.Count} file!", 
                        "Hợp nhất file thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                    Process.Start(new ProcessStartInfo(saveFileDialog.FileName) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Đã xảy ra lỗi khi hợp nhất các tệp PDF: {ex.Message}", 
                        "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // SPLIT FUNCTIONS
        // ═══════════════════════════════════════════════════════════

        private void SelectSplitFile_MouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog 
            { 
                Filter = "PDF Files (*.pdf)|*.pdf", 
                Title = "Select PDF File to Split" 
            };
            if (openFileDialog.ShowDialog() == true)
            {
                FileToSplit = openFileDialog.FileName;
            }
        }

        private void SplitFile_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                var firstPdf = files.FirstOrDefault(f => Path.GetExtension(f)
                    .Equals(".pdf", StringComparison.OrdinalIgnoreCase));
                if (firstPdf != null)
                {
                    FileToSplit = firstPdf;
                }
            }
        }

        private void ClearSplitFile_Click(object sender, RoutedEventArgs e)
        {
            FileToSplit = null;
            SplitPageRange = "";
        }

        private void SplitFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FileToSplit))
            {
                MessageBox.Show("Please select a PDF file to split.", 
                    "No File Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SplitOption selectedOption = (SplitIndividualPagesRadio.IsChecked == true) 
                ? SplitOption.Individual 
                : SplitOption.Range;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF File (*.pdf)|*.pdf",
                Title = "Select Output Location",
                FileName = $"{Path.GetFileNameWithoutExtension(FileToSplit)}_split.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputDirectory = Path.GetDirectoryName(saveFileDialog.FileName);
                string baseFileName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);

                try
                {
                    using (PdfDocument inputDocument = PdfReader.Open(FileToSplit, PdfDocumentOpenMode.Import))
                    {
                        if (selectedOption == SplitOption.Individual)
                        {
                            for (int i = 0; i < inputDocument.PageCount; i++)
                            {
                                using (PdfDocument outputDocument = new PdfDocument())
                                {
                                    outputDocument.AddPage(inputDocument.Pages[i]);
                                    string outputFileName = Path.Combine(outputDirectory, 
                                        $"{baseFileName}_page_{i + 1}.pdf");
                                    outputDocument.Save(outputFileName);
                                }
                            }
                            MessageBox.Show($"Successfully split the file into {inputDocument.PageCount} pages.", 
                                "Split Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else // Split by range
                        {
                            List<int> pagesToExtract = ParsePageRange(SplitPageRange, inputDocument.PageCount);
                            if (pagesToExtract == null || pagesToExtract.Count == 0)
                            {
                                MessageBox.Show("Invalid page range specified. Please use formats like '1-3', '5', '1,3,5-7'.", 
                                    "Invalid Range", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }

                            using (PdfDocument outputDocument = new PdfDocument())
                            {
                                foreach (int pageIndex in pagesToExtract)
                                {
                                    outputDocument.AddPage(inputDocument.Pages[pageIndex]);
                                }
                                outputDocument.Save(saveFileDialog.FileName);
                            }
                            MessageBox.Show($"Successfully extracted {pagesToExtract.Count} pages.", 
                                "Split Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                        }

                        Process.Start(new ProcessStartInfo(outputDirectory) { UseShellExecute = true });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred during splitting: {ex.Message}", 
                        "Split Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private List<int> ParsePageRange(string range, int maxPage)
        {
            var pageNumbers = new SortedSet<int>();
            if (string.IsNullOrWhiteSpace(range)) return null;

            try
            {
                foreach (var part in range.Split(','))
                {
                    var trimmedPart = part.Trim();
                    if (trimmedPart.Contains('-'))
                    {
                        var limits = trimmedPart.Split('-');
                        int start = int.Parse(limits[0].Trim());
                        int end = int.Parse(limits[1].Trim());
                        for (int i = start; i <= end; i++)
                        {
                            if (i > 0 && i <= maxPage) pageNumbers.Add(i - 1);
                        }
                    }
                    else
                    {
                        int page = int.Parse(trimmedPart);
                        if (page > 0 && page <= maxPage) pageNumbers.Add(page - 1);
                    }
                }
            }
            catch { return null; }
            return pageNumbers.ToList();
        }

        // ═══════════════════════════════════════════════════════════
        // COMPRESS FUNCTIONS
        // ═══════════════════════════════════════════════════════════

        private void SelectCompressFile_MouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog 
            { 
                Filter = "PDF Files (*.pdf)|*.pdf", 
                Title = "Select PDF File to Compress" 
            };
            if (openFileDialog.ShowDialog() == true)
            {
                FileToCompress = openFileDialog.FileName;
            }
        }

        private void CompressFile_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                var firstPdf = files.FirstOrDefault(f => Path.GetExtension(f)
                    .Equals(".pdf", StringComparison.OrdinalIgnoreCase));
                if (firstPdf != null)
                {
                    FileToCompress = firstPdf;
                }
            }
        }

        private void ClearCompressFile_Click(object sender, RoutedEventArgs e)
        {
            FileToCompress = null;
        }

        private void CompressFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FileToCompress))
            {
                MessageBox.Show("Please select a PDF file to compress.", 
                    "No File Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF File (*.pdf)|*.pdf",
                Title = "Save Compressed PDF",
                FileName = $"{Path.GetFileNameWithoutExtension(FileToCompress)}_compressed.pdf"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    long originalSize = new FileInfo(FileToCompress).Length;

                    using (PdfDocument document = PdfReader.Open(FileToCompress, PdfDocumentOpenMode.Modify))
                    {
                        var options = document.Options;
                        options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
                        options.UseFlateDecoderForJpegImages = PdfUseFlateDecoderForJpegImages.Always;
                        options.CompressContentStreams = true;
                        options.NoCompression = false;
                        options.EnableCcittCompressionForBilevelImages = true;

                        if (ConvertToGrayscaleCheck.IsChecked == true)
                        {
                            MessageBox.Show("Grayscale conversion is not yet implemented. " +
                                "This feature requires a more advanced PDF library.", 
                                "Feature Not Available", MessageBoxButton.OK, MessageBoxImage.Information);
                        }

                        document.Save(saveFileDialog.FileName);
                    }

                    long newSize = new FileInfo(saveFileDialog.FileName).Length;
                    double reduction = 100 - ((double)newSize / originalSize * 100);
                    
                    string message = reduction > 0 
                        ? $"Compression complete! File size reduced by {reduction:F2}%."
                        : "Compression complete! Note: This PDF may already be optimized, so size reduction is minimal.";
                    
                    MessageBox.Show(message, "Compression Success", 
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    Process.Start(new ProcessStartInfo(saveFileDialog.FileName) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred during compression: {ex.Message}", 
                        "Compression Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // SHARED FUNCTIONS
        // ═══════════════════════════════════════════════════════════

        private void FileDrop_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }
    }
}