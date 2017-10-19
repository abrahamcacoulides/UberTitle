using System;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using Microsoft.Win32;
using Engine.EventArgs;
using Engine.ViewModels;

namespace WpfApp2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Session _session;

        public MainWindow()
        {
            InitializeComponent();

            _session = new Session();

            _session.OnMessageRaised += OnMessageRaised;

            _session.RefreshVersions();

            DataContext = _session;

        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "dwg files (*.dwg)|*.dwg";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Please Select Source File(s) to change Title";
            if (openFileDialog.ShowDialog() == true)
                txtEditor.Text = String.Empty;
                _session.RefreshList(openFileDialog.FileNames.ToList());
                txtEditor.Text = _session.SelectedDWGS();
        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            _session.GoButton();
        }

        private void OnMessageRaised(object sender, MessageEventArgs e)
        {
            Messages.Document.Blocks.Add(new Paragraph(new Run(e.Message)));
            Messages.ScrollToEnd();
        }
    }
}
