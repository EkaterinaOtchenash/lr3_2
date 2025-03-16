using lr3.Classes;
using System.Windows;


namespace lr3
{
    public partial class MainWindow : Window
    {
        private APIInteraction apiInteraction;

        public MainWindow()
        {
            InitializeComponent();
            apiInteraction = new APIInteraction();
        }

        private void ButtonGetFullName_Click(object sender, RoutedEventArgs e)
        {
            TextBlockFullName.Text = apiInteraction.GetFullName();
        }

        private void ButtonSendResult_Click(object sender, RoutedEventArgs e)
        {
            TextBlockResult.Text = apiInteraction.FillDocument();
        }
    }
}