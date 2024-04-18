using PlannerOpenXML.ViewModel;
using System.Windows;

namespace PlannerOpenXML
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }
    }
}