using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Navigation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfAppSmetaGraf.View
{
    /// <summary>
    /// Логика взаимодействия для WindowFirst.xaml
    /// </summary>
    public partial class WindowFirst : NavigationWindow
    {
        public WindowFirst()
        {
            InitializeComponent();
            this.NavigationService.Navigate(new PageMain());
        }

      
    }
}
