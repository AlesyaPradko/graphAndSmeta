using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfAppSmetaGraf.View
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class PageMain : Page
    {
        public PageMain()
        {
            InitializeComponent();
        }
        

        private void ToggleButton_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ToggleButton toggle=sender as ToggleButton;
            switch(toggle.Name)
            {
                case "ButtonTE": this.NavigationService.Navigate(new PageTE());break;
                case "ButtonGraph": this.NavigationService.Navigate(new PageGraph());break;
            }
            
        }

    }
}
