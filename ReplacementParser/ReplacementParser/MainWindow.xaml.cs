using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Newtonsoft.Json;
using xNet;
namespace ReplacementParser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
                
        }
        public void SendReplacement()
        {
            
            Replacement CHTOTIBReplacement = new Replacement(AppDomain.CurrentDomain.BaseDirectory + @"\main.docx");
            CHTOTIBReplacement.GetInfo();
            string replacementJson = JsonConvert.SerializeObject(CHTOTIBReplacement.ResultReplacements);
            using (var request = new HttpRequest())
            {
                var reqParams = new RequestParams();
                reqParams["count"] = CHTOTIBReplacement.ResultReplacements.Count;
                reqParams["replacement"] = replacementJson;
                string content = request.Post("localhost/replacement/add", reqParams).ToString();
            }
            MessageBox.Show("Отправлено");

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendReplacement();
        }
    }
}
