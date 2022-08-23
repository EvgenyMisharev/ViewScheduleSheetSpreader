using Autodesk.Revit.DB;
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

namespace ViewScheduleSheetSpreader
{
    public partial class ReadonlyParameterWPF : Window
    {
        public ReadonlyParameterWPF(string errorMassage, List<ElementId> readonlyGruppingParameterElementIdList)
        {
            InitializeComponent();
            textBox_ErrorMassage.Text = errorMassage;
            
            listBox_Ids.ItemsSource = readonlyGruppingParameterElementIdList;
            listBox_Ids.DisplayMemberPath = "";
        }
        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void ReadonlyParameterWPF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                DialogResult = true;
                Close();
            }

            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}
