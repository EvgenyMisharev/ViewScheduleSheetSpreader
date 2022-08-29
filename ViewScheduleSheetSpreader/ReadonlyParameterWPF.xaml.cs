using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

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
