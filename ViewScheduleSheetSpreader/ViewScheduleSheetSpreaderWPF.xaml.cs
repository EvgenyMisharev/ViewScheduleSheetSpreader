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
using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;

namespace ViewScheduleSheetSpreader
{
    public partial class ViewScheduleSheetSpreaderWPF : Window
    {
        Document Doc;
        ObservableCollection<ViewSchedule> ViewScheduleInProjectCollection;
        ObservableCollection<ViewSchedule> SelectedViewScheduleCollection;
        ObservableCollection<Family> TitleBlocksFor1stSheetCollection;
        ObservableCollection<Family> TitleBlocksForFollowingSheetsCollection;
        ObservableCollection<FamilySymbol> TitleBlocksFor1stSheetTypeCollection;
        ObservableCollection<FamilySymbol> TitleBlocksForFollowingSheetsTypeCollection;
        ObservableCollection<Parameter> FamilyInstanceParametersCollection;
        public ViewScheduleSheetSpreaderWPF(Document doc, List<ViewSchedule> viewScheduleList, List<Family> titleBlockFamilysList)
        {
            Doc = doc;
            ViewScheduleInProjectCollection = new ObservableCollection<ViewSchedule>(viewScheduleList);
            SelectedViewScheduleCollection = new ObservableCollection<ViewSchedule>();
            TitleBlocksFor1stSheetCollection = new ObservableCollection<Family>(titleBlockFamilysList);
            TitleBlocksForFollowingSheetsCollection = new ObservableCollection<Family>(titleBlockFamilysList);

            InitializeComponent();

            listBox_ViewScheduleInProjectCollection.ItemsSource = ViewScheduleInProjectCollection;
            listBox_ViewScheduleInProjectCollection.DisplayMemberPath = "Name";

            listBox_SelectedViewScheduleCollection.ItemsSource = SelectedViewScheduleCollection;
            listBox_SelectedViewScheduleCollection.DisplayMemberPath = "Name";

            comboBox_1stSheetFamily.ItemsSource = TitleBlocksFor1stSheetCollection;
            comboBox_1stSheetFamily.DisplayMemberPath = "Name";
            comboBox_1stSheetFamily.SelectedItem = comboBox_1stSheetFamily.Items.GetItemAt(0);

            comboBox_FollowingSheetsFamily.ItemsSource = TitleBlocksForFollowingSheetsCollection;
            comboBox_FollowingSheetsFamily.DisplayMemberPath = "Name";
            comboBox_FollowingSheetsFamily.SelectedItem = comboBox_FollowingSheetsFamily.Items.GetItemAt(0);
        }
        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
        private void btn_Add_Click(object sender, RoutedEventArgs e)
        {
            List<ViewSchedule> viewScheduleInProjectCollectionSelectedItems = listBox_ViewScheduleInProjectCollection
                .SelectedItems
                .Cast<ViewSchedule>()
                .ToList();
            foreach(ViewSchedule vs in viewScheduleInProjectCollectionSelectedItems)
            {
                ViewScheduleInProjectCollection.Remove(vs);
                SelectedViewScheduleCollection.Add(vs);
            }
        }

        private void btn_Exclude_Click(object sender, RoutedEventArgs e)
        {
            List<ViewSchedule> selectedViewScheduleCollectionSelectedItems = listBox_SelectedViewScheduleCollection
                .SelectedItems
                .Cast<ViewSchedule>()
                .ToList();
            foreach (ViewSchedule vs in selectedViewScheduleCollectionSelectedItems)
            {
                SelectedViewScheduleCollection.Remove(vs);
                ViewScheduleInProjectCollection.Add(vs);
            }
            ViewScheduleInProjectCollection = new ObservableCollection<ViewSchedule>(ViewScheduleInProjectCollection.OrderBy(vs => vs.Name, new AlphanumComparatorFastString()));
            listBox_ViewScheduleInProjectCollection.ItemsSource = ViewScheduleInProjectCollection;
        }

        private void btn_MoveUp_Click(object sender, RoutedEventArgs e)
        {
            List<ViewSchedule> selectedViewScheduleCollectionSelectedItems = listBox_SelectedViewScheduleCollection
                .SelectedItems
                .Cast<ViewSchedule>()
                .ToList();
            foreach (ViewSchedule vs in selectedViewScheduleCollectionSelectedItems)
            {
                if(SelectedViewScheduleCollection.IndexOf(vs) != 0)
                {
                    int oldIndex = SelectedViewScheduleCollection.IndexOf(vs);
                    int newIndex = SelectedViewScheduleCollection.IndexOf(vs) - 1;
                    SelectedViewScheduleCollection.Move(oldIndex, newIndex);
                }
            }
                
        }
        private void btn_MoveDown_Click(object sender, RoutedEventArgs e)
        {
            List<ViewSchedule> selectedViewScheduleCollectionSelectedItems = listBox_SelectedViewScheduleCollection
                .SelectedItems
                .Cast<ViewSchedule>()
                .ToList();
            for (int i = selectedViewScheduleCollectionSelectedItems.Count - 1; i >= 0; i--)
            {
                if (SelectedViewScheduleCollection.IndexOf(selectedViewScheduleCollectionSelectedItems[i]) != SelectedViewScheduleCollection.Count - 1)
                {
                    int oldIndex = SelectedViewScheduleCollection.IndexOf(selectedViewScheduleCollectionSelectedItems[i]);
                    int newIndex = SelectedViewScheduleCollection.IndexOf(selectedViewScheduleCollectionSelectedItems[i]) + 1;
                    SelectedViewScheduleCollection.Move(oldIndex, newIndex);
                }
            }
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        private void ViewScheduleSheetSpreaderWPF_KeyDown(object sender, KeyEventArgs e)
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

        private void comboBox_1stSheetFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<ElementId> familySymbolsIdList = ((sender as ComboBox).SelectedItem as Family).GetFamilySymbolIds().ToList();
            TitleBlocksFor1stSheetTypeCollection = new ObservableCollection<FamilySymbol>();
            if (familySymbolsIdList.Count != 0)
            {
                foreach (ElementId symbolId in familySymbolsIdList)
                {
                    TitleBlocksFor1stSheetTypeCollection.Add(Doc.GetElement(symbolId) as FamilySymbol);
                }
            }
            TitleBlocksFor1stSheetTypeCollection = new ObservableCollection<FamilySymbol>(TitleBlocksFor1stSheetTypeCollection.OrderBy(fs => fs.Name).ToList());
            comboBox_1stSheetType.ItemsSource = TitleBlocksFor1stSheetTypeCollection;
            comboBox_1stSheetType.DisplayMemberPath = "Name";
            comboBox_1stSheetType.SelectedItem = comboBox_1stSheetType.Items.GetItemAt(0);
        }

        private void comboBox_FollowingSheetsFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<ElementId> familySymbolsIdList = ((sender as ComboBox).SelectedItem as Family).GetFamilySymbolIds().ToList();
            TitleBlocksForFollowingSheetsTypeCollection = new ObservableCollection<FamilySymbol>();
            if (familySymbolsIdList.Count != 0)
            {
                foreach (ElementId symbolId in familySymbolsIdList)
                {
                    TitleBlocksForFollowingSheetsTypeCollection.Add(Doc.GetElement(symbolId) as FamilySymbol);
                }
            }
            TitleBlocksForFollowingSheetsTypeCollection = new ObservableCollection<FamilySymbol>(TitleBlocksForFollowingSheetsTypeCollection.OrderBy(fs => fs.Name).ToList());
            comboBox_FollowingSheetsType.ItemsSource = TitleBlocksForFollowingSheetsTypeCollection;
            comboBox_FollowingSheetsType.DisplayMemberPath = "Name";
            comboBox_FollowingSheetsType.SelectedItem = comboBox_FollowingSheetsType.Items.GetItemAt(0);
        }

        private void comboBox_1stSheetType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FamilySymbol selectedFamilySymbol = (sender as ComboBox).SelectedItem as FamilySymbol;
            ParameterSet parameterSet = null;
            if (selectedFamilySymbol != null)
            {
                ElementClassFilter filter = new ElementClassFilter(typeof(FamilyInstance));
                IList<ElementId> selectedFamilyInstanceList = selectedFamilySymbol.GetDependentElements(filter);
                if(selectedFamilyInstanceList.Count != 0)
                {
                    parameterSet = Doc.GetElement(selectedFamilyInstanceList.First()).Parameters;

                    List<Parameter> tmpParametersList = new List<Parameter>();
                    foreach (Parameter param in parameterSet)
                    {
                        tmpParametersList.Add(param);
                    }
                    FamilyInstanceParametersCollection = new ObservableCollection<Parameter>(tmpParametersList.OrderBy(p => p.Definition.Name).ToList());
                    comboBox_SheetFormatParameter.ItemsSource = FamilyInstanceParametersCollection;
                    comboBox_SheetFormatParameter.DisplayMemberPath = "Definition.Name";
                    comboBox_SheetFormatParameter.SelectedItem = comboBox_SheetFormatParameter.Items.GetItemAt(0);
                }
            }
        }

        private void radioButton_Checked(object sender, RoutedEventArgs e)
        {
            string selectedButtonName = (this.groupBox_SheetSize.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            if (selectedButtonName == "radioButton_Instance")
            {
                label_SheetFormatParameter.IsEnabled = true;
                comboBox_SheetFormatParameter.IsEnabled = true;
                label_SheetFormatParameterValue.IsEnabled = true;
                textBox_SheetFormatParameterValue.IsEnabled = true;
            }
            else if (selectedButtonName == "radioButton_Type")
            {
                label_SheetFormatParameter.IsEnabled = false;
                comboBox_SheetFormatParameter.IsEnabled = false;
                label_SheetFormatParameterValue.IsEnabled = false;
                textBox_SheetFormatParameterValue.IsEnabled = false;
            }
        }
    }
}
