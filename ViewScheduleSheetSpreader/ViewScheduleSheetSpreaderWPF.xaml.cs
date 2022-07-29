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
        public ObservableCollection<ViewSchedule> SelectedViewScheduleCollection;
        ObservableCollection<Family> TitleBlocksForFirstSheetCollection;
        ObservableCollection<Family> TitleBlocksForFollowingSheetsCollection;
        ObservableCollection<FamilySymbol> TitleBlocksForFirstSheetTypeCollection;
        ObservableCollection<FamilySymbol> TitleBlocksForFollowingSheetsTypeCollection;
        ObservableCollection<Parameter> FamilyInstanceParametersCollection;
        ObservableCollection<Definition> ParamDefinitionsCollection;
        public FamilySymbol FirstSheetType;
        public FamilySymbol FollowingSheetType;
        public Parameter SheetFormatParameter;
        public Definition GroupingParameterDefinition;
        public string SheetSizeVariantName;
        public double XOffset;
        public double YOffset;
        public int FirstSheetNumber;

        public ViewScheduleSheetSpreaderWPF(Document doc, List<ViewSchedule> viewScheduleList, List<Family> titleBlockFamilysList, List<Definition> paramDefinitionsList)
        {
            Doc = doc;
            ViewScheduleInProjectCollection = new ObservableCollection<ViewSchedule>(viewScheduleList);
            SelectedViewScheduleCollection = new ObservableCollection<ViewSchedule>();
            TitleBlocksForFirstSheetCollection = new ObservableCollection<Family>(titleBlockFamilysList);
            TitleBlocksForFollowingSheetsCollection = new ObservableCollection<Family>(titleBlockFamilysList);
            ParamDefinitionsCollection = new ObservableCollection<Definition>(paramDefinitionsList.OrderBy(df =>df.Name));

            InitializeComponent();

            listBox_ViewScheduleInProjectCollection.ItemsSource = ViewScheduleInProjectCollection;
            listBox_ViewScheduleInProjectCollection.DisplayMemberPath = "Name";

            listBox_SelectedViewScheduleCollection.ItemsSource = SelectedViewScheduleCollection;
            listBox_SelectedViewScheduleCollection.DisplayMemberPath = "Name";

            comboBox_FirstSheetFamily.ItemsSource = TitleBlocksForFirstSheetCollection;
            comboBox_FirstSheetFamily.DisplayMemberPath = "Name";
            comboBox_FirstSheetFamily.SelectedItem = comboBox_FirstSheetFamily.Items.GetItemAt(0);

            comboBox_FollowingSheetsFamily.ItemsSource = TitleBlocksForFollowingSheetsCollection;
            comboBox_FollowingSheetsFamily.DisplayMemberPath = "Name";
            comboBox_FollowingSheetsFamily.SelectedItem = comboBox_FollowingSheetsFamily.Items.GetItemAt(0);

            comboBox_GroupingParameter.ItemsSource = ParamDefinitionsCollection;
            comboBox_GroupingParameter.DisplayMemberPath = "Name";
            comboBox_GroupingParameter.SelectedItem = comboBox_GroupingParameter.Items.GetItemAt(0);
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

        private void comboBox_FirstSheetFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<ElementId> familySymbolsIdList = ((sender as ComboBox).SelectedItem as Family).GetFamilySymbolIds().ToList();
            TitleBlocksForFirstSheetTypeCollection = new ObservableCollection<FamilySymbol>();
            if (familySymbolsIdList.Count != 0)
            {
                foreach (ElementId symbolId in familySymbolsIdList)
                {
                    TitleBlocksForFirstSheetTypeCollection.Add(Doc.GetElement(symbolId) as FamilySymbol);
                }
            }
            TitleBlocksForFirstSheetTypeCollection = new ObservableCollection<FamilySymbol>(TitleBlocksForFirstSheetTypeCollection.OrderBy(fs => fs.Name).ToList());
            comboBox_FirstSheetType.ItemsSource = TitleBlocksForFirstSheetTypeCollection;
            comboBox_FirstSheetType.DisplayMemberPath = "Name";
            comboBox_FirstSheetType.SelectedItem = comboBox_FirstSheetType.Items.GetItemAt(0);
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

        private void comboBox_FirstSheetType_SelectionChanged(object sender, SelectionChangedEventArgs e)
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
            }
            else if (selectedButtonName == "radioButton_Type")
            {
                label_SheetFormatParameter.IsEnabled = false;
                comboBox_SheetFormatParameter.IsEnabled = false;
            }
        }
        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            FirstSheetType = comboBox_FirstSheetType.SelectedItem as FamilySymbol;
            FollowingSheetType = comboBox_FollowingSheetsType.SelectedItem as FamilySymbol;
            SheetFormatParameter = comboBox_SheetFormatParameter.SelectedItem as Parameter;
            GroupingParameterDefinition = comboBox_GroupingParameter.SelectedItem as Definition;
            double.TryParse(textBox_XOffset.Text, out XOffset);
            double.TryParse(textBox_YOffset.Text, out YOffset);
            int.TryParse(textBox_FirstSheetNumber.Text, out FirstSheetNumber);
            SheetSizeVariantName = (this.groupBox_SheetSize.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            DialogResult = true;
            Close();
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
                FirstSheetType = comboBox_FirstSheetType.SelectedItem as FamilySymbol;
                FollowingSheetType = comboBox_FollowingSheetsType.SelectedItem as FamilySymbol;
                SheetFormatParameter = comboBox_SheetFormatParameter.SelectedItem as Parameter;
                GroupingParameterDefinition = comboBox_GroupingParameter.SelectedItem as Definition;
                double.TryParse(textBox_XOffset.Text, out XOffset);
                double.TryParse(textBox_YOffset.Text, out YOffset);
                int.TryParse(textBox_FirstSheetNumber.Text, out FirstSheetNumber);
                SheetSizeVariantName = (this.groupBox_SheetSize.Content as System.Windows.Controls.Grid)
                    .Children.OfType<RadioButton>()
                    .FirstOrDefault(rb => rb.IsChecked.Value == true)
                    .Name;
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
