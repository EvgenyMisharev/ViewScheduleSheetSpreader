using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.ObjectModel;
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
        public string HeaderInSpecificationHeaderVariantName;
        public double SpecificationHeaderHeight;

        ViewScheduleSheetSpreaderSettings ViewScheduleSheetSpreaderSettingsItem;

        public ViewScheduleSheetSpreaderWPF(Document doc, List<ViewSchedule> viewScheduleList, List<Family> titleBlockFamilysList, List<Definition> paramDefinitionsList)
        {
            Doc = doc;
            ViewScheduleInProjectCollection = new ObservableCollection<ViewSchedule>(viewScheduleList);
            SelectedViewScheduleCollection = new ObservableCollection<ViewSchedule>();
            TitleBlocksForFirstSheetCollection = new ObservableCollection<Family>(titleBlockFamilysList);
            TitleBlocksForFollowingSheetsCollection = new ObservableCollection<Family>(titleBlockFamilysList);
            ParamDefinitionsCollection = new ObservableCollection<Definition>(paramDefinitionsList.OrderBy(df => df.Name));

            ViewScheduleSheetSpreaderSettingsItem = new ViewScheduleSheetSpreaderSettings().GetSettings();

            InitializeComponent();

            listBox_ViewScheduleInProjectCollection.ItemsSource = ViewScheduleInProjectCollection;
            listBox_ViewScheduleInProjectCollection.DisplayMemberPath = "Name";

            listBox_SelectedViewScheduleCollection.ItemsSource = SelectedViewScheduleCollection;
            listBox_SelectedViewScheduleCollection.DisplayMemberPath = "Name";

            comboBox_FirstSheetFamily.ItemsSource = TitleBlocksForFirstSheetCollection;
            comboBox_FirstSheetFamily.DisplayMemberPath = "Name";

            comboBox_FollowingSheetsFamily.ItemsSource = TitleBlocksForFollowingSheetsCollection;
            comboBox_FollowingSheetsFamily.DisplayMemberPath = "Name";

            comboBox_GroupingParameter.ItemsSource = ParamDefinitionsCollection;
            comboBox_GroupingParameter.DisplayMemberPath = "Name";

            SetSavedSettingsValueToForm();
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
        private void radioButton_HeaderInSpecificationHeader_Checked(object sender, RoutedEventArgs e)
        {
            string selectedButtonName = (this.groupBox_HeaderInSpecificationHeader.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            if (selectedButtonName == "radioButton_Yes")
            {
                label_SpecificationHeaderHeight.IsEnabled = false;
                textBox_SpecificationHeaderHeight.IsEnabled = false;
            }
            else if (selectedButtonName == "radioButton_No")
            {
                label_SpecificationHeaderHeight.IsEnabled = true;
                textBox_SpecificationHeaderHeight.IsEnabled = true;
            }
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            SaveDialogResultValues();
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
                SaveDialogResultValues();
                DialogResult = true;
                Close();
            }

            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
        private void SaveDialogResultValues()
        {
            ViewScheduleSheetSpreaderSettingsItem = new ViewScheduleSheetSpreaderSettings();
            ViewScheduleSheetSpreaderSettingsItem.FirstSheetFamilyName = (comboBox_FirstSheetFamily.SelectedItem as Family).Name;

            FirstSheetType = comboBox_FirstSheetType.SelectedItem as FamilySymbol;
            ViewScheduleSheetSpreaderSettingsItem.FirstSheetTypeName = FirstSheetType.Name;

            ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsFamilyName = (comboBox_FollowingSheetsFamily.SelectedItem as Family).Name;

            FollowingSheetType = comboBox_FollowingSheetsType.SelectedItem as FamilySymbol;
            ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsTypeName = FollowingSheetType.Name;

            SheetFormatParameter = comboBox_SheetFormatParameter.SelectedItem as Parameter;
            ViewScheduleSheetSpreaderSettingsItem.SheetFormatParameterName = SheetFormatParameter.Definition.Name;

            GroupingParameterDefinition = comboBox_GroupingParameter.SelectedItem as Definition;
            ViewScheduleSheetSpreaderSettingsItem.GroupingParameterName = GroupingParameterDefinition.Name;

            double.TryParse(textBox_XOffset.Text, out XOffset);
            ViewScheduleSheetSpreaderSettingsItem.XOffsetValue = textBox_XOffset.Text;

            double.TryParse(textBox_YOffset.Text, out YOffset);
            ViewScheduleSheetSpreaderSettingsItem.YOffsetValue = textBox_YOffset.Text;

            int.TryParse(textBox_FirstSheetNumber.Text, out FirstSheetNumber);
            ViewScheduleSheetSpreaderSettingsItem.FirstSheetNumberValue = textBox_FirstSheetNumber.Text;

            SheetSizeVariantName = (this.groupBox_SheetSize.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            ViewScheduleSheetSpreaderSettingsItem.SheetSizeSelectedButtonName = SheetSizeVariantName;

            HeaderInSpecificationHeaderVariantName = (this.groupBox_HeaderInSpecificationHeader.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            ViewScheduleSheetSpreaderSettingsItem.HeaderInSpecificationHeaderSelectedButtonName = HeaderInSpecificationHeaderVariantName;

            double.TryParse(textBox_SpecificationHeaderHeight.Text, out SpecificationHeaderHeight);
            ViewScheduleSheetSpreaderSettingsItem.SpecificationHeaderHeightValue = textBox_SpecificationHeaderHeight.Text;
            ViewScheduleSheetSpreaderSettingsItem.SaveSettings();
        }
        private void SetSavedSettingsValueToForm()
        {
            if (ViewScheduleSheetSpreaderSettingsItem.SheetSizeSelectedButtonName != null)
            {
                if (ViewScheduleSheetSpreaderSettingsItem.SheetSizeSelectedButtonName == "radioButton_Type")
                {
                    radioButton_Type.IsChecked = true;
                }
            }

            if (TitleBlocksForFirstSheetCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FirstSheetFamilyName) != null)
            {
                comboBox_FirstSheetFamily.SelectedItem = TitleBlocksForFirstSheetCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FirstSheetFamilyName);
            }
            else
            {
                comboBox_FirstSheetFamily.SelectedItem = comboBox_FirstSheetFamily.Items.GetItemAt(0);
            }

            if (TitleBlocksForFirstSheetTypeCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FirstSheetTypeName) != null)
            {
                comboBox_FirstSheetType.SelectedItem = TitleBlocksForFirstSheetTypeCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FirstSheetTypeName);
            }
            else
            {
                comboBox_FirstSheetType.SelectedItem = comboBox_FirstSheetType.Items.GetItemAt(0);
            }

            if (TitleBlocksForFollowingSheetsCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsFamilyName) != null)
            {
                comboBox_FollowingSheetsFamily.SelectedItem = TitleBlocksForFollowingSheetsCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsFamilyName);
            }
            else
            {
                comboBox_FollowingSheetsFamily.SelectedItem = comboBox_FollowingSheetsFamily.Items.GetItemAt(0);
            }

            if (TitleBlocksForFollowingSheetsTypeCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsTypeName) != null)
            {
                comboBox_FollowingSheetsType.SelectedItem = TitleBlocksForFollowingSheetsTypeCollection.FirstOrDefault(tb => tb.Name == ViewScheduleSheetSpreaderSettingsItem.FollowingSheetsTypeName);
            }
            else
            {
                comboBox_FollowingSheetsType.SelectedItem = comboBox_FollowingSheetsType.Items.GetItemAt(0);
            }

            if (FamilyInstanceParametersCollection.FirstOrDefault(p => p.Definition.Name == ViewScheduleSheetSpreaderSettingsItem.SheetFormatParameterName) != null)
            {
                comboBox_SheetFormatParameter.SelectedItem = FamilyInstanceParametersCollection.FirstOrDefault(p => p.Definition.Name == ViewScheduleSheetSpreaderSettingsItem.SheetFormatParameterName);
            }
            else
            {
                comboBox_SheetFormatParameter.SelectedItem = comboBox_SheetFormatParameter.Items.GetItemAt(0);
            }

            if (ParamDefinitionsCollection.FirstOrDefault(pd => pd.Name == ViewScheduleSheetSpreaderSettingsItem.GroupingParameterName) != null)
            {
                comboBox_GroupingParameter.SelectedItem = ParamDefinitionsCollection.FirstOrDefault(pd => pd.Name == ViewScheduleSheetSpreaderSettingsItem.GroupingParameterName);
            }
            else
            {
                comboBox_GroupingParameter.SelectedItem = comboBox_GroupingParameter.Items.GetItemAt(0);
            }

            if (ViewScheduleSheetSpreaderSettingsItem.XOffsetValue != null)
            {
                textBox_XOffset.Text = ViewScheduleSheetSpreaderSettingsItem.XOffsetValue;
            }
            else
            {
                textBox_XOffset.Text = "-400";
            }
            if (ViewScheduleSheetSpreaderSettingsItem.YOffsetValue != null)
            {
                textBox_YOffset.Text = ViewScheduleSheetSpreaderSettingsItem.YOffsetValue;
            }
            else
            {
                textBox_YOffset.Text = "292";
            }
            if (ViewScheduleSheetSpreaderSettingsItem.FirstSheetNumberValue != null)
            {
                textBox_FirstSheetNumber.Text = ViewScheduleSheetSpreaderSettingsItem.FirstSheetNumberValue;
            }
            else
            {
                textBox_FirstSheetNumber.Text = "69";
            }

            if (ViewScheduleSheetSpreaderSettingsItem.HeaderInSpecificationHeaderSelectedButtonName != null)
            {
                if (ViewScheduleSheetSpreaderSettingsItem.HeaderInSpecificationHeaderSelectedButtonName == "radioButton_No")
                {
                    radioButton_No.IsChecked = true;
                }
            }

            if (ViewScheduleSheetSpreaderSettingsItem.SpecificationHeaderHeightValue != null)
            {
                textBox_SpecificationHeaderHeight.Text = ViewScheduleSheetSpreaderSettingsItem.SpecificationHeaderHeightValue;
            }
            else
            {
                textBox_SpecificationHeaderHeight.Text = "40";
            }
        }
    }
}
