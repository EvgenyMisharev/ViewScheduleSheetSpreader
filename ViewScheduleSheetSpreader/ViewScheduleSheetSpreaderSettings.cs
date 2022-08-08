using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ViewScheduleSheetSpreader
{
    public class ViewScheduleSheetSpreaderSettings
    {
        public string SheetSizeSelectedButtonName { get; set; }
        public string FirstSheetFamilyName { get; set; }
        public string FirstSheetTypeName { get; set; }
        public string FollowingSheetsFamilyName { get; set; }
        public string FollowingSheetsTypeName { get; set; }
        public string SheetFormatParameterName { get; set; }
        public string GroupingParameterName { get; set; }
        public string XOffsetValue { get; set; }
        public string YOffsetValue { get; set; }
        public string FirstSheetNumberValue { get; set; }
        public string HeaderInSpecificationHeaderSelectedButtonName { get; set; }
        public string SpecificationHeaderHeightValue { get; set; }

        public ViewScheduleSheetSpreaderSettings GetSettings()
        {
            ViewScheduleSheetSpreaderSettings viewScheduleSheetSpreaderSettings = null;
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "ViewScheduleSheetSpreaderSettings.xml";
            string assemblyPath = assemblyPathAll.Replace("ViewScheduleSheetSpreader.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                using (FileStream fs = new FileStream(assemblyPath, FileMode.Open))
                {
                    XmlSerializer xSer = new XmlSerializer(typeof(ViewScheduleSheetSpreaderSettings));
                    viewScheduleSheetSpreaderSettings = xSer.Deserialize(fs) as ViewScheduleSheetSpreaderSettings;
                    fs.Close();
                }
            }
            else
            {
                viewScheduleSheetSpreaderSettings = new ViewScheduleSheetSpreaderSettings();
            }

            return viewScheduleSheetSpreaderSettings;
        }

        public void SaveSettings()
        {
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "ViewScheduleSheetSpreaderSettings.xml";
            string assemblyPath = assemblyPathAll.Replace("ViewScheduleSheetSpreader.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                File.Delete(assemblyPath);
            }

            using (FileStream fs = new FileStream(assemblyPath, FileMode.Create))
            {
                XmlSerializer xSer = new XmlSerializer(typeof(ViewScheduleSheetSpreaderSettings));
                xSer.Serialize(fs, this);
                fs.Close();
            }
        }
    }
}
