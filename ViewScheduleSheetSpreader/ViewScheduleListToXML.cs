using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ViewScheduleSheetSpreader
{
    public class ViewScheduleListToXML
    {
        public List<string> GetSettings()
        {
            List<string> viewScheduleList = null;
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "ViewScheduleListToXML.xml";
            string assemblyPath = assemblyPathAll.Replace("ViewScheduleSheetSpreader.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                using (FileStream fs = new FileStream(assemblyPath, FileMode.Open))
                {
                    XmlSerializer xSer = new XmlSerializer(typeof(List<string>));
                    viewScheduleList = xSer.Deserialize(fs) as List<string>;
                    fs.Close();
                }
            }
            else
            {
                viewScheduleList = null;
            }

            return viewScheduleList;
        }

        public void SaveList(List<string> viewScheduleList)
        {
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "ViewScheduleListToXML.xml";
            string assemblyPath = assemblyPathAll.Replace("ViewScheduleSheetSpreader.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                File.Delete(assemblyPath);
            }

            using (FileStream fs = new FileStream(assemblyPath, FileMode.Create))
            {
                XmlSerializer xSer = new XmlSerializer(typeof(List<string>));
                xSer.Serialize(fs, viewScheduleList);
                fs.Close();
            }
        }
    }
}
