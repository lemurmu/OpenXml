using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlDemo
{
    class Program
    {
        static void Main(string[] args) {
            OpenXmlUtil.InsertIntoBookmark("Report.docx", "DC_1MΩ_50mV_CH1", "9.36mV");
            OpenXmlUtil.InsertIntoBookmark("Report.docx", "DC_1MΩ_50mV_CH1", "8.45mV");
            OpenXmlUtil.InsertIntoBookmark("Report.docx", "DC_1MΩ_50mV_CH1", "8.88mV");
            OpenXmlUtil.InsertIntoBookmark("Report.docx", "DC_1MΩ_100mV_CH1", "8.88mV");
        }
    }
}
