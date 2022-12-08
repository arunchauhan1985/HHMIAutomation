using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEA_Docx_to_Widget
{
    public class Toc
    {
        public class Component
        {
            public string TemplateID { get; set; }
            public string TemplateName { get; set; }
            public string TemplateDescription { get; set; }
            public string Title { get; set; }
            public string MainText { get; set; }
            public string TeacherOnly { get; set; }
            public string Instruction { get; set; }
            public string Lesson { get; set; }
            public string Due_Date { get; set; }
            public string Discoverable { get; set; }
            public string Notes { get; set; }
            public string DownloadPDF { get; set; }
            public string DownloadButtonText { get; set; }
            public List<Panel> Panels { get; set; }
        }

        public class Panel
        {
            public string Panel_Heading { get; set; }
            public string Panel_RevealText { get; set; }
            public string Template_id { get; set; }
        }

        public class Root
        {
            public string id { get; set; }
            public List<Component> components { get; set; }
        }
    }
}
