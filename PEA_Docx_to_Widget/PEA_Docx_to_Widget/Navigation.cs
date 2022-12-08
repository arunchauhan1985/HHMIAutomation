using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEA_Docx_to_Widget
{
    public class Navigation
    {
        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
        public class Meta
        {
            public string Type { get; set; }
            public string File_name { get; set; }
            public string Version { get; set; }
            public string Due_Date { get; set; }
            public string Instruction { get; set; }
            public string Title { get; set; }
            public string Brief_description { get; set; }
            public string Long_description { get; set; }
            public string Learning_intention { get; set; }
            public string Success_criteria { get; set; }
            public string Subject { get; set; }
            public string Year_level { get; set; }
            public string Course { get; set; }
            public string Unit { get; set; }
            public string State { get; set; }
            public string AC_code { get; set; }
            public string AC_descriptor { get; set; }
            public string Estimated_time { get; set; }
            public string TocTitle { get; set; }
            public string MainText { get; set; }
            public string Template_Description { get; set; }
            public string Teacher_only { get; set; }
            public string Notes { get; set; }
            public string Discoverable { get; set; }
        }

        public class Page
        {
            public int id { get; set; }
            public string filename { get; set; }
        }

        public class Root
        {
            public Meta meta { get; set; }
            public List<Page> pages { get; set; }
        }

    }
}
