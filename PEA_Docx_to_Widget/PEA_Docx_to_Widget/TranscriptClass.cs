using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEA_Docx_to_Widget
{
    public class TranscriptClass
    {
        public class Component
        {
            public string TemplateID { get; set; }
            public List<Transcript> Transcripts { get; set; }
        }

        public class Root
        {
            public string id { get; set; }
            public List<Component> components { get; set; }
        }

        public class Transcript
        {
            public string id { get; set; }
            public string Transcript_Text { get; set; }
        }

    }
}
