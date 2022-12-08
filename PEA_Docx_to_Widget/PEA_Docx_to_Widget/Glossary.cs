using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEA_Docx_to_Widget
{
    public class Glossary
    {
        public class GlossaryItem
        {
            public string Glossary_Term { get; set; }
            public string Glossary_definition { get; set; }
            public string glossaryImage { get; set; }
            public string glossaryIframe { get; set; }
        }

        public class Root
        {
            public List<GlossaryItem> glossaries { get; set; }
        }
    }
}
