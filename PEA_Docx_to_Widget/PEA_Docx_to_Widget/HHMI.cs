using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace hhmi_Docx_to_Widget
{
    public class HHMI
    {
        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
        public class AppendixElement
        {
            public string id { get; set; }
            public string templateId { get; set; }
            public string pageId { get; set; }
            public string name { get; set; }
            public string img { get; set; }
            public string imgtext { get; set; }
            public bool child { get; set; }
            public bool appendix { get; set; }
            public List<ChildElement> childElements { get; set; }
        }

        public class ChData
        {
            public string pageId { get; set; }
            public string content { get; set; }
        }

        public class ChildElement
        {
            public string id { get; set; }
            public string pageId { get; set; }
            public string templateId { get; set; }
            public string name { get; set; }
        }

        public class Footer
        {
            public string id { get; set; }
            public string contenttype { get; set; }
            public string content { get; set; }
            public string link { get; set; }
            public string title { get; set; }
            public string popupContent { get; set; }
        }

        public class Navigation
        {
            public string id { get; set; }
            public string pageId { get; set; }
            public string templateId { get; set; }
            public string name { get; set; }
            public string imgtext { get; set; }
            public string img { get; set; }
            public bool child { get; set; }
            public List<ChildElement> childElements { get; set; }
            public bool appendix { get; set; }
            public List<AppendixElement> appendixElements { get; set; }
        }

        public class Root
        {
            public string title { get; set; }
            public string search { get; set; }
            public string home { get; set; }
            public string searchicon { get; set; }
            public string help { get; set; }
            public string popupSearchImg { get; set; }
            public string download { get; set; }
            public string headerimg { get; set; }
            public string close { get; set; }
            public string collapse { get; set; }
            public string hhmi { get; set; }
            public string youtube { get; set; }
            public string facebook { get; set; }
            public string twitter { get; set; }
            public string helpPopupContent { get; set; }
            public List<Footer> footer { get; set; }
            public List<Navigation> navigation { get; set; }
            public List<ChData> chData { get; set; }
        }
    }
}
