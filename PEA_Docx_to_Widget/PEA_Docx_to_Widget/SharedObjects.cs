using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PEA_Docx_to_Widget.Glossary;
using static PEA_Docx_to_Widget.TranscriptClass;

namespace PEA_Docx_to_Widget
{
    public static class SharedObjects
    {
        public static Dictionary<string, HtmlNode> idNodes { get; set; }
        public static Dictionary<string, HtmlNode> TablePanDoc=new Dictionary<string, HtmlNode>();
        public static Dictionary<string, HtmlNode> TablesStyled = new Dictionary<string, HtmlNode>();
        public static Dictionary<string, string> ScreenTitles = new Dictionary<string, string>();
        public static List<Toc.Panel> Panels = new List<Toc.Panel>();
        public static List<string> missingglossaries = new List<string>();
        public static Toc.Component component = new Toc.Component();
        public static Dictionary<string, string> Maths = new Dictionary<string, string>();
        public static Dictionary<string, string> latexes = new Dictionary<string, string>();
        public static bool mathEnabled = false;
        public static int mathseq = 0;
        public static bool GlossaryEnable = true;
        public static int GlossaryPage { get; set; }
        public static int pageIndex { get; set; }
        public static int LastNumber { get; set; }
        public static string DocPath { get; set; }
        public static string GlossaryPageId { get; set; }
        public static Dictionary<string, GlossaryItem> glossaryList { get; set; }
        public static Dictionary<string, string> ColorList = new Dictionary<string, string>();
        public static Dictionary<string, string> popupList { get; set; }
        public static Dictionary<string, string> RubyList = new Dictionary<string, string>();
        public static Dictionary<string, Transcript> TranscriptList = new Dictionary<string, Transcript>();
        public static Dictionary<string, string> nameList = new Dictionary<string, string>();
        public static Dictionary<string, string> imageList = new Dictionary<string, string>();
    }
}
