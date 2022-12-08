using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PEA_Docx_to_Widget
{
    public class Page
    {
        public class Tab
        {
            public string Tab_Heading { get; set; }
            public string Tab_Reveal_text { get; set; }
            public List<TabRevealVideo> Tab_Reveal_Video { get; set; }
        }
        public class Panel
        {
            public int id { get; set; }
            public string Panel_Heading { get; set; }
            public string PanelRevealText { get; set; }
        }
        public class TabRevealVideo
        {
            public int id { get; set; }
            public string title = "Read Transcript";
            public string videoTime = "sec";
            public string Transcript { get; set; }
        }
        public class Slide
        {
            public int id { get; set; }
            public string Slide_Title { get; set; }
            public string Slide_Text { get; set; }
            public string Slide_Graphic { get; set; }
            public string Slide_FFN { get; set; }
            public string Slide_Caption { get; set; }
            public string Slide_Acknowledgements { get; set; }
            public string Slide_Alt_text { get; set; }
            public string TextAlignment { get; set; }
            public string imageName { get; set; }
        }
        public class Main_Text
        {
            public int id { get; set; }
            public string Text_Heading { get; set; }
            public string Text_Text { get; set; }
            public string Glossary_Term { get; set; }
            public string Glossary_Definition { get; set; }
        }
        public class Text
        {
            public int id { get; set; }
            public string Text_Heading { get; set; }
            public string Text_Text { get; set; }
        }
        public class TranscriptDataVideo
        {
            public int id { get; set; }
            public string title { get; set; }
            public string videoTime = "0.00";
            public string Transcript { get; set; }
        }
        public class TranscriptDataAudio
        {
            public int id { get; set; }
            public string title { get; set; }

            public string audioTime = "0.00";
            public string Transcript_txt { get; set; }
        }
        public class Hotspot
        {
            public int id { get; set; }
            public string Hotspot_Title { get; set; }
            public string Hotspot_Text { get; set; }
            public string Hotspot_Position { get; set; }
            public string Hotspot_Reveal_title { get; set; }
            public string Hotspot_Reveal_text { get; set; }
        }
        public class Card
        {
            public string Card_Front_Heading = "";
            public string Card_Front_Content = "";
            public Card_Front_Audio Card_Front_Audio { get; set; }
            public string Card_Front_Image = "";
            public Card_Front_Video Card_Front_Video { get; set; }

            public string Card_Back_Heading = "";
            public string Card_Back_Content = "";
            public Card_Back_Audio Card_Back_Audio { get; set; }
            public string Card_Back_Image = "";
            public Card_Back_Video Card_Back_Video { get; set; }
           
            
        }

        public class Card_Front_Audio {
            public string audio = "";
            public string transcript = "";
        }
        public class Card_Front_Video
        {
            public string video = "";
            public string transcript = "";
        }
        public class Card_Back_Audio
        {
            public string audio = "";
            public string transcript = "";
        }
        public class Card_Back_Video
        {
            public string video = "";
            public string transcript = "";
        }
        public class Component
        {
            public string TemplateID { get; set; }
            public string PopupID { get; set; }
            public string PopupDataID { get; set; }
            public string TemplateName { get; set; }
            //public string Page_Title { get; set; }
            public string TemplateDescription { get; set; }
            public string Title { get; set; }
            public string MainText { get; set; }
            public string Main_text1 { get; set; }
            public string Transcript { get; set; }
            public string Graphic { get; set; }
            public string FFN { get; set; }
            public string Caption { get; set; }
            public string Acknowledgements { get; set; }
            public string AltText { get; set; }
            public string MediaPosition { get; set; }
            public string TeacherOnly { get; set; }
            public string Discoverable { get; set; }
            public string Notes { get; set; }
            public string DownloadButtonText { get; set; }
            public string DownloadPDF { get; set; }
            public List<TranscriptDataVideo> video_transcript { get; set; }
            public List<TranscriptDataAudio> audio_transcript { get; set; }
            public string ActivityID { get; set; }
            public string SRT_VTT { get; set; }
            public string Transcript_ID { get; set; }
            public string Media_position { get; set; }
            public string Timeline_graphic { get; set; }
            public List<Panel> Panels { get; set; }
            public List<Slide> Slides { get; set; }
            public List<Main_Text> Main_Texts { get; set; }
            public List<Tab> tabData { get; set; }
            public List<Hotspot> Hotspots { get; set; }
            public List<Card> flipCardData { get; set; }
            public List<Text> Texts { get; set; }
        }

        public class Root
        {
            public string id { get; set; }
            public List<Component> components { get; set; }
            public string Page_Title { get; set; }
        }
    }
}
