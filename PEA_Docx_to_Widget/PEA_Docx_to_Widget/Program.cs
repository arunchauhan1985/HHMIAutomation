using DocumentFormat.OpenXml.Packaging;
using eps2math;
using HtmlAgilityPack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Xsl;
using static PEA_Docx_to_Widget.Navigation;
using static PEA_Docx_to_Widget.Page;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;
using AForge.Imaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = PEA_Docx_to_Widget.Page.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Globalization;
using static PEA_Docx_to_Widget.Glossary;
using static hhmi_Docx_to_Widget.HHMI;
using hhmi_Docx_to_Widget;

namespace PEA_Docx_to_Widget
{
    class Program
    {
        static void Main(string[] args)
        {
            //string inputDoc = null;
            //string glossaryEnable = "-glossary=yes";
            //string output_path = null;
            //string indexHtml = null;
            //if (args.Length < 3)
            //{
            //    return;
            //}
            //else
            //{
            //    if (args.Length == 3)
            //    {
            //        inputDoc = args[0];
            //        output_path = args[1];
            //        indexHtml = args[2];
            //    }
            //    else
            //    {
            //        glossaryEnable = args[0];
            //        inputDoc = args[1];
            //        output_path = args[2];
            //        indexHtml = args[3];
            //    }
            //}



            string inputDoc = @"D:\HHMI\Input\HHMI Genes Guide.docx";
            string glossaryEnable = "-glossary=no";
            string output_path = @"D:\HHMI\Input";
            if (inputDoc == null)
            {
                Console.WriteLine("Invalid Path.");
                Console.ReadLine();
            }
            else
            {
                try
                {
                    if (glossaryEnable.Split('=')[1].Trim().ToLower() == "no")
                    {
                        SharedObjects.GlossaryEnable = false;
                    }
                }
                catch (Exception) { }
                string tempFolder = Path.GetTempPath();
                string newName = Guid.NewGuid().ToString() + Path.GetExtension(inputDoc);
                output_path = WordExportProcess(inputDoc, newName, tempFolder, output_path);
                string titleName = GetTitleName(output_path + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + ".html");
                Dictionary<string, List<HtmlNode>> screensList = ImplementId(output_path + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + ".html");
                List<ChData> data = new List<ChData>();
                StringBuilder strContent = new StringBuilder();
                Dictionary<string, string> chData = new Dictionary<string, string>();
                List<HtmlNode> receipeNodes = new List<HtmlNode>();

                #region Footnote Items
                List<HHMI.Footer> footnoteItems = new List<HHMI.Footer>();
                if (screensList.ContainsKey("footnote"))
                {
                    List<HtmlNode> footnoteItem = screensList["footnote"];
                    foreach (HtmlNode node in footnoteItem)
                    {
                        HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                        hDoc.LoadHtml(node.OuterHtml);
                        HtmlAgilityPack.HtmlNodeCollection tdNodes = hDoc.DocumentNode.SelectNodes("//td//strong[text()='contenttype']|//td//strong[text()='contenttype']");

                        if (tdNodes != null)
                        {
                            HtmlNode tdNode = tdNodes[0];
                            HtmlNodeCollection trNodes = tdNode.Ancestors("table").First().SelectNodes("//tr");
                            footnoteItems = LoadFootnoteData(trNodes, footnoteItems);
                        }
                    }
                }
                #endregion

                #region Filter Receipe Nodes
                List<HtmlNode> contents = screensList["chapterContent"];
                if (contents != null)
                {
                    foreach (HtmlNode node in contents.ToList())
                    {
                        string inText = node.InnerText.Replace("\r", "").Replace("\n", "");
                        if ((inText.ToLower().Contains("imgtextrecipe cards"))|| (inText.ToLower().Contains("imgtext recipe cards")))
                        {
                            receipeNodes.Add(node);
                            contents.Remove(node);
                        }
                    }
                }
                #endregion

                #region Load Navigation Items
                List<HHMI.Navigation> NavigationItems = new List<HHMI.Navigation>();
                if (screensList.ContainsKey("toc"))
                {
                    int firstlevel = 1;
                    List<HtmlNode> tocItem = screensList["toc"];
                    foreach (HtmlNode node in tocItem)
                    {
                        HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                        hDoc.LoadHtml(node.OuterHtml);
                        HtmlAgilityPack.HtmlNodeCollection tdNodes = hDoc.DocumentNode.SelectNodes("//td//strong[text()='name']|//td//strong[text()='name']");
                        HHMI.Navigation navigation = new HHMI.Navigation();

                        if (tdNodes != null)
                        {
                            HtmlNode tdNode = tdNodes[0];
                            HtmlNodeCollection trNodes = tdNode.Ancestors("table").First().SelectNodes("//tr");
                            navigation = LoadNavData(trNodes, navigation, firstlevel);
                        }
                        NavigationItems.Add(navigation);
                        firstlevel++;
                    }
                }
                #endregion

                #region Load Recipe Items
                List<HHMI.Navigation> ReceipeItems = new List<HHMI.Navigation>();
                if (screensList.ContainsKey("recipe-toc"))
                {
                    int firstlevel = NavigationItems.Count + 1;
                    List<HtmlNode> tocItem = screensList["recipe-toc"];
                    foreach (HtmlNode node in tocItem)
                    {
                        HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                        hDoc.LoadHtml(node.OuterHtml);
                        HtmlAgilityPack.HtmlNodeCollection tdNodes = hDoc.DocumentNode.SelectNodes("//td//strong[text()='name']|//td//strong[text()='name']");
                        HHMI.Navigation navigation = new HHMI.Navigation();

                        if (tdNodes != null)
                        {
                            HtmlNode tdNode = tdNodes[0];
                            HtmlNodeCollection trNodes = tdNode.Ancestors("table").First().SelectNodes("//tr");
                            navigation = LoadRecipeData(trNodes, navigation, firstlevel, receipeNodes);
                        }
                        ReceipeItems.Add(navigation);
                        firstlevel++;
                    }
                }
                #endregion

                #region Load Chapter Data Items
                foreach (HHMI.Navigation item in NavigationItems)
                {
                    if ((item.childElements != null) && (item.childElements.Count > 0))
                    {
                        ChData data1 = new ChData();
                        string navtitle = item.name;
                        string content = CheckContent(navtitle, screensList);
                        if (content != null)
                            content = content.Replace("\"", "'");                        
                            else
                                content = "";
                        chData.Add(item.pageId, content);
                        strContent.AppendLine("\"" + item.pageId + "\"" + ":{");
                        strContent.AppendLine("\"" + "content" + "\"" + ":" + content);
                        strContent.AppendLine("},");
                        data1.pageId = item.pageId;
                        data1.content = content;
                        data.Add(data1);

                        foreach (ChildElement element in item.childElements)
                        {
                            data1 = new ChData();
                            navtitle = element.name;
                            content = CheckContent(navtitle, screensList);
                            if (content != null)
                                content = content.Replace("\"", "'");
                            else
                                content = "";
                            chData.Add(element.pageId, content);
                            strContent.AppendLine("\"" + element.pageId + "\"" + ":{");
                            strContent.AppendLine("\"" + "content" + "\"" + ":" + content);
                            strContent.AppendLine("},");
                            data1.pageId = element.pageId;
                            data1.content = content;
                            data.Add(data1);
                        }
                    }
                }
                #endregion

                string templatejson = System.Windows.Forms.Application.StartupPath + "\\template\\Guide.json";
                HHMI.Root hhmitemplateJsonObjects =
                JsonConvert.DeserializeObject<HHMI.Root>(File.ReadAllText(templatejson));
                hhmitemplateJsonObjects.navigation = NavigationItems;
                hhmitemplateJsonObjects.chData = data;
                hhmitemplateJsonObjects.footer = footnoteItems;
                string finalJson = JsonConvert.SerializeObject(hhmitemplateJsonObjects, Newtonsoft.Json.Formatting.Indented);
                string guid = Guid.NewGuid().ToString();
                string tempJson = System.Windows.Forms.Application.StartupPath + "\\Temp\\"+ guid+".json";
                File.WriteAllText(tempJson, finalJson);
                string[] allLines = File.ReadAllLines(tempJson);

                #region Update Json
                bool startEdit = false;
                StringBuilder newJson = new StringBuilder();
                foreach (string line in allLines)
                {
                    if (line.Contains("chData"))
                    {
                        startEdit = true;
                    }
                    if (startEdit == true)
                    {
                        if (line.Contains("chData"))
                        {
                            string newline = line.Replace("[", "{");
                            newJson.AppendLine(newline);
                        }
                        else 
                        {
                            if (line.Contains("pageId"))
                            {
                                string newline = line.Replace(": ", ":");
                                Regex reg = new Regex("\"" + "pageId" + "\"" + ":" + "(.+?)" + ",");
                                newline = reg.Replace(newline, "$1" + ":{");
                                newJson.AppendLine(newline);
                            }
                            else
                            {
                                if (line.Trim() == "]")
                                {
                                    string newline = line.Replace("]", "}");
                                    newJson.AppendLine(newline);
                                }
                                else 
                                {
                                    if (line.Trim() == "{")
                                    {
                                    }
                                    else
                                    {
                                        newJson.AppendLine(line);
                                    }
                                }
                            }
                        }
                    }
                    else 
                    {
                        newJson.AppendLine(line);
                    }
                }
                File.WriteAllText(tempJson, newJson.ToString());
                #endregion

                File.Copy(tempJson, Path.GetDirectoryName(inputDoc) + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + ".json");
            }
        }
        private static List<HHMI.Footer> LoadFootnoteData(HtmlNodeCollection trNodes, List<HHMI.Footer> footnoteItems)
        {
            foreach (HtmlNode trNode in trNodes)
            {
                HtmlNode chNode = GetCellNode(trNode);
                if (chNode.Name != "#text")
                {
                    if (chNode.InnerText.Trim() == "contenttype")
                    {
                        HHMI.Footer footer = new HHMI.Footer();
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        footer.contenttype = value;
                        footer.id = Convert.ToString(footnoteItems.Count+1);
                        footnoteItems.Add(footer);
                    }
                    if (chNode.InnerText.Trim() == "title")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();

                        footnoteItems[footnoteItems.Count - 1].title = value;
                    }
                    if (chNode.InnerText.Trim() == "content")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.ToLower().Trim().Replace("\r","").Replace("\n", "").Trim()+".png";

                        footnoteItems[footnoteItems.Count - 1].content = value;
                    }
                    if (chNode.InnerText.Trim() == "link")
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string value = nextSibling.InnerText.ToLower().Trim();

                        footnoteItems[footnoteItems.Count - 1].link = value;
                    }
                }
            }
            return footnoteItems;
        }
        private static HHMI.Navigation LoadNavData(HtmlNodeCollection trNodes, HHMI.Navigation navigation, int firstlevel)
        {
            bool child = false;
            bool appendix = false;
            foreach (HtmlNode trNode in trNodes)
            {
                HtmlNode chNode = GetCellNode(trNode);
                if (chNode.Name != "#text")
                {
                    if ((child == false) && (chNode.InnerText.Trim() == "name"))
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.name = value;
                        navigation.id = firstlevel.ToString();
                        navigation.pageId = firstlevel.ToString();
                    }
                    if ((child == true) && (chNode.InnerText.Trim() == "name"))
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();

                        navigation.childElements[navigation.childElements.Count - 1].name = value;
                    }
                    if (chNode.InnerText.Trim() == "child")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.ToLower().Trim();
                        if (value.ToLower().Trim().Contains("true"))
                        {
                            child = true;
                            navigation.child = true;
                        }
                        else
                        {
                            navigation.child = false;
                        }
                    }
                    if (chNode.InnerText.Trim() == "appendix")
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string value = nextSibling.InnerText.ToLower().Trim();
                        if (value.ToLower().Trim().Contains("true"))
                        {
                            appendix = true;
                            navigation.appendix = true;
                        }
                        else { navigation.appendix = false; }
                    }
                    if (chNode.InnerText.Trim() == "imgtext")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.imgtext = value;
                    }
                    if (chNode.InnerText.Trim() == "img")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.img = value;
                    }
                    if ((child == false) && (chNode.InnerText.Trim() == "templateId"))
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string templateId = nextSibling.InnerText.Trim();
                        navigation.templateId = templateId;
                    }
                    if ((child == true) && (chNode.InnerText.Trim() == "templateId"))
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string templateId = nextSibling.InnerText.Trim();

                        ChildElement childElement = new ChildElement();
                        childElement.templateId = templateId;
                        if (navigation.childElements == null)
                        {
                            List<ChildElement> elements = new List<ChildElement>();
                            navigation.childElements = elements;
                        }

                        childElement.id = Convert.ToString(navigation.childElements.Count + 1);
                        childElement.pageId = navigation.pageId + "-" + childElement.id;
                        navigation.childElements.Add(childElement);
                    }
                }
            }
            return navigation;
        }

        private static HHMI.Navigation LoadRecipeData(HtmlNodeCollection trNodes, HHMI.Navigation navigation, int firstlevel, List<HtmlNode> receipeNodes)
        {
            bool child = false;
            bool appendix = false;
            foreach (HtmlNode trNode in trNodes)
            {
                HtmlNode chNode = GetCellNode(trNode);
                if (chNode.Name != "#text")
                {
                    if ((child == false) && (chNode.InnerText.Trim() == "name"))
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.name = value;
                        navigation.id = firstlevel.ToString();
                        navigation.pageId = firstlevel.ToString();
                    }
                    if ((child == true) && (chNode.InnerText.Trim() == "name"))
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();

                        navigation.childElements[navigation.childElements.Count - 1].name = value;
                    }
                    if (chNode.InnerText.Trim() == "child")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.ToLower().Trim();
                        if (value.ToLower().Trim().Contains("true"))
                        {
                            child = true;
                            navigation.child = true;
                        }
                        else
                        {
                            navigation.child = false;
                        }
                    }
                    if (chNode.InnerText.Trim() == "appendix")
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string value = nextSibling.InnerText.ToLower().Trim();
                        if (value.ToLower().Trim().Contains("true"))
                        {
                            appendix = true;
                            navigation.appendix = true;
                        }
                        else { navigation.appendix = false; }
                    }
                    if (chNode.InnerText.Trim() == "imgtext")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.imgtext = value;
                    }
                    if (chNode.InnerText.Trim() == "img")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string value = nextSib.InnerText.Trim();
                        navigation.img = value;
                    }
                    if ((child == false) && (chNode.InnerText.Trim() == "templateId"))
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string templateId = nextSibling.InnerText.Trim();
                        navigation.templateId = templateId;
                    }
                    if ((child == true) && (chNode.InnerText.Trim() == "templateId"))
                    {
                        HtmlNode nextSibling = GetNextSibling(chNode);
                        string templateId = nextSibling.InnerText.Trim();

                        ChildElement childElement = new ChildElement();
                        childElement.templateId = templateId;
                        if (navigation.childElements == null)
                        {
                            List<ChildElement> elements = new List<ChildElement>();
                            navigation.childElements = elements;
                        }

                        childElement.id = Convert.ToString(navigation.childElements.Count + 1);
                        childElement.pageId = navigation.pageId + "-" + childElement.id;
                        navigation.childElements.Add(childElement);
                    }
                }
            }

            if (appendix == true)
            {
                foreach (HtmlNode trNode in receipeNodes)
                { 
                
                }
            }
            return navigation;
        }

        private static HtmlNode CheckChild(HtmlNode node)
        {
            HtmlNode nextSib = node.NextSibling;
            for (int k = 0; k < 4; k++)
            {
                if (nextSib.Name != "#text")
                {
                    if (nextSib.InnerText == "child")
                    { 
                    
                    }
                }
                else
                    nextSib = nextSib.NextSibling;
            }
            return nextSib;
        }
        private static string CheckContent(string title, Dictionary<string, List<HtmlNode>> screensList)
        {
            string content = null;
            List<HtmlNode> contents = screensList["chapterContent"];
            int i = 0;
            foreach (HtmlNode node in contents)
            {
                HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                hDoc.LoadHtml(node.OuterHtml);
                HtmlAgilityPack.HtmlNodeCollection tdNodes = hDoc.DocumentNode.SelectNodes("//td");
                bool matched = false;
                foreach (HtmlNode chNode in tdNodes)
                {
                    if (chNode.InnerText.Trim() == "name")
                    {
                        HtmlNode nextSib = GetNextSibling(chNode);
                        string titleContent = nextSib.InnerText;
                        if (title == titleContent)
                        {
                            matched = true;
                        }
                    }
                    if (matched == true)
                    {
                        if (chNode.InnerText.ToLower().Trim() == "content")
                        {
                            HtmlNode nextSib = GetNextSibling(chNode);
                            content = nextSib.InnerHtml;
                            break;
                        }
                    }
                }
                if (content != null)
                {
                    break;
                }
                i++;
            }
            return content;
        }
        private static HtmlNode GetCellNode(HtmlNode node)
        {
            HtmlNode childNode= node;
            foreach (HtmlNode chNode in node.ChildNodes)
            {
                if ((chNode.Name == "th") || (chNode.Name == "td"))
                {
                    childNode = chNode;
                    break;
                }
            }
            return childNode;
        }
        private static HtmlNode GetNextSibling(HtmlNode node)
        {
            HtmlNode nextSib = node.NextSibling;
            if (nextSib != null)
            {
                for (int k = 0; k < 4; k++)
                {
                    if (nextSib.Name != "#text")
                        break;
                    else
                        nextSib = nextSib.NextSibling;
                }
            }
            return nextSib;
        }
        private static HtmlNode GetTrNextSibling(HtmlNode node)
        {
            HtmlNode trNode = node.Ancestors("td").First();
            HtmlNode nextSib = trNode.NextSibling;
            if (nextSib != null)
            {
                for (int k = 0; k < 4; k++)
                {
                    if (nextSib.Name != "#text")
                        break;
                    else
                        nextSib = nextSib.NextSibling;
                }
            }
            return nextSib;
        }
        private static string WordExportProcess(string inputDoc, string newName, string tempFolder, string output_path)
        {
            File.Copy(inputDoc, tempFolder + "\\" + newName, true);
            string docPath = tempFolder + "\\" + newName;
            string tempPath = System.Windows.Forms.Application.StartupPath + "\\Temp";
            if (!Directory.Exists(tempPath))
                Directory.CreateDirectory(tempPath);
            string htmlPath = tempPath + "\\" + Path.GetFileNameWithoutExtension(docPath) + ".html";
            string tempFolderApp = tempPath + "\\" + Path.GetFileNameWithoutExtension(docPath);
            if (!Directory.Exists(tempFolderApp))
                Directory.CreateDirectory(tempFolderApp);
            Console.WriteLine("Template conversion is in process.....");
            #region Math Execution
            string MathFunctionsEnable = ConfigurationManager.AppSettings.Get("MathFunctionsEnable");
            //if (SharedObjects.mathEnabled == true)
            {

                if (MathFunctionsEnable.ToLower().Trim() == "true")
                {
                    docPath = AddDocPreandSuf(docPath);
                    docPath = ConvertintoLatex(docPath);
                }
            }
            docPath = UpdateList(docPath);
            docPath = UpdateColor(docPath);
            docPath = AddRuby(docPath);

            //docPath = UpdateMath(docPath);
            //docPath = getmaths(docPath, tempFolderApp);

            #endregion
            string command = System.Windows.Forms.Application.StartupPath + "\\lib\\" + "pandoc.exe -t html -s " + docPath + " -o " + htmlPath + " -N --extract-media=" + System.IO.Path.GetDirectoryName(docPath) + "\\" + System.IO.Path.GetFileNameWithoutExtension(docPath).Replace(" ", "_") + "_images";
            ExecuteCommandMain(command);
            ConvertDocx(docPath, htmlPath);

            string htmlContent = File.ReadAllText(htmlPath).Replace("[color]", "</span>");
            Regex regcolor = new Regex(@"\[color\((.+?)\)\]");
            htmlContent = regcolor.Replace(htmlContent, "<span style=\"color:$1" + "\"" + ">");

            htmlContent = htmlContent.Replace("\\[", "$").Replace("\\]", "$");
            if (MathFunctionsEnable.ToLower().Trim() == "true")
            {
                MatchCollection matches = Regex.Matches(htmlContent, @"&lt;math&gt;(.+?)&lt;/math&gt;", RegexOptions.Singleline);
                foreach (Match item in matches)
                {
                    string latex = item.Value.Replace("\r", "").Replace("</p>\n<p>", "").Replace("&lt;math&gt;", "").Replace("&lt;/math&gt;", "").Trim().Trim('$').Replace("&amp;", "&");
                    latex = "<p>" + latex + "</p>";
                    latex = latex.Replace("<p>\\[", "").Replace("\\]</p>", "").Replace("<p>", "").Replace("</p>", "");
                    string correctedValue = "<span id=" + "\"" + "mathjax" + "\"" + " latex=" + "\"" + latex + "\"" + "></span>";
                    htmlContent = htmlContent.Replace(item.Value, correctedValue);
                }
            }
            else
            {
                MatchCollection matches = Regex.Matches(htmlContent, @"\$(.+?)\$", RegexOptions.Singleline);
                foreach (Match item in matches)
                {
                    string latex = item.Value.Replace("\r", "").Replace("</p>\n<p>", "").Replace("&lt;math&gt;", "").Replace("&lt;/math&gt;", "").Trim().Trim('$').Replace("&amp;", "&").Replace("</span>", "");
                    Regex regMath = new Regex(@"<span(.+?)>");
                    latex = regMath.Replace(latex, "");
                    latex = "<p>" + latex + "</p>";
                    latex = latex.Replace("<p>\\[", "").Replace("\\]</p>", "").Replace("<p>", "").Replace("</p>", "");
                    string correctedValue = "<span id=" + "\"" + "mathjax" + "\"" + " latex=" + "\"" + latex + "\"" + "></span>";
                    htmlContent = htmlContent.Replace(item.Value, correctedValue);
                }
            }

            #region Filter Color Values

            HtmlAgilityPack.HtmlDocument cDoc = new HtmlAgilityPack.HtmlDocument();
            cDoc.OptionWriteEmptyNodes = true;
            cDoc.LoadHtml(htmlContent);
            HtmlNodeCollection spanNodes = cDoc.DocumentNode.SelectNodes("//span[@style]");

            foreach (HtmlNode colorNode in spanNodes.ToList())
            {
                try
                {
                    if (colorNode.Attributes["style"].Value.Contains("color:"))
                    {
                        string colorCode = colorNode.Attributes["style"].Value.Split(':')[1];
                        string systemColor = getColor(colorCode);
                        if (systemColor.StartsWith("#"))
                        {
                            System.Drawing.Color c = System.Drawing.ColorTranslator.FromHtml(colorCode);
                            systemColor = GetKnownColor(c.ToArgb(), colorCode);
                        }
                        if (systemColor == "black")
                        {
                            colorNode.Attributes.RemoveAll();
                            colorNode.Name = "temp";
                        }
                    }
                }
                catch (Exception)
                { }
            }
            spanNodes = cDoc.DocumentNode.SelectNodes("//span[@style]");
            if (spanNodes != null)
            {
                foreach (HtmlNode colorNode in spanNodes.ToList())
                {
                    if ((colorNode.Attributes != null) && (colorNode.Attributes.Count > 0))
                    {
                        if (colorNode.Attributes["style"].Value.Contains("color:"))
                        {
                            string firstcolorCode = colorNode.Attributes["style"].Value.Split(':')[1];
                            HtmlNode nextSibling = colorNode.NextSibling;
                            for (int i = 0; i < 10000; i++)
                            {
                                if ((nextSibling != null) && (nextSibling.Name == "span") && (nextSibling.Attributes.Contains("style")))
                                {
                                    if (nextSibling.Attributes["style"].Value.Contains("color:"))
                                    {
                                        string nextcolorCode = nextSibling.Attributes["style"].Value.Split(':')[1];
                                        if (firstcolorCode == nextcolorCode)
                                        {
                                            if (!colorNode.Attributes.Contains("class"))
                                            {
                                                colorNode.Attributes.Add("class", "first tomerge");
                                            }
                                            if (!nextSibling.Attributes.Contains("class"))
                                            {
                                                nextSibling.Attributes.Add("class", "tomerge");
                                            }
                                            nextSibling = nextSibling.NextSibling;
                                        }
                                        else { break; }
                                    }
                                }
                                else { break; }
                            }
                        }
                    }
                }
            }
            spanNodes = cDoc.DocumentNode.SelectNodes("//span[@style]");
            if (spanNodes != null)
            {
                foreach (HtmlNode colorNode in spanNodes.ToList())
                {
                    if (colorNode.Attributes.Contains("class"))
                    {
                        string className = colorNode.Attributes["class"].Value;
                        if (className.Contains("first "))
                        {
                            colorNode.Attributes.Remove("class");
                        }
                        else
                        {
                            colorNode.Attributes.Remove("style");
                            colorNode.Attributes.Remove("class");
                        }
                    }
                }
            }
            #endregion
            htmlContent = cDoc.DocumentNode.InnerHtml.Replace("<temp>", "").Replace("</temp>", "").Replace("<temp/>", "").Replace("</span><span>", "");
            if (SharedObjects.RubyList.Count > 0)
            {
                foreach (KeyValuePair<string, string> item in SharedObjects.RubyList)
                {
                    string id = item.Key;
                    string ruby = item.Value;
                    ruby = ruby.Replace("w:", "");
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.OptionWriteEmptyNodes = true;
                    tDoc.LoadHtml(ruby);
                    HtmlNode rtNode = tDoc.DocumentNode.SelectSingleNode("//rt");
                    HtmlNode tNode = tDoc.DocumentNode.SelectSingleNode("//rt//t");
                    rtNode.InnerHtml = tNode.InnerHtml;
                    HtmlNode rtbNode = tDoc.DocumentNode.SelectSingleNode("//rubybase");
                    HtmlNode tbNode = tDoc.DocumentNode.SelectSingleNode("//rubybase//t");
                    rtbNode.InnerHtml = tbNode.InnerHtml;

                    string newTag = "<ruby>" + rtbNode.InnerHtml + rtNode.OuterHtml + "</ruby>";
                    htmlContent = htmlContent.Replace(id, newTag.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", ""));
                }
            }
            //htmlContent = MathHTMLUpdate(htmlContent);
            File.WriteAllText(htmlPath, htmlContent);
            GetAllTables(htmlPath);
            string[] dataDir = Directory.GetDirectories(output_path, "data", SearchOption.AllDirectories);
            if (dataDir.Length > 0)
            {
                output_path = dataDir[0];
            }
            File.Copy(htmlPath, output_path + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + ".html", true);
            return output_path;
        }
        private static String GetKnownColor(int iARGBValue, string hexcode)
        {
            System.Drawing.Color someColor;
            string knowncolor = hexcode;
            Array aListofKnownColors = Enum.GetValues(typeof(KnownColor));
            foreach (KnownColor eKnownColor in aListofKnownColors)
            {
                someColor = System.Drawing.Color.FromKnownColor(eKnownColor);
                int arg = someColor.ToArgb();
                if (iARGBValue == someColor.ToArgb() && !someColor.IsSystemColor)
                {
                    knowncolor = someColor.Name;
                }
            }
            return knowncolor.ToLower();
        }
        private static string getColor(string hexCode)
        {
            string color = hexCode;
            if (hexCode == "#FFFFFF") { color = "white"; }
            if (hexCode == "#EF0208") { color = "red"; }
            if (hexCode == "#C00000") { color = "red-dark"; }
            if (hexCode == "#FFFF00") { color = "yellow"; }
            if (hexCode == "#FF8000") { color = "orange"; }
            if (hexCode == "#C0FFC0") { color = "green"; }
            if (hexCode == "#00C000") { color = "green-dark"; }
            if (hexCode == "#00FFFF") { color = "blue"; }
            if (hexCode == "#006FC7") { color = "blue-dark"; }
            if (hexCode == "#FF80FF") { color = "pink"; }
            if (hexCode == "#7E2FA8") { color = "purple-dark"; }
            if (hexCode == "#FDC9A0") { color = "brown"; }
            if (hexCode == "#C9500B") { color = "brown-dark"; }
            if (hexCode == "#7B867C") { color = "grey"; }
            if (hexCode == "#000000") { color = "black"; }
            return color;
        }
        private static string AddDocPreandSuf(string docPath)
        {
            string newdocPath = docPath;
            Dictionary<string, string> maths = new Dictionary<string, string>();
            string outFile = Path.GetDirectoryName(docPath) + "\\" + Path.GetFileNameWithoutExtension(docPath) + "_math" + Path.GetExtension(docPath);
            File.Copy(docPath, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);
            string[] wmfes = Directory.GetFiles(inputDocFolder, "*.wmf", SearchOption.AllDirectories);
            if (wmfes.Length > 0)
            {
                SharedObjects.mathEnabled = true;
                if (documentFile.Length > 0)
                {
                    string documentFileText = File.ReadAllText(documentFile[0]).Replace("<w:object", "<w:r w:rsidRPr=\"0057557C\"><w:rPr><w:rFonts w:ascii=\"Calibri\" w:eastAsia=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\" /></w:rPr><w:t>&lt;math&gt;</w:t></w:r><w:object").Replace("</w:object>", "</w:object><w:r w:rsidRPr=\"0057557C\"><w:rPr><w:rFonts w:ascii=\"Calibri\" w:eastAsia=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\" /></w:rPr><w:t>&lt;/math&gt;</w:t></w:r>");
                    File.WriteAllText(documentFile[0], documentFileText);
                    newdocPath = zipFile(inputDocFolder, Path.GetDirectoryName(inputDocFolder), docPath);
                }
            }

            return newdocPath;
        }
        private static string AddRuby(string docPath)
        {
            string newdocPath = docPath;
            string outFile = Path.GetDirectoryName(docPath) + "\\" + Path.GetFileNameWithoutExtension(docPath) + "_ruby" + Path.GetExtension(docPath);
            File.Copy(docPath, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);

            string documentText = File.ReadAllText(documentFile[0]).Replace("w:ruby", "ruby");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(documentText);
            XmlNodeList xmlNodeList = doc.DocumentElement.SelectNodes("//ruby");
            if (xmlNodeList!=null)
            {
                int id = 1;
                foreach(XmlNode xmlNode in xmlNodeList)
                {
                    string rubyId = "[RUBYID" + id + "]";
                    SharedObjects.RubyList.Add(rubyId, xmlNode.OuterXml);
                    xmlNode.Attributes.RemoveAll();
                    xmlNode.InnerXml = "<w:t>" + rubyId + "</w:t>";
                    id++;
                }
                doc.DocumentElement.InnerXml = doc.DocumentElement.InnerXml.Replace("<ruby>", "").Replace("</ruby>", "");
                doc.Save(documentFile[0]);
                newdocPath = zipFile(inputDocFolder, Path.GetDirectoryName(inputDocFolder), docPath);
            }

            return newdocPath;
        }
        private static string UpdateColor(string docPath)
        {
            string newdocPath = docPath;
            string outFile = Path.GetDirectoryName(docPath) + "\\" + Path.GetFileNameWithoutExtension(docPath) + "_colored" + Path.GetExtension(docPath);
            File.Copy(docPath, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);

            string documentText = File.ReadAllText(documentFile[0]).Replace("w:color", "w_color");
                //.Replace("w:pStyle", "w_pStyle").Replace("w:val=", "w_val=");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(documentText);
            XmlNodeList xmlNodeList = doc.DocumentElement.SelectNodes("//w_color");
            if (xmlNodeList != null)
            {
                int id = 1;
                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    string colorId = "Color" + id;
                    string colorval = "#" + xmlNode.Attributes["w:val"].Value;
                    XmlNode tNode = null;
                    for (int i = 0; i < 5; i++)
                    {
                        tNode = xmlNode.ParentNode;
                        XmlNode xmlNode1 = tNode.NextSibling;
                        bool done = false;
                        for (int j = 0; j < 5; j++)
                        {
                            if (xmlNode1 == null)
                            {
                                break;
                            }
                            else
                            {
                                if (xmlNode1.Name == "w:t")
                                {
                                    xmlNode1.InnerXml = "[color(" + colorval + ")]" + xmlNode1.InnerXml + "[color]";
                                    done = true;
                                    break;
                                }
                            }
                        }

                        if (done == true)
                        {
                            break;
                        }
                    }
                    //if(xmlNode)
                    id++;
                }
            //}
            //XmlNodeList ListNodes = doc.DocumentElement.SelectNodes("//w_pStyle[@w_val='ListParagraph']");
            //if (ListNodes != null)
            //{
            //    int j = 1;
            //    foreach (XmlNode xmlNode in ListNodes)
            //    {
            //        XmlNode ancestorNode = xmlNode.ParentNode;
            //        for (int i = 0; i < 5; i++)
            //        {
            //            if (ancestorNode.Name == "w:p")
            //            {
            //                break;
            //            }
            //            else 
            //            {
            //                ancestorNode = ancestorNode.ParentNode;
            //            }
            //        }
            //        int k = 0;
            //        foreach (XmlNode chNode in ancestorNode.ChildNodes)
            //        {
            //            if (chNode.Name == "w:r")
            //            {
            //                if (k == 0)
            //                {
                               
            //                }
            //            }
            //        }
            //        string wt = "<w:r><w:t>[#MYLIST-ID"+(j)+"]</w:t></w:r>";
            //        ancestorNode.InnerXml = ancestorNode.InnerXml + wt;
            //        j++;
            //    }

                doc.DocumentElement.InnerXml = doc.DocumentElement.InnerXml.Replace("w_color", "w:color").Replace("\\(", "$").Replace("\\)", "$").Replace("\\[", "$").Replace("\\]", "$");
                doc.Save(documentFile[0]);
                newdocPath = zipFile(inputDocFolder, Path.GetDirectoryName(inputDocFolder), docPath);
            }

            return newdocPath;
        }
        private static string UpdateList(string docPath)
        {
            string newdocPath = docPath;
            string outFile = Path.GetDirectoryName(docPath) + "\\" + Path.GetFileNameWithoutExtension(docPath) + "_colored" + Path.GetExtension(docPath);
            File.Copy(docPath, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);

            string documentText = File.ReadAllText(documentFile[0]).Replace("w:pStyle", "w_pStyle").Replace("w:val=", "w_val=").Replace("w:ilvl", "w_ilvl").Replace("w:numPr", "w_numPr");
            //.Replace("w:pStyle", "w_pStyle").Replace("w:val=", "w_val=");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(documentText);

            XmlNodeList ListNodes = doc.DocumentElement.SelectNodes("//w_pStyle[@w_val='ListParagraph']|//w_ilvl");
            if (ListNodes != null)
            {
                foreach (XmlNode xmlNode in ListNodes)
                {
                    XmlNode ancestorNode = xmlNode.ParentNode;
                    if (xmlNode.Name == "w_ilvl")
                    {
                        ancestorNode = ancestorNode.ParentNode;
                    }
                    for (int i = 0; i < 5; i++)
                    {
                        if (ancestorNode.Name == "w:p")
                        {
                            break;
                        }
                        else
                        {
                            ancestorNode = ancestorNode.ParentNode;
                        }
                    }
                    for (int i = 0; i < ancestorNode.ChildNodes.Count; i++)
                    {
                        XmlNode chNode = ancestorNode.ChildNodes[i];
                        if (chNode.Name == "w:r")
                        {
                            foreach (XmlNode xmlNode1 in chNode.ChildNodes)
                            {
                                if (xmlNode1.Name == "w:t")
                                {
                                    string id = Guid.NewGuid().ToString();
                                    xmlNode1.InnerXml = "&lt;listitem id="+"\""+ id + "\""+"&gt;" + xmlNode1.InnerXml;
                                }
                            }
                            break;
                        }
                    }
                    for (int i = ancestorNode.ChildNodes.Count - 1; i >= 0; i--)
                    {
                        XmlNode chNode = ancestorNode.ChildNodes[i];
                        if (chNode.Name == "w:r")
                        {
                            foreach (XmlNode xmlNode1 in chNode.ChildNodes)
                            {
                                if (xmlNode1.Name == "w:t")
                                {
                                    xmlNode1.InnerXml = xmlNode1.InnerXml+ "&lt;/listitem&gt;";
                                }
                            }
                            break;
                        }
                    }
                }
                doc.DocumentElement.InnerXml = doc.DocumentElement.InnerXml.Replace("w_pStyle", "w:pStyle").Replace("w_val=", "w:val=").Replace("w_ilvl", "w:ilvl").Replace("w_numPr", "w:numPr");
                doc.Save(documentFile[0]);
                newdocPath = zipFile(inputDocFolder, Path.GetDirectoryName(inputDocFolder), docPath);
            }

            return newdocPath;
        }
        private static string UpdateMath(string inputdoc)
        {
            //using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(inputdoc, true))
            //{
            //    var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            //    foreach (var paragraph in paragraphs)
            //    {
            //        if (paragraph.InnerText.Contains("math"))
            //        { 

            //        }
            //        Console.WriteLine(paragraph.InnerText);
            //    }
            //}
            return inputdoc;
        }
        private static string MathHTMLUpdate(string htmlContent)
        {
            HtmlAgilityPack.HtmlDocument reldoc = new HtmlAgilityPack.HtmlDocument();
            reldoc.LoadHtml(htmlContent);
            HtmlNodeCollection tNodes = reldoc.DocumentNode.SelectNodes("//*[contains(text(),'$')]|//*[contains(text(),'\\[')]");
            if (tNodes.Count > 0)
            {
                foreach (HtmlNode mathNode in tNodes)
                {
                    string mathdata = mathNode.InnerText.Trim();
                    if (((mathdata.StartsWith("$")) && (mathdata.EndsWith("$"))) || ((mathdata.StartsWith("$")) && (mathdata.EndsWith("$"))))
                    {

                    }
                    else
                    {
                        if ((mathdata.StartsWith("$")) || (mathdata.StartsWith("$")))
                        {
                            HtmlNode mathStartNode = reldoc.CreateElement("mathstart");
                            mathNode.ParentNode.InsertBefore(mathStartNode, mathNode);
                        }
                        if ((mathdata.EndsWith("$")) || (mathdata.EndsWith("$")))
                        {
                            HtmlNode mathStartNode = reldoc.CreateElement("mathend");
                            mathNode.ParentNode.InsertAfter(mathStartNode, mathNode);
                        }
                    }
                }
            }
            string relHtml = reldoc.DocumentNode.InnerHtml
                .Replace("\r\n<temp></temp>", "").Replace("\r\n<temp />", "").Replace("\r\n<temp>", "")
                .Replace("\n<temp></temp>", "").Replace("\n<temp />", "").Replace("\n<temp>", "").Replace("\r", "").Replace("\n", "");
            return relHtml;
        }
        private static string MathHTMLUpdate2(string htmlContent)
        {
            HtmlAgilityPack.HtmlDocument reldoc = new HtmlAgilityPack.HtmlDocument();
            reldoc.LoadHtml(htmlContent);
            HtmlNodeCollection tNodes = reldoc.DocumentNode.SelectNodes("//*[contains(text(),'$')]|//*[contains(text(),'\\[')]");
            if (tNodes.Count > 0)
            {
                foreach (HtmlNode mathNode in tNodes)
                {
                    StringBuilder strmath = new StringBuilder();
                    HtmlNode nextSibling = mathNode.NextSibling;
                    bool mathcont = false;
                    for (int i = 0; i < 500; i++)
                    {
                        if ((nextSibling == null) || ((nextSibling != null) && (!nextSibling.InnerText.Contains("&lt;"))))
                        {
                            if (mathcont == false)
                            {
                                nextSibling = mathNode.ParentNode.NextSibling;
                                if (nextSibling.Name == "#text")
                                {
                                    nextSibling = nextSibling.NextSibling;
                                }
                            }
                            else
                            {
                                if ((nextSibling != null) && (!nextSibling.InnerText.Contains("&lt;")))
                                {
                                    strmath.AppendLine(nextSibling.InnerText);
                                    nextSibling.InnerHtml = "[DEL]";
                                    nextSibling = nextSibling.NextSibling;
                                    if (nextSibling.Name == "#text")
                                    {
                                        nextSibling = nextSibling.NextSibling;
                                    }
                                }
                            }
                        }
                        else
                        {
                            mathcont = true;
                            strmath.AppendLine(nextSibling.InnerText);
                            nextSibling.InnerHtml = "[DEL]";
                            nextSibling = nextSibling.NextSibling;
                            if (nextSibling.Name == "#text")
                            {
                                nextSibling = nextSibling.NextSibling;
                            }
                        }
                        if (nextSibling.InnerHtml.Contains("&lt;/math"))
                        {
                            mathcont = false;
                            strmath.AppendLine(nextSibling.InnerText);
                            nextSibling.InnerHtml = "[DEL]";
                            break;
                        }
                    }
                    string math = strmath.Replace("&lt;", "<").Replace("&gt;", ">").ToString();
                    mathNode.InnerHtml = mathNode.InnerHtml.Replace("&lt;math&gt;", math).Replace("&lt;math display='block'&gt;", math);
                }
            }
            tNodes = reldoc.DocumentNode.SelectNodes("//*[contains(text(),'[DEL]')]");
            if (tNodes != null)
            {
                foreach (HtmlNode node in tNodes.ToList())
                {
                    HtmlNode ttemp = reldoc.CreateElement("temp");
                    node.ParentNode.ReplaceChild(ttemp, node);
                }
            }
            string relHtml = reldoc.DocumentNode.InnerHtml
                .Replace("\r\n<temp></temp>", "").Replace("\r\n<temp />", "").Replace("\r\n<temp>", "")
                .Replace("\n<temp></temp>", "").Replace("\n<temp />", "").Replace("\n<temp>", "").Replace("\r", "").Replace("\n", "");
            return relHtml;
        }
        private static int GetWMFsCount(string inputDoc)
        {
            int count = 0;
            string newdocPath = inputDoc;
            Dictionary<string, string> maths = new Dictionary<string, string>();
            string outFile = Path.GetDirectoryName(inputDoc) + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + "_math" + Path.GetExtension(inputDoc);
            File.Copy(inputDoc, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);
            string[] reldocumentFile = Directory.GetFiles(inputDocFolder, "document.xml.rels", SearchOption.AllDirectories);
            string relFile = "";
            if (reldocumentFile.Length > 0)
            {
                relFile = reldocumentFile[0];
            }
            if (reldocumentFile.Length > 0)
            {
                string textFile = File.ReadAllText(reldocumentFile[0]).Replace(" xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"", "");
                string documentFileText = File.ReadAllText(documentFile[0]).Replace(" r:id=", " r_id=").Replace("w:drawing", "w_drawing").Replace("w:object", "w_object");
                File.WriteAllText(documentFile[0], documentFileText);

                MatchCollection matchesdraw = Regex.Matches(documentFileText, @"</w_drawing>");
                MatchCollection matchesobject = Regex.Matches(documentFileText, @"</w_object>");
                count = matchesdraw.Count + matchesobject.Count;
            }
            return count;
        }
        private static List<string> GetAllWMFs(string sourceFile)
        {
            List<string> images = new List<string>();
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(sourceFile, true))
            {
                var imageParts = myDocument.MainDocumentPart.ImageParts;
                foreach (ImagePart imagePart in imageParts)
                {
                    var uri = imagePart.Uri;
                    var filename = uri.ToString().Split('/').Last();
                    if (imagePart.ContentType.Equals("image/x-wmf"))
                    {
                        images.Add(filename);
                    }
                }
            }
            return images;
        }
        private static string ConvertintoLatex(string docFile)
        {
            int mathCount = GetWMFsCount(docFile);
            Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            app.Activate();
            app.Documents.Open(docFile);
            for (int i = 0; i < mathCount; i++)
            {
                app.Run("BrowseEquationsForward");
                app.Run("MTCommand_TeXToggle");
                Thread.Sleep(500);
            }
            //app.Run("MTCommand_ConvertEqns");
            //app.Run("MTCommand_ConvertEqns");

            string newdoc = Path.GetDirectoryName(docFile) + "\\" + Path.GetFileNameWithoutExtension(docFile) + "_math.docx";
            app.ActiveDocument.SaveAs2(newdoc, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
            app.ActiveDocument.Close();
            try
            {
                //app.Visible = false;
                app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

                //app.Visible = false;
                app.Quit(false);
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
            }
            catch (Exception)
            { app.Visible = false; }
            //Thread.Sleep(1000);
            //InputSimulator inputSimulator = new InputSimulator();
            //inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
            //Thread.Sleep(500);
            //inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
            return newdoc;
            //System.Windows.Forms.SendKeys.Send("{ENTER}");
        }
        private static string ConvertintoLatexMultiThread(string docFile)
        {
            Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            app.Documents.Open(docFile);
            //MessageBox.Show("Don't press any button. Only select the option 'MathJax Latex'. Otherwise ignore.");

            //MessageBox.Show("Conversion will start automatically. This process will take sometime. Please wait the document to close.");

            Thread p1 = new Thread(() => runCommand(app));
            p1.Start();

            Thread p2 = new Thread(() => ExecuteButton());
            p2.Start();
            p1.Join();
            p2.Join();
            string newdoc = Path.GetDirectoryName(docFile) + "\\" + Path.GetFileNameWithoutExtension(docFile) + "_math.docx";
            app.Documents.Save(newdoc);
            app.ActiveDocument.Close();
            app.Quit();
            return newdoc;
            //
        }
        private static System.Drawing.Rectangle searchBitmap(Bitmap smallBmp, Bitmap bigBmp, double tolerance)
        {
            BitmapData smallData =
              smallBmp.LockBits(new System.Drawing.Rectangle(0, 0, smallBmp.Width, smallBmp.Height),
                       System.Drawing.Imaging.ImageLockMode.ReadOnly,
                       System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            BitmapData bigData =
              bigBmp.LockBits(new System.Drawing.Rectangle(0, 0, bigBmp.Width, bigBmp.Height),
                       System.Drawing.Imaging.ImageLockMode.ReadOnly,
                       System.Drawing.Imaging.PixelFormat.Format24bppRgb);

            int smallStride = smallData.Stride;
            int bigStride = bigData.Stride;

            int bigWidth = bigBmp.Width;
            int bigHeight = bigBmp.Height - smallBmp.Height + 1;
            int smallWidth = smallBmp.Width * 3;
            int smallHeight = smallBmp.Height;

            System.Drawing.Rectangle location = System.Drawing.Rectangle.Empty;
            int margin = Convert.ToInt32(255.0 * tolerance);

            unsafe
            {
                byte* pSmall = (byte*)(void*)smallData.Scan0;
                byte* pBig = (byte*)(void*)bigData.Scan0;

                int smallOffset = smallStride - smallBmp.Width * 3;
                int bigOffset = bigStride - bigBmp.Width * 3;

                bool matchFound = true;

                for (int y = 0; y < bigHeight; y++)
                {
                    for (int x = 0; x < bigWidth; x++)
                    {
                        byte* pBigBackup = pBig;
                        byte* pSmallBackup = pSmall;

                        //Look for the small picture.
                        for (int i = 0; i < smallHeight; i++)
                        {
                            int j = 0;
                            matchFound = true;
                            for (j = 0; j < smallWidth; j++)
                            {
                                //With tolerance: pSmall value should be between margins.
                                int inf = pBig[0] - margin;
                                int sup = pBig[0] + margin;
                                if (sup < pSmall[0] || inf > pSmall[0])
                                {
                                    matchFound = false;
                                    break;
                                }

                                pBig++;
                                pSmall++;
                            }

                            if (!matchFound) break;

                            //We restore the pointers.
                            pSmall = pSmallBackup;
                            pBig = pBigBackup;

                            //Next rows of the small and big pictures.
                            pSmall += smallStride * (1 + i);
                            pBig += bigStride * (1 + i);
                        }

                        //If match found, we return.
                        if (matchFound)
                        {
                            location.X = x;
                            location.Y = y;
                            location.Width = smallBmp.Width;
                            location.Height = smallBmp.Height;
                            break;
                        }
                        //If no match found, we restore the pointers and continue.
                        else
                        {
                            pBig = pBigBackup;
                            pSmall = pSmallBackup;
                            pBig += 3;
                        }
                    }

                    if (matchFound) break;

                    pBig += bigOffset;
                }
            }

            bigBmp.UnlockBits(bigData);
            smallBmp.UnlockBits(smallData);

            return location;
        }
        private static void checkImage()
        {
            for (int i = 0; i < 100; i++)
            {
                string screenshotFolder = System.Windows.Forms.Application.StartupPath + "\\Temp\\Screenshot";
                if (!Directory.Exists(screenshotFolder))
                    Directory.CreateDirectory(screenshotFolder);
                Thread.Sleep(5000);
                Bitmap sourceImage = GetSreenshot();
                System.Drawing.Bitmap template = (Bitmap)Bitmap.FromFile(System.Windows.Forms.Application.StartupPath + "\\lib\\ok.png");
                System.Drawing.Image sourceImageImg = (System.Drawing.Image)sourceImage;
                // create template matching algorithm's instance
                // (set similarity threshold to 92.1%)

                var templateImage24bpp = ConvertTo24bpp(System.Drawing.Image.FromFile(System.Windows.Forms.Application.StartupPath + "\\lib\\ok.png"));
                var sourceImage24bpp = ConvertTo24bpp(sourceImageImg);

                ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(0.921f);
                // find all matchings with specified above similarity

                System.Drawing.Rectangle location = searchBitmap(sourceImage24bpp, templateImage24bpp, 35);

                TemplateMatch[] matchings = tm.ProcessImage(sourceImage24bpp, templateImage24bpp);
                // highlight found matchings

                BitmapData data = sourceImage.LockBits(
                     new System.Drawing.Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                     ImageLockMode.ReadWrite, sourceImage.PixelFormat);
                foreach (TemplateMatch m in matchings)
                {

                    AForge.Imaging.Drawing.Rectangle(data, m.Rectangle, System.Drawing.Color.White);

                    MessageBox.Show(m.Rectangle.Location.ToString());
                    // do something else with matching
                }

                if (matchings.Length > 0)
                {
                    InputSimulator inputSimulator = new InputSimulator();
                    inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                    break;
                }
                else
                {

                }
            }
        }
        private static Bitmap GetSreenshot()
        {
            Bitmap bm = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(bm);
            g.CopyFromScreen(0, 0, 0, 0, bm.Size);
            return bm;
        }
        public static Bitmap ConvertTo24bpp(System.Drawing.Image img)
        {
            var bmp = new Bitmap(img.Width, img.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var gr = Graphics.FromImage(bmp))
                gr.DrawImage(img, new System.Drawing.Rectangle(0, 0, img.Width, img.Height));
            return bmp;
        }
        private static void runCommand(Application app)
        {
            app.Run("MTCommand_ConvertEqns");
        }
        private static void ExecuteButton()
        {
            Thread.Sleep(3000);
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
            checkImage();
            //System.Windows.Forms.SendKeys.Send("{ENTER}");
        }
        private static string getmaths(string inputDoc, string tempFolderApp)
        {
            string newdocPath = inputDoc;
            Dictionary<string, string> maths = new Dictionary<string, string>();
            string outFile = Path.GetDirectoryName(inputDoc) + "\\" + Path.GetFileNameWithoutExtension(inputDoc) + "_math" + Path.GetExtension(inputDoc);
            File.Copy(inputDoc, outFile);
            unzipfile(outFile);
            string inputDocFolder = Path.GetDirectoryName(outFile) + "\\" + Path.GetFileNameWithoutExtension(outFile);
            string[] documentFile = Directory.GetFiles(inputDocFolder, "document.xml", SearchOption.AllDirectories);
            string[] reldocumentFile = Directory.GetFiles(inputDocFolder, "document.xml.rels", SearchOption.AllDirectories);
            string relFile = "";
            if (reldocumentFile.Length > 0)
            {
                relFile = reldocumentFile[0];
            }
            if (reldocumentFile.Length > 0)
            {
                string textFile = File.ReadAllText(reldocumentFile[0]).Replace(" xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"", "");
                string documentFileText = File.ReadAllText(documentFile[0]).Replace(" r:id=", " r_id=");
                File.WriteAllText(documentFile[0], documentFileText);
                XmlDocument doc = new XmlDocument();
                doc.Load(documentFile[0]);
                HtmlAgilityPack.HtmlDocument reldoc = new HtmlAgilityPack.HtmlDocument();
                reldoc.LoadHtml(textFile);
                HtmlNodeCollection wmfNodes = reldoc.DocumentNode.SelectNodes("//*[contains(@target,'.wmf')]");
                if (wmfNodes.Count > 0)
                {
                    foreach (HtmlNode mathNode in wmfNodes)
                    {
                        string target = mathNode.Attributes["target"].Value;
                        string id = mathNode.Attributes["id"].Value;
                        string wmfFile = null;
                        string[] imagePath = Directory.GetFiles(inputDocFolder, Path.GetFileName(target), SearchOption.AllDirectories);
                        if (imagePath.Length > 0)
                        {
                            wmfFile = imagePath[0];
                        }
                        XmlNodeList objectNodes = doc.DocumentElement.SelectNodes("//*[@r_id='" + id + "']");
                        if (objectNodes.Count > 0)
                        {
                            foreach (XmlNode objNode in objectNodes)
                            {
                                if (!maths.ContainsKey("[" + id + "]"))
                                {
                                    string mathml = CreateMath(objNode, wmfFile, documentFile[0], tempFolderApp, id, doc);
                                    maths.Add("[" + id + "]", mathml);
                                }
                            }
                        }
                    }
                }
                doc.Save(documentFile[0]);
                textFile = File.ReadAllText(documentFile[0]).Replace("_separator_", ":").Replace("</temp>", "").Replace("<temp>", "").Replace("_spt_", ":").Replace("r_id", "r:id");
                File.WriteAllText(documentFile[0], textFile);
                newdocPath = zipFile(inputDocFolder, Path.GetDirectoryName(inputDocFolder), inputDoc);
            }
            SharedObjects.Maths = maths;
            return newdocPath;
        }
        public static string zipFile(string f_path, string movableDir, string docFile)
        {
            string newdocFile = docFile;
            string DirName = Path.GetFileName(f_path);
            using (Ionic.Zip.ZipFile zip1 = new Ionic.Zip.ZipFile())
            {
                DirectoryInfo dir_all = new DirectoryInfo(f_path);
                DirectoryInfo[] dir_all1 = dir_all.GetDirectories("*.*");
                FileInfo[] Files = dir_all.GetFiles("*.*");
                string fileName = dir_all.Name + ".docx";
                string newfileName = dir_all.Name + "_ready.docx";
                //foreach (FileInfo file_all in Files)
                //{
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile(dir_all.FullName + "\\" + fileName))
                {
                    foreach (DirectoryInfo directory_all in dir_all1)
                    {
                        DirectoryInfo dir_al2 = new DirectoryInfo(directory_all.FullName);
                        DirectoryInfo[] dir_all3 = dir_al2.GetDirectories("*.*");
                        FileInfo[] Files1 = dir_al2.GetFiles("*.*");
                        zip.AddDirectory(dir_all.FullName + "\\" + directory_all.Name, directory_all.Name);
                    }
                    foreach (FileInfo directory_all in Files)
                    {
                        zip.AddFile(dir_all.FullName + "\\" + directory_all.Name, "");
                    }
                    zip.Save(f_path + "\\" + dir_all.Name + ".zip");
                }
                // File.Delete(f_path + "\\" + fileName);
                File.Move(f_path + "\\" + dir_all.Name + ".zip", f_path + "\\" + fileName);
                // }
                string str = f_path + "\\" + fileName;
                string dir = f_path + "\\";
                newdocFile = movableDir + "\\" + newfileName;
                try
                {
                    // File.Delete(movableDir + "\\" + fileName);
                    File.Move(str, movableDir + "\\" + newfileName);
                }
                catch (Exception ex)
                {
                    File.Move(str, docFile);
                }
                //try
                //{
                //    Directory.Delete(f_path + "\\", true);
                //}
                //catch (Exception ex)
                //{
                //    Directory.Delete(f_path, true);
                //}
            }
            return newdocFile;
        }
        private static string CreateMath(XmlNode objectNode, string wmfFile, string documentxml, string tempFolderApp, string rid, XmlDocument maindoc)
        {
            string math = "";
            string imageFile = "";
            string outXml = objectNode.OuterXml.Replace("v:", "");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(outXml);

            XmlNodeList chNodes = doc.GetElementsByTagName("imagedata");
            if (chNodes != null)
            {
                XmlNode imagedata = chNodes[0];
                string image = wmfFile;
                if (Path.GetExtension(image) == ".wmf")
                {
                    string folderPath = Path.GetDirectoryName(documentxml);
                    imageFile = tempFolderApp + "\\Math\\" + Path.GetFileNameWithoutExtension(image) + ".jpg";
                    if (!Directory.Exists(tempFolderApp + "\\Math"))
                        Directory.CreateDirectory(tempFolderApp + "\\Math");
                    WMFtoImage(image, imageFile);
                    string mathresult = GenerateMathfromImage(imageFile);
                    string mathMLJSON = Path.GetDirectoryName(imageFile) + "\\MathML\\" + Path.GetFileNameWithoutExtension(imageFile) + ".json";
                    if (File.Exists(mathMLJSON))
                    {
                        using (StreamReader reader = new StreamReader(mathMLJSON))
                        {
                            string json = reader.ReadToEnd();
                            int mathCount = Regex.Matches(json, "<math(.+?)</math>").Count;
                            if (mathCount > 0)
                            {
                                string mathValue = Regex.Matches(json, "<math(.+?)</math>")[0].Value;
                                Regex reg = new Regex(@"\\u(....)");
                                mathValue = reg.Replace(mathValue, "&#x$1;").Replace("\\" + "\"", "\"");
                                math = mathValue;

                                if (math.Length > 3)
                                {
                                    XmlNode wpNode = maindoc.CreateElement("temp");
                                    string textAdd = "<w_separator_t>[" + rid + "]</w_separator_t>";
                                    wpNode.InnerXml = textAdd;
                                    XmlNode mainObjectNode = objectNode.ParentNode;
                                    for (int i = 0; i < 10; i++)
                                    {
                                        if (mainObjectNode.Name == "w:object")
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            mainObjectNode = mainObjectNode.ParentNode;
                                        }
                                    }
                                    mainObjectNode.ParentNode.ReplaceChild(wpNode, mainObjectNode);
                                }
                            }
                        }
                    }
                }
            }
            return math;
        }
        public static void WMFtoImage(string wmfFile, string imageFile)
        {
            System.Drawing.Image i = System.Drawing.Image.FromFile(wmfFile, true);
            Bitmap b = new Bitmap(i);
            Graphics g = Graphics.FromImage(b);
            g.Clear(System.Drawing.Color.White);
            g.DrawImage(i, 0, 0, i.Width, i.Height);
            b.Save(imageFile, ImageFormat.Jpeg);
        }
        public static string GenerateMathfromImage(string imagePath)
        {
            executebatch exeObj = new executebatch();
            string result = string.Empty;
            try
            {
                string command = System.Windows.Forms.Application.StartupPath + "\\lib\\Image2Math\\Image2MathML.exe";
                string argument = " " + imagePath;
                result = exeObj.executeCommand(command, argument);

            }
            catch (Exception e)
            {

            }

            return result;
        }
        private static void ConvertDocx(string docFile, string panhtml)
        {
            string odtFile = Path.GetDirectoryName(docFile) + "\\" + Path.GetFileNameWithoutExtension(docFile) + ".html";
            DocumentCore dc = DocumentCore.Load(docFile);
            dc.Save(odtFile, SaveOptions.HtmlFlowingDefault);
            string newodt = Path.GetDirectoryName(odtFile) + "\\" + Path.GetFileNameWithoutExtension(odtFile) + "_1" + Path.GetExtension(odtFile);
            File.Copy(odtFile, newodt,true);
            Regex reg = new Regex(@"&lt;listitem(.+?)&gt;");
            
            var results = reg.Matches(File.ReadAllText(odtFile));
            string htmlText = File.ReadAllText(odtFile).Replace("&lt;/listitem&gt;", "");
            foreach (Match match in results)
            {
                string matchvalue = match.Value;
                //string replacevalue = match.Value.Replace("&lt;","<").Replace("&gt;", ">").Replace("&quot;", "\"");
                htmlText = htmlText.Replace(matchvalue, "");
            }

            File.WriteAllText(odtFile, htmlText);

            var resultsNew = reg.Matches(File.ReadAllText(newodt));
            string htmlTextNew = File.ReadAllText(newodt).Replace("&lt;/listitem&gt;", "</listitem>");
            foreach (Match match in resultsNew)
            {
                string matchvalue = match.Value;
                string replacevalue = match.Value.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"");
                htmlTextNew = htmlTextNew.Replace(matchvalue, replacevalue);
            }

            File.WriteAllText(newodt, htmlTextNew);

            var panresults = reg.Matches(File.ReadAllText(panhtml));
            string panhtmlText = File.ReadAllText(panhtml).Replace("&lt;/listitem&gt;", "</listitem>");
            foreach (Match match in panresults)
            {
                string matchvalue = match.Value;
                string replacevalue = match.Value.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"");
                panhtmlText = panhtmlText.Replace(matchvalue, replacevalue);
            }

            File.WriteAllText(panhtml, panhtmlText);
            panhtmlText = File.ReadAllText(panhtml);
            HtmlAgilityPack.HtmlDocument tDocNew = new HtmlAgilityPack.HtmlDocument();
            tDocNew.Load(newodt);
            HtmlAgilityPack.HtmlNodeCollection listNodes = tDocNew.DocumentNode.SelectNodes("//listitem");
            if (listNodes != null)
            {
                foreach (HtmlNode node in listNodes.ToList())
                {
                    string id = node.Attributes["id"].Value;
                    HtmlNode parentNode = node.ParentNode;
                    for (int i = 0; i < 5; i++)
                    {
                        if (parentNode != null)
                        {
                            if ((parentNode.Name == "p") || (parentNode.Name == "li"))
                            {
                                if (parentNode.Name == "p")
                                {
                                    node.ParentNode.RemoveChild(node);
                                    string innerText = parentNode.InnerText.Replace("\r", "").Replace("\n", "").Replace("\t(", "(").Replace("\t(", "(").Replace("\t(", "(").Replace("\t(", "(").Replace("\t(", "(").Replace("\t", "");
                                    if (innerText.Trim().Length > 0)
                                    {
                                        panhtmlText = panhtmlText.Replace("<listitem id=" + "\"" + id + "\"" + ">", "<span class=" + "\"" + "tabbed" + "\"" + ">" + innerText + "</span>\t" + "<listitem id=" + "\"" + id + "\"" + ">");
                                    }
                                }
                                else
                                {
                                    //node.ParentNode.
                                }
                                break;
                            }
                            else { parentNode = parentNode.ParentNode; }
                        }
                        else { break; }
                    }
                }
            }

            Regex reg1 = new Regex(@"<listitem(.+?)>");
            var resultsPan = reg1.Matches(panhtmlText);
            panhtmlText = panhtmlText.Replace("</listitem>", "");
            foreach (Match match in resultsPan)
            {
                string matchvalue = match.Value;
                //string replacevalue = match.Value.Replace("&lt;","<").Replace("&gt;", ">").Replace("&quot;", "\"");
                panhtmlText = panhtmlText.Replace(matchvalue, "");
            }

            File.WriteAllText(panhtml, panhtmlText);

            HtmlAgilityPack.HtmlDocument tDocHtml = new HtmlAgilityPack.HtmlDocument();
            tDocHtml.Load(panhtml);

            HtmlAgilityPack.HtmlNodeCollection spanTabbedNodes = tDocHtml.DocumentNode.SelectNodes("//span[@class='tabbed']");
            if (spanTabbedNodes != null)
            {
                foreach (HtmlNode node in spanTabbedNodes)
                {
                    HtmlNode parentNode = node.ParentNode;
                    for (int i = 0; i < 5; i++)
                    {
                        if (((parentNode.Name == "ol") && (parentNode.Attributes.Contains("type"))) || ((parentNode.Name == "ul") && (parentNode.Attributes.Contains("type"))))
                        {
                            //parentNode.Attributes.Remove("type");
                            parentNode.Attributes["type"].Value= "none";
                            break;
                        }
                        else
                        {
                            if ((parentNode.Name == "ul") || (parentNode.Name == "ol")) { }
                            else
                            {
                                parentNode = parentNode.ParentNode;
                            }
                        }
                    }
                }
                foreach (HtmlNode node in spanTabbedNodes.ToList())
                {
                    //node.ParentNode.RemoveChild(node);
                }
            }
            tDocHtml.Save(panhtml);


            HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
            tDoc.Load(odtFile);

            HtmlAgilityPack.HtmlNodeCollection spanNodes = tDoc.DocumentNode.SelectNodes("//span[@style]");
            if (spanNodes != null)
            {
                foreach (HtmlNode node in spanNodes)
                {
                    string innText = node.InnerHtml;
                    string[] styles = node.Attributes["style"].Value.Split(';');
                    if (styles.Length > 0)
                    {
                        string styleStr = "";
                        foreach (string style in styles)
                        {
                            if (style.Contains("font-weight:"))
                            {
                                innText = "<strong>" + innText + "</strong>";
                            }
                            if (style.Contains("font-style:"))
                            {
                                innText = "<em>" + innText + "</em>";
                            }
                            if (style.Contains(" color:"))
                            {
                                styleStr = styleStr + style.Trim() + ";";
                            }
                        }
                        if (styleStr.Length > 0)
                        {
                            node.Attributes["style"].Value = styleStr;
                        }
                        node.InnerHtml = innText;
                    }
                }
            }
            HtmlAgilityPack.HtmlNodeCollection styleNodes = tDoc.DocumentNode.SelectNodes("//*[@style]");
            if (styleNodes != null)
            {
                foreach (HtmlNode node in styleNodes)
                {
                    if (node.Name != "span")
                    {
                        //node.Attributes.Remove("style");
                    }
                }
            }
            
            HtmlAgilityPack.HtmlNodeCollection tableNodes = tDoc.DocumentNode.SelectNodes("//table");
            if (tableNodes != null)
            {
                int num = 1;
                foreach (HtmlNode node in tableNodes)
                {
                    HtmlNode preNode = node.PreviousSibling;
                    for (int i = 0; i < 3; i++)
                    {
                        if (preNode == null)
                        {
                            break;
                        }
                        if (preNode.Name != "#text")
                        {
                            if (preNode.InnerText.Trim().StartsWith("Table "))
                            {
                                string caption = preNode.InnerHtml;
                                HtmlNode capNode = tDoc.CreateElement("caption");
                                capNode.InnerHtml = caption;
                                node.InnerHtml = capNode.OuterHtml + node.InnerHtml;
                                //preNode.ParentNode.RemoveChild(preNode);
                                break;
                            }
                            else { preNode = preNode.PreviousSibling; }
                        }
                        else { preNode = preNode.PreviousSibling; }
                    }
                }
            }
            HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//table");
            if (tdNodes != null)
            {
                int num = 1;
                foreach (HtmlNode htmlNode in tdNodes)
                {
                    string innerText = htmlNode.InnerText.Replace("\r", "").Replace("\n", "").Replace("\t", "").Trim().Replace("Â","").Trim();
                    if (innerText.Length > 10)
                    {
                        htmlNode.Attributes.Add("id", "Table_" + num);
                    }
                    htmlNode.InnerHtml = htmlNode.InnerHtml.Replace("\\[", "$").Replace("\\]", "$");
                    //if (SharedObjects.mathEnabled == true)
                    {
                        string MathFunctionsEnable = ConfigurationManager.AppSettings.Get("MathFunctionsEnable");
                        if (MathFunctionsEnable.ToLower().Trim() == "true")
                        {
                            MatchCollection matches = Regex.Matches(htmlNode.InnerHtml, @"&lt;math&gt;(.+?)&lt;/math&gt;", RegexOptions.Singleline);
                            foreach (Match match in matches)
                            {
                                string matchText = match.Value;
                                if ((match.Value.Contains("</span>")) && (match.Value.Contains("<span")))
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                                    htmlDocument.LoadHtml(match.Value);
                                    matchText = htmlDocument.DocumentNode.InnerText;
                                }
                                string latex = matchText.Replace("\r", "").Replace("</p>\n<p>", "").Replace("&lt;math&gt;", "").Replace("&lt;/math&gt;", "").Trim().Trim('$').Replace("&amp;", "&");
                                latex = "<p>" + latex + "</p>";
                                latex = latex.Replace("<p>\\[", "").Replace("\\]</p>", "").Replace("<p>", "").Replace("</p>", "");
                                string correctedValue = "<span id=" + "\"" + "mathjax" + "\"" + " latex=" + "\"" + latex + "\"" + "></span>";
                                if ((match.Value.Contains("</span>")) && (match.Value.Contains("<span")))
                                {
                                    correctedValue = "</span>" + correctedValue + "<span>";
                                }
                                htmlNode.InnerHtml = htmlNode.InnerHtml.Replace(match.Value, correctedValue);
                            }
                        }
                        else
                        {
                            MatchCollection matches = Regex.Matches(htmlNode.InnerHtml, @"\$(.+?)\$", RegexOptions.Singleline);
                            foreach (Match match in matches)
                            {
                                string matchText = match.Value;
                                if ((match.Value.Contains("</span>")) && (match.Value.Contains("<span")))
                                {
                                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                                    htmlDocument.LoadHtml(match.Value);
                                    matchText = htmlDocument.DocumentNode.InnerText;
                                }
                                string latex = matchText.Replace("\r", "").Replace("</p>\n<p>", "").Replace("&lt;math&gt;", "").Replace("&lt;/math&gt;", "").Trim().Trim('$').Replace("&amp;", "&");
                                latex = "<p>" + latex + "</p>";
                                latex = latex.Replace("<p>\\[", "").Replace("\\]</p>", "").Replace("<p>", "").Replace("</p>", "");
                                string correctedValue = "<span id=" + "\"" + "mathjax" + "\"" + " latex=" + "\"" + latex + "\"" + "></span>";
                                if ((match.Value.Contains("</span>")) && (match.Value.Contains("<span")))
                                {
                                    correctedValue = "</span>" + correctedValue + "<span>";
                                }
                                htmlNode.InnerHtml = htmlNode.InnerHtml.Replace(match.Value, correctedValue);
                            }
                        }
                        //if (SharedObjects.Maths.Count > 0)
                        //{
                        //    foreach (KeyValuePair<string, string> item in SharedObjects.Maths)
                        //    {
                        //        string id = item.Key;
                        //        string math = item.Value;
                        //        string mathml = item.Value.Split('|')[0];
                        //        if (mathml.Trim().Contains("<math"))
                        //        {
                        //            htmlNode.InnerHtml = htmlNode.InnerHtml.Replace(id, math.Split('|')[0]);

                        //        }
                        //    }
                        //}
                        if (SharedObjects.RubyList.Count > 0)
                        {
                            foreach (KeyValuePair<string, string> item in SharedObjects.RubyList)
                            {
                                string id = item.Key;
                                string ruby = item.Value;
                                ruby = ruby.Replace("w:", "");
                                HtmlAgilityPack.HtmlDocument tiDoc = new HtmlAgilityPack.HtmlDocument();
                                tiDoc.OptionWriteEmptyNodes = true;
                                tiDoc.LoadHtml(ruby);
                                HtmlNode rtNode = tiDoc.DocumentNode.SelectSingleNode("//rt");
                                HtmlNode tNode = tiDoc.DocumentNode.SelectSingleNode("//rt//t");
                                rtNode.InnerHtml = tNode.InnerHtml;
                                HtmlNode rtbNode = tiDoc.DocumentNode.SelectSingleNode("//rubybase");
                                HtmlNode tbNode = tiDoc.DocumentNode.SelectSingleNode("//rubybase//t");
                                rtbNode.InnerHtml = tbNode.InnerHtml;

                                string newTag = "<ruby>" + rtbNode.InnerHtml + rtNode.OuterHtml + "</ruby>";
                                htmlNode.InnerHtml = htmlNode.InnerHtml.Replace(id, newTag.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", ""));
                            }
                        }
                    }
                    htmlNode.InnerHtml = htmlNode.InnerHtml.Replace("[color]", "</span>");
                    Regex regcolor = new Regex(@"\[color\((.+?)\)\]");
                    htmlNode.InnerHtml = regcolor.Replace(htmlNode.InnerHtml, "<span style=\"color:$1" + "\"" + ">");
                    if (innerText.Length > 10)
                    {
                        SharedObjects.TablesStyled.Add("Table_" + num, htmlNode);
                        num++;
                    }
                }
                tDoc.Save(odtFile);
            }
        }
        private static void GetAllTables(string htmlFile)
        {   
            List<HtmlNode> tables = new List<HtmlNode>();
            HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
            tDoc.OptionWriteEmptyNodes = true;
            tDoc.Load(htmlFile);
            HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//table");
            if (tdNodes != null)
            {
                int num = 1;
                foreach (HtmlNode htmlNode in tdNodes)
                {
                    htmlNode.Attributes.Add("id", "Table_" + num);

                    SharedObjects.TablePanDoc.Add("Table_" + num, htmlNode);


                    num++;
                }
                tDoc.Save(htmlFile);
            }
        }
        private static void unzipfile(string f_path)
        {
            string dirName = Path.GetDirectoryName(f_path);
            string file_Names = Path.GetFileNameWithoutExtension(f_path);
            if (Path.GetExtension(f_path) == ".docx")
            {
                File.Move(f_path, dirName + "\\" + file_Names + ".zip");
            }
            using (Ionic.Zip.ZipFile zip1 = new Ionic.Zip.ZipFile(dirName + "\\" + file_Names + ".zip"))
            {
                Directory.CreateDirectory(dirName + "\\" + file_Names);
                zip1.ExtractAll(dirName + "\\" + file_Names);
            }
            //File.Delete(dirName + "\\" + file_Names + ".zip");


        }
        private static void ConvertDocxToODT(object Sourcepath, object TargetPath)
        {
            Microsoft.Office.Interop.Word._Application newApp = new Microsoft.Office.Interop.Word.Application();
            newApp.Visible = true;
            Microsoft.Office.Interop.Word.Documents d = newApp.Documents;
            //object Unknown = Type.Missing;
            object Unknown = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document od = d.Open(ref Sourcepath, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown, ref Unknown);
            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatOpenDocumentText;
            newApp.ActiveDocument.SaveAs(ref TargetPath, ref format,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown);

            newApp.Documents.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            newApp.Quit();

        }
        private static string CleanIndexHTML(string indexhtml)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(indexhtml);
            HtmlNodeCollection scripts = hDoc.DocumentNode.SelectNodes("//script");
            if (scripts != null)
            {
                foreach (HtmlNode node in scripts.ToList())
                {
                    if ((node.Attributes.Contains("src")) && ((node.Attributes["src"].Value.Contains("data/page_")) || (node.Attributes["src"].Value.Contains("navigation.js")) || (node.Attributes["src"].Value.Contains("Toc_001.js")) || (node.Attributes["src"].Value.Contains("toc_001.js"))))
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                    if (node.InnerHtml.Contains("window.page_"))
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
            string outHtml = hDoc.DocumentNode.OuterHtml.Replace("\r", "").Replace("\n", "").Replace("</script>", "</script>\n");
            return hDoc.DocumentNode.OuterHtml;
        }
        private static List<Navigation.Page> LoadDatatoJSON(Dictionary<string, List<HtmlNode>> screensList, string output_path, string imagesDir, string docName, Dictionary<string, GlossaryItem> glossaryList)
        {
            StringBuilder logs = new StringBuilder();
            List<Navigation.Page> pages = new List<Navigation.Page>();
            if (screensList != null)
            {
                int pno = 1;

                foreach (KeyValuePair<string, List<HtmlNode>> screenItem in screensList)
                {   
                    Navigation.Page NavPage = new Navigation.Page();
                    string ScrTitle = screenItem.Key.Replace("\r", "").Replace("\n", "");
                    Regex REG = new Regex(@"(.+?)creen(.+?) (.+)", RegexOptions.IgnoreCase);
                    string PageTitle = REG.Replace(ScrTitle, "$3").Trim('-').Trim(':').Trim().Trim('–').Trim();
                    string pageName = GetPageName(pno);
                    Page.Root page = new Page.Root();
                    page.id = pageName;
                    page.Page_Title = PageTitle;
                    NavPage.id = Convert.ToInt32(pageName.Replace("page_00","").Replace("page_0", "").Replace("page_", ""));
                    NavPage.filename = pageName;
                    pages.Add(NavPage);
                    List<HtmlNode> components = screenItem.Value;
                    List<Component> componentsArr = new List<Component>();
                    Toc.Panel panel = new Toc.Panel();
                    panel.Panel_Heading = PageTitle;
                    if (components != null)
                    {
                        StringBuilder strpanels = new StringBuilder();
                        int j = 1;
                        foreach (HtmlNode node in components)
                        {
                            string templateStr = "<span id='toc-navigate' navigateTo='[TemplateID]'>[TITLE]</span>";
                            Component cmp = GetComponentData(node, imagesDir, page);
                            if (SharedObjects.popupList != null)
                            {
                                if (SharedObjects.popupList.Count > 0)
                                {
                                    foreach (KeyValuePair<string, string> item in SharedObjects.popupList)
                                    {
                                        if (cmp.PopupDataID == item.Value.Trim())
                                        {
                                            cmp.TemplateID = "";
                                        }
                                    }
                                }
                            }
                            if (j == 1)
                            {
                                if ((cmp.Title == null) || (cmp.Title.Contains("<temp")) || (cmp.Title == ""))
                                {
                                    panel.Panel_Heading = templateStr.Replace("[TemplateID]", page.id).Replace("[TITLE]", PageTitle);
                                    //cmp.Page_Title = PageTitle;
                                }
                                else
                                {
                                    panel.Panel_Heading = templateStr.Replace("[TemplateID]", page.id).Replace("[TITLE]", PageTitle);
                                    //cmp.Page_Title = PageTitle;
                                }
                            }
                            componentsArr.Add(cmp);
                            if (j == components.Count)
                            {
                                if ((cmp.Title == null) || (cmp.Title.Contains("<temp")) || (cmp.Title == "")) { }
                                else
                                {
                                    strpanels.Append(templateStr.Replace("[TemplateID]", page.id).Replace("[TITLE]", cmp.Title));
                                }
                            }
                            else
                            {
                                if ((cmp.Title == null) || (cmp.Title.Contains("<temp")) || (cmp.Title == "")) { }
                                else
                                {
                                    strpanels.Append(templateStr.Replace("[TemplateID]", page.id).Replace("[TITLE]", cmp.Title) + "<br/>");
                                }
                            }
                            j++;
                        }
                        panel.Panel_RevealText = strpanels.ToString();
                    }
                    if (SharedObjects.GlossaryEnable == false)
                    {
                        if (pno != SharedObjects.GlossaryPage)
                        {
                            SharedObjects.Panels.Add(panel);
                        }
                        else { SharedObjects.GlossaryPageId = pageName; }
                    }
                    else { SharedObjects.Panels.Add(panel); }
                    page.components = componentsArr;
                    string log = validation(page, pageName);
                    logs.Append(log + Environment.NewLine);
                    pno++;
                    //string json = JsonConvert.SerializeObject(page, Formatting.Indented);
                    var json = JsonConvert.SerializeObject(page, Newtonsoft.Json.Formatting.None,
new JsonSerializerSettings
{
    NullValueHandling = NullValueHandling.Ignore
});
                    json = Replacement(json);

                    //json = ReplacementGlossary(json, glossaryList);
                    string jsonFormatted = "var " + pageName + " = " + json.ToString();
                    try
                    {
                        jsonFormatted = "var " + pageName + " = " + JValue.Parse(json).ToString(Newtonsoft.Json.Formatting.Indented);
                    }
                    catch (Exception)
                    {
                    }

                    // json=CleanNodes(json);
                    File.WriteAllText(output_path + "\\" + pageName + ".js", jsonFormatted.Replace(" type='none'"," style='list-style-type:none'"));
                    Console.WriteLine("Screen: " + pno + "\t\tCreated." + Environment.NewLine + jsonFormatted);
                }
            }
            if (logs.Length > 0)
            {
                File.WriteAllText(output_path + "\\" + docName + ".log", logs.ToString());
            }
            return pages;
        }
        private static string ReplacementGlossary(string json, Dictionary<string, string> glossaryList)
        {
            MatchCollection matches = Regex.Matches(json, @"<strong>(.+?)</strong>");

            // Use foreach-loop.
            foreach (Match match in matches)
            {
                foreach (Capture capture in match.Captures)
                {
                    string term = capture.Value.Trim();
                    HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                    hDoc.LoadHtml(term);
                    string termText = hDoc.DocumentNode.InnerText.Trim();
                    if (glossaryList.ContainsKey(termText))
                    {
                        string glossDef = "<a data-glossary='" + glossaryList[termText] + "' class='global-glossary'>" + termText + "</a>";
                        json = json.Replace(term, glossDef);
                    }
                }
            }
            return json;
        }
        private static string ReplacementPopup(string json, Dictionary<string, string> glossaryList)
        {
            MatchCollection matches = Regex.Matches(json, @"<strong>(.+?)</strong>");

            // Use foreach-loop.
            foreach (Match match in matches)
            {
                foreach (Capture capture in match.Captures)
                {
                    string term = capture.Value.Trim();
                    HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                    hDoc.LoadHtml(term);
                    string termText = hDoc.DocumentNode.InnerText.Trim();
                    if (glossaryList.ContainsKey(termText))
                    {
                        string glossDef = "<a data-glossary='" + glossaryList[termText] + "' class='global-glossary'>" + termText + "</a>";
                        json = json.Replace(term, glossDef);
                    }
                }
            }
            return json;
        }
        private static string Replacement(string json)
        {
            json = json.Replace(">\\r\\n<", "><").Replace("AltText", "Alt-text")
                .Replace("MediaPosition", "Media-position")
                .Replace("TeacherOnly", "Teacher-only")
                .Replace("TemplateDescription", "Template-Description")
                 .Replace("Slides", "sliderData")
                 .Replace("glossaryterm", "glossaryTerm")
                  .Replace("popdata", "popData")
                .Replace("PanelRevealText", "Panel_RevealText")
                .Replace("\"" + "Notes" + "\"" + ": " + "\"" + "Null" + "\"" + ",", "")
                 .Replace("\"" + "Notes" + "\"" + ":" + "\"" + "Null" + "\"" + ",", "")
                 .Replace("Transcript_txt", "Transcript").Replace("\"Texts\": []", "").Replace("\"Texts\":[]", "")
                .Replace("SRT_VTT", "SRT/VTT")
                .Replace("[PIPESTART]", "\\"+"\"").Replace("[PIPEEND]", "\\" + "\"");
            json = json.Replace("<temp>", "").Replace("</temp>", "")
                .Replace("Main-texts", "MainText").Replace("Main_Texts", "MainText").Replace("</temp>", "").Replace("controls=''", "controls").Replace("'=''", "");
            return json;
        }
        private static string ReplacementToc(string json)
        {
            json = json.Replace(">\\r\\n<", "><")
                .Replace("TeacherOnly", "Teacher-only")
                .Replace("TemplateDescription", "Template-Description");
            return json;
        }
        private static Component GetComponentData(HtmlNode node, string imagesDir, Page.Root page)
        {
            Component cmp = new Component();
            string table = node.OuterHtml;
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(table);
            HtmlAgilityPack.HtmlNodeCollection trNodes = hDoc.DocumentNode.SelectNodes("//tr");
            if (trNodes != null)
            {
                trNodes = ImplementIdFunc(hDoc, trNodes);
                foreach (HtmlNode tNode in trNodes.ToList())
                {
                    string id = tNode.Attributes["id"].Value;
                    if (SharedObjects.idNodes == null)
                    {
                        string trText = tNode.OuterHtml;
                        HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                        tDoc.LoadHtml(trText);
                        HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                        if (tdNodes != null)
                        {
                            string term = tdNodes[0].InnerHtml;
                            string def = tdNodes[1].InnerHtml;
                            cmp = GetComponent(cmp, term, def, tNode, imagesDir, page);
                        }
                    }
                    else
                    {
                        if (!SharedObjects.idNodes.ContainsKey(id))
                        {
                            string trText = tNode.OuterHtml;
                            HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                            tDoc.LoadHtml(trText);
                            HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                            if (tdNodes != null)
                            {
                                string term = tdNodes[0].InnerHtml;
                                string def = tdNodes[1].InnerHtml;
                                cmp = GetComponent(cmp, term, def, tNode, imagesDir, page);
                            }
                        }
                    }
                }
            }
            return cmp;
        }
        private static Meta GetNavigationData(HtmlNode node)
        {
            Meta meta = new Meta();
            string table = node.OuterHtml;
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(table);
            HtmlAgilityPack.HtmlNodeCollection trNodes = hDoc.DocumentNode.SelectNodes("//tr");
            if (trNodes != null)
            {
                trNodes = ImplementIdFunc(hDoc, trNodes);
                foreach (HtmlNode tNode in trNodes.ToList())
                {
                    string id = tNode.Attributes["id"].Value;
                    if (SharedObjects.idNodes == null)
                    {
                        string trText = tNode.OuterHtml;
                        HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                        tDoc.LoadHtml(trText);
                        HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                        if (tdNodes != null)
                        {
                            string term = tdNodes[0].InnerHtml;
                            string def = tdNodes[1].InnerHtml;
                            meta = GetNavigationData(meta, term, def, tNode, null);
                        }
                    }
                    else
                    {
                        if (!SharedObjects.idNodes.ContainsKey(id))
                        {
                            string trText = tNode.OuterHtml;
                            HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                            tDoc.LoadHtml(trText);
                            HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                            if (tdNodes != null)
                            {
                                string term = tdNodes[0].InnerHtml;
                                string def = tdNodes[1].InnerHtml;
                                meta = GetNavigationData(meta, term, def, tNode, null);
                            }
                        }
                    }
                }
            }
            return meta;
        }

        //private static string GetDuration(string url)
        //{
        //    //WebClient myDownloader = new WebClient();
        //    //myDownloader.Encoding = System.Text.Encoding.UTF8;

        //    //string jsonResponse = myDownloader.DownloadString(
        //    //"https://www.googleapis.com/youtube/v3/videos?id=" + yourvideoID + "&key="
        //    //+ youtubekey + "&part=contentDetails");
        //    //dynamic dynamicObject = Json.Decode(jsonResponse);
        //    //string tmp = dynamicObject.items[0].contentDetails.duration;
        //    //var Duration = Convert.ToInt32
        //    //(System.Xml.XmlConvert.ToTimeSpan(tmp).TotalSeconds);
        //}
        private static string RemoveFormatting(string def)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(def);
            string data = hDoc.DocumentNode.InnerHtml.Replace("<b>", "").Replace("</b>", "")
                .Replace("<strong>", "").Replace("</strong>", "").Replace("<i>", "").Replace("</i>", "")
                .Replace("<em>", "").Replace("</em>", "");
            return data;
        }
        private static void WriteTocJson(string assetFolderPath)
        {
            Toc.Root toc = new Toc.Root();
            //toc.id = "toc_001";
            //SharedObjects.component.TemplateID = "nav_001";
            //SharedObjects.component.Panels = SharedObjects.Panels;
            //SharedObjects.component.TemplateName = "Table of Content";
            //List<Toc.Component> components = new List<Toc.Component>();
            //components.Add(SharedObjects.component);
            //toc.components = components;
            var json = JsonConvert.SerializeObject(toc, Newtonsoft.Json.Formatting.None,
   new JsonSerializerSettings
   {
       NullValueHandling = NullValueHandling.Ignore
   });

            json = ReplacementToc(json);
            string jsonFormatted = json;
            try
            {
                jsonFormatted = "var toc_001 = " + JValue.Parse(json).ToString(Newtonsoft.Json.Formatting.Indented);
            }
            catch (Exception ex)
            { 
            
            }
            // json=CleanNodes(json);
            File.WriteAllText(assetFolderPath + "\\Toc_001.js", jsonFormatted);
            Console.WriteLine("Toc Creation: \t\tCompleted Successfully." + Environment.NewLine + jsonFormatted);
        }
        private static void WriteTranscriptsListJson(string assetFolderPath)
        {
            TranscriptClass.Root transcriptList = new TranscriptClass.Root();
            transcriptList.id = "readTranscript_001";
            List<TranscriptClass.Component> components = new List<TranscriptClass.Component>();
            TranscriptClass.Component comp = new TranscriptClass.Component();
            if (SharedObjects.TranscriptList.Count > 0)
            {
                comp.TemplateID = "readTranscript_001";
                foreach (KeyValuePair<string, TranscriptClass.Transcript> item in SharedObjects.TranscriptList)
                {
                    if (comp.Transcripts == null)
                    {
                        List<TranscriptClass.Transcript> transcripts = new List<TranscriptClass.Transcript>();
                        transcripts.Add(item.Value);
                        comp.Transcripts = transcripts;
                    }
                    else
                    {
                        comp.Transcripts.Add(item.Value);
                    }
                }
            }
            components.Add(comp);
            transcriptList.components = components;
            var json = JsonConvert.SerializeObject(transcriptList, Newtonsoft.Json.Formatting.None,
   new JsonSerializerSettings
   {
       NullValueHandling = NullValueHandling.Ignore
   });

            json = ReplacementToc(json);
            string jsonFormatted = json;
            try
            {
                jsonFormatted = "var readTranscript_001 = " + JValue.Parse(json).ToString(Newtonsoft.Json.Formatting.Indented);
            }
            catch (Exception ex)
            {

            }
            // json=CleanNodes(json);
            File.WriteAllText(assetFolderPath + "\\readTranscript_001.js", jsonFormatted);
            Console.WriteLine("Transcript List Creation: \t\tCompleted Successfully." + Environment.NewLine + jsonFormatted);
        }
        private static Meta GetNavigationData(Meta cmp, string term, string def, HtmlNode node, Page page)
        {
            HtmlAgilityPack.HtmlDocument DOCTInner = new HtmlAgilityPack.HtmlDocument();
            DOCTInner.LoadHtml(term.Trim().ToLower());
            String termInner = DOCTInner.DocumentNode.InnerText;
            if (term.Trim().ToLower().Contains("type"))
            {
                cmp.Type = CleanUpText(def, "", false, null, null);
                HtmlAgilityPack.HtmlDocument DOCT = new HtmlAgilityPack.HtmlDocument();
                DOCT.LoadHtml(cmp.Type);
                String type = DOCT.DocumentNode.InnerText;
                SharedObjects.component.Lesson = type.Replace("\r", "").Replace("\n", "");
                if (cmp.Type == null)
                {
                    SharedObjects.component.Lesson = "Lesson";
                }
            }

            if (term.Trim().ToLower().Contains("file_name"))
                cmp.File_name = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("version"))
                cmp.Version = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("due_date"))
            {
                HtmlAgilityPack.HtmlDocument DOCT = new HtmlAgilityPack.HtmlDocument();
                DOCT.LoadHtml(CleanUpText(def, "", false, null, null));
                String date = DOCT.DocumentNode.InnerText;
                SharedObjects.component.Due_Date = date.Replace("\r", "").Replace("\n", "");
            }
            if (termInner == "instruction")
            {
                HtmlAgilityPack.HtmlDocument DOCT = new HtmlAgilityPack.HtmlDocument();
                DOCT.LoadHtml(CleanUpText(def, "", false, null, null));
                string instruction = DOCT.DocumentNode.InnerText;
                SharedObjects.component.Instruction = instruction.Replace("\r","").Replace("\n", "");
            }
            if (term.Trim().ToLower().Contains("brief_description"))
                cmp.Brief_description = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("long_description"))
                cmp.Long_description = CleanUpText(def, "", true, null, null);

            if ((term.Trim().ToLower().Contains("learning_intention")) || (term.Trim().ToLower().Contains("learning intention")))
                cmp.Learning_intention = CleanUpText(def, "", false, null, null);

            if ((term.Trim().ToLower().Contains("success_criteria")) || (term.Trim().ToLower().Contains("success criteria")))
                cmp.Success_criteria = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("subject"))
                cmp.Subject = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("year level"))
                cmp.Year_level = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("course"))
                cmp.Course = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("unit"))
                cmp.Unit = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("state"))
                cmp.State = CleanUpText(def, "", false, null, null);

            if ((term.Trim().ToLower().Contains("ac_code")) || (term.Trim().ToLower().Contains("ac code")))
                cmp.AC_code = CleanUpText(def, "", false, null, null);

            if ((term.Trim().ToLower().Contains("ac_descriptor")) || (term.Trim().ToLower().Contains("ac descriptor")))
                cmp.AC_descriptor = CleanUpText(def, "", false, null, null);


            if ((term.Trim().ToLower().Contains("estimated_time")) || (term.Trim().ToLower().Contains("estimated time")))
                cmp.Estimated_time = CleanUpText(def, "", false, null, null);

            if (term.Trim().ToLower().Contains("toc title"))
            {
                cmp.TocTitle = CleanUpText(def, "", false, null, null);
                SharedObjects.component.Title = cmp.TocTitle;
            }

            if (term.Trim().ToLower().Contains("notes"))
            {
                cmp.Notes = CleanUpText(def, "", false, null, null);
                SharedObjects.component.Notes = cmp.Notes;
            }
            if (term.Trim().ToLower().Contains("teacher only"))
            {
                cmp.Teacher_only = CleanUpText(def, "", false, null, null);
                SharedObjects.component.TeacherOnly = cmp.Teacher_only;
            }

            if (term.Trim().ToLower().Contains("toc main text"))
            {
                cmp.MainText = CleanUpText(def, "", false, null, null);
                SharedObjects.component.MainText = cmp.MainText;
            }
            if (term.Trim().ToLower().Contains("discoverable"))
            {
                cmp.Discoverable = CleanUpText(def, "", false, null, null);
                SharedObjects.component.Discoverable = cmp.Discoverable;
            }
            if (term.Trim().ToLower().Contains("toc template-description"))
            {
                cmp.Template_Description = CleanUpText(def, "", false, null, null);
                SharedObjects.component.TemplateDescription = cmp.Template_Description;
            }

            return cmp;
        }
        private static Component GetComponent(Component cmp, string term, string def, HtmlNode node, string imagesDir, Page.Root page)
        {
            if (term.Trim().ToLower().Contains("templateid"))
            {
                cmp.TemplateID = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));
                cmp.PopupID = cmp.TemplateID;
            }
            if (term.Trim().ToLower().Contains("title"))
            {
                def = CleanUpText(def, imagesDir, false, cmp, page);
                cmp.Title = RemoveFormatting(def);
            }
            if (term.Trim().ToLower().Contains("template_num"))
                cmp.PopupDataID = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("templatename"))
                cmp.TemplateName = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("template description"))
                cmp.TemplateDescription = CleanUpText(def, imagesDir, false, cmp, page);

            if (term.Trim().ToLower().Contains("graphic"))
                cmp.Graphic = GetImageName(CleanUpText(def, imagesDir, false, cmp, page), imagesDir);

            if (term.Trim().ToLower().Contains("ffn"))
            {
                string imageName= GetImageName(CleanUpText(def, imagesDir, false, cmp, page), imagesDir);
                Regex reg = new Regex(@"(.+?)>(.+?)<(.+)");
                imageName = reg.Replace(imageName, "$2");
                string extension = Path.GetExtension(imageName);
                cmp.FFN = imageName;
                if (extension.Length > 0)
                {
                    if (!imageName.StartsWith("http"))
                    {
                        if (extension.StartsWith(".htm"))
                        { cmp.FFN = "./assets/" + imageName; }
                        else
                        {
                            cmp.FFN = "./assets/images/" + imageName;
                        }
                    }
                }
                //
            }
            if (term.Trim().ToLower().Contains("caption"))
                cmp.Caption = CleanUpText(def, imagesDir, true, cmp, page);

            if (term.Trim().ToLower().Contains("acknowledgements"))
                cmp.Acknowledgements = CleanUpText(def, imagesDir, false, cmp, page);

            if (term.Trim().ToLower().Contains("alt text"))
                cmp.AltText = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("media position"))
                cmp.MediaPosition = CleanUpText(def, imagesDir, false, cmp, page);

            if (term.Trim().ToLower().Contains("transcript id"))
                cmp.Transcript_ID = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("srt/vtt"))
                cmp.SRT_VTT = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("timeline graphic"))
                cmp.Timeline_graphic = CleanUpText(def, imagesDir, false, cmp, page);

            if (term.Trim().ToLower().Contains("teacher only"))
                cmp.TeacherOnly = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("discoverable"))
                cmp.Discoverable = RemoveFormatting(CleanUpText(def, imagesDir, false, cmp, page));

            if (term.Trim().ToLower().Contains("notes"))
            {
                if (cmp.Notes == null)
                {
                    cmp.Notes = CleanUpText(def, imagesDir, true, cmp, page);
                }
            }

            if (term.Trim().ToLower().Contains("activityid"))
                cmp.ActivityID = CleanUpText(def, imagesDir, true, cmp, page);

            #region Array Type data Allocation
            if ((term.Trim().ToLower().Contains("panel")) || (term.Trim().ToLower().Contains("tab")) || (term.Trim().ToLower().Contains("hotspot")) || (term.Trim().ToLower().Contains("card")) || (term.Trim().ToLower().Contains("text")) || (term.Trim().ToLower().Contains("slide")) || (term.Trim().ToLower().Contains("main-text")) || (term.Trim().ToLower().Contains("transcript")))
            {
                if (!term.Trim().ToLower().Contains("transcript id"))
                {
                    cmp = CreateArray(term.Trim().ToLower(), node, cmp, def, imagesDir, page);
                }
            }
            #endregion
            return cmp;
        }
        private static Component CreateArray(string templateText, HtmlNode node, Component cmp, string def, string imagesDir, Page.Root page)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(templateText);
            templateText = hDoc.DocumentNode.InnerText.Replace("\r", "").Replace("\n", "").Replace("\t", "").Trim();
            templateText = templateText.Split(' ')[0].Trim();
            if (templateText == "panel")
            {
                List<PEA_Docx_to_Widget.Page.Panel> panels = ListAllPanels(node, imagesDir, page, cmp);
                cmp.Panels = panels;
            }
            if (templateText == "slide")
            {
                List<Slide> slides = ListAllSlides(node, imagesDir, page, cmp);
                cmp.Slides = slides;
            }
            if (templateText == "transcript")
            {
                if (cmp.TemplateName.ToLower().Contains("video"))
                {
                    List<PEA_Docx_to_Widget.Page.TranscriptDataVideo> slides = ListAllVideoTranscript(node, imagesDir, page, cmp);
                    cmp.video_transcript = slides;
                }
                else
                {
                    List<PEA_Docx_to_Widget.Page.TranscriptDataAudio> slides = ListAllAudioTranscript(node, imagesDir, page, cmp);
                    cmp.audio_transcript = slides;
                }
            }
            if (templateText == "main-text")
            {
                List<Main_Text> main_Texts = ListAllMainText(node, imagesDir, page, cmp);
                if (main_Texts.Count == 0)
                {
                    if (templateText.Trim().ToLower().Contains("main-text"))
                    {
                        if ((def.Contains("&lt;")) && (def.Contains("&gt;")) && (def.Contains("|")))
                        {
                            MatchCollection matches = Regex.Matches(def, @"&lt;(.+?)&gt;");
                            foreach (Match match in matches)
                            {
                                string text = match.Value;
                                if ((text.Contains("gloss:")) || (text.Contains("pop-up|")))
                                { cmp.MainText = CleanUpText(def, imagesDir, true, cmp, page); }
                                else
                                {
                                    if (text.Contains("|"))
                                    {
                                        if (text.Contains("image|"))
                                        {
                                            string alt = text.Split('|')[text.Split('|').Length - 2].Split('=')[1].Trim();
                                            string imageName = text.Split('|')[text.Split('|').Length - 1].Trim().Replace("FFN", "").Replace(":", "").Trim();
                                            //<img id='expand-image' src=""/>
                                            string imageTag = "<img id='expand-image' src='./assets/images/" + imageName + "' alt='" + alt + "'/>";
                                            def = def.Replace(match.Value, imageTag);
                                            cmp.MainText = CleanUpText(def, imagesDir, true, cmp, page);
                                        }
                                        else
                                        {
                                            if (text.Contains("media|"))
                                            {
                                                PEA_Docx_to_Widget.TranscriptClass.Transcript transcript = new PEA_Docx_to_Widget.TranscriptClass.Transcript();
                                                string transcriptText = text.Split('|')[text.Split('|').Length - 2].Split('=')[1].Trim();
                                                string assetName = text.Split('|')[text.Split('|').Length - 1].Trim().Replace("FFN", "").Replace(":", "").Trim().Replace("&gt;", "");
                                                string id = "AudioTranscriptID_" + SharedObjects.TranscriptList.Count;
                                                if (Path.GetExtension(assetName).ToLower() == ".mp3")
                                                {
                                                    string audioTag = "<audio src='./assets/audio/" + Path.GetFileName(assetName) + "' controls controlsList='nodownload' class='audio-container'></audio><span id='read-transcript' transcriptID='" + id + "'></span>";
                                                    def = def.Replace(match.Value, audioTag);
                                                    cmp.MainText = CleanUpText(def, imagesDir, true, cmp, page);
                                                }
                                                else {
                                                    id = "VideoTranscriptID_" + SharedObjects.TranscriptList.Count;
                                                    string videoTag = "<div class='iframeVideoContainer'><iframe src='" + Path.GetFileName(assetName) + "' allow='fullscreen'></iframe></div><span id='read-transcript' transcriptID='" + id + "'></span><span id='read-transcript' transcriptID='" + id + "'></span>";
                                                    def = def.Replace(match.Value, videoTag);
                                                    cmp.MainText = CleanUpText(def, imagesDir, true, cmp, page);
                                                }
                                                transcript.id = id;
                                                transcript.Transcript_Text = transcriptText;
                                                SharedObjects.TranscriptList.Add(id, transcript);
                                            }
                                            else
                                            {
                                                string DownloadButton = text.Split('|')[text.Split('|').Length - 1].Replace("&lt;", "").Replace("&gt;", "").Trim();
                                                string notes = text.Split('|')[text.Split('|').Length - 2].Replace("&lt;", "").Replace("&gt;", "").Trim();
                                                if (cmp.TemplateName.ToLower().Trim().Contains("widget and text"))
                                                {
                                                    cmp.Notes = "Null";
                                                    cmp.DownloadPDF = notes;
                                                }
                                                else { cmp.Notes = notes; }
                                                cmp.DownloadButtonText = DownloadButton;

                                                cmp.MainText = CleanUpText(def.Replace(text, "<span id=\"file-download\" downloadItem=" + "\"" + notes + "\"" + " downloadButtonText=" + "\"" + DownloadButton + "\"" + "></span>"), imagesDir, true, cmp, page);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            cmp.MainText = CleanUpText(def, imagesDir, true, cmp, page);
                        }
                    }
                }
                else
                {
                    cmp.Main_Texts = main_Texts;
                    CreateDownloadNotes(cmp, imagesDir, page);
                }
            }
            if (templateText == "text")
            {
                List<Text> texts = ListAllTexts(node, imagesDir, page, cmp);
                cmp.Texts = texts;
            }
            if (templateText == "hotspot")
            {
                List<Hotspot> hotspots = ListAllHotspots(node, imagesDir, page, cmp);
                cmp.Hotspots = hotspots;
            }
            if (templateText == "tab")
            {
                List<Tab> tabs = ListAllTabs(node, imagesDir, page, cmp);
                cmp.tabData = tabs;
            }
            if (templateText == "card")
            {
                List<Card> cards = ListAllCards(node, imagesDir, page, cmp);

                //if ((cards[cards.Count - 1].Card_Front_Video != null) && (cards[cards.Count - 1].Card_Front_Video.video == "") && (cards[cards.Count - 1].Card_Front_Video.transcript == ""))
                //{
                //    cards[cards.Count - 1].Card_Front_Video = null;
                //}
                //if ((cards[cards.Count - 1].Card_Back_Video != null) && (cards[cards.Count - 1].Card_Back_Video.video == "") && (cards[cards.Count - 1].Card_Back_Video.transcript == ""))
                //{
                //    cards[cards.Count - 1].Card_Back_Video = null;
                //}
                //if ((cards[cards.Count - 1].Card_Front_Audio != null) && (cards[cards.Count - 1].Card_Front_Audio.audio == "") && (cards[cards.Count - 1].Card_Front_Audio.transcript == ""))
                //{
                //    cards[cards.Count - 1].Card_Front_Audio = null;
                //}
                //if ((cards[cards.Count - 1].Card_Back_Audio != null) && (cards[cards.Count - 1].Card_Back_Audio.audio == "") && (cards[cards.Count - 1].Card_Back_Audio.transcript == ""))
                //{
                //    cards[cards.Count - 1].Card_Back_Audio = null;
                //}

                if (cmp.flipCardData == null)
                {
                    cmp.flipCardData = cards;
                }
                else 
                {
                    foreach (Card card in cards)
                    {
                        cmp.flipCardData.Add(card);
                    }
                }
            }
            return cmp;
        }
        private static List<PEA_Docx_to_Widget.Page.TranscriptDataVideo> ListAllVideoTranscript(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.TranscriptDataVideo> main_Texts = new List<PEA_Docx_to_Widget.Page.TranscriptDataVideo>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.TranscriptDataVideo cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataVideo();
            int temp = 1;
            bool foundDuration = false;
            for (int i = 0; i < 300; i++)
            {
                if (nextSib == null) 
                { 
                    break;
                }
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[0].InnerText;
                        string def = CleanUpText(tdNodes[1].InnerHtml, imagesDir, true, cmp, page);
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower() == "transcript")|| (term.Trim().ToLower() == "duration"))
                        {
                            if (term.Trim().ToLower().Contains("transcript"))
                            {
                                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataVideo();
                                cPanel.title = "Read Transcript";
                                cPanel.id = temp;
                                cPanel.Transcript = def;
                                temp++;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("duration"))
                            {
                                foundDuration = true;
                                HtmlAgilityPack.HtmlDocument DDoc = new HtmlAgilityPack.HtmlDocument();
                                DDoc.LoadHtml(def);
                                cPanel.videoTime = DDoc.DocumentNode.InnerText.Trim();
                                main_Texts.Add(cPanel);
                                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataVideo();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                        else { nextSib = nextSib.NextSibling; }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (foundDuration == false)
            {
                main_Texts.Add(cPanel);
                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataVideo();
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<PEA_Docx_to_Widget.Page.TranscriptDataAudio> ListAllAudioTranscript(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.TranscriptDataAudio> main_Texts = new List<PEA_Docx_to_Widget.Page.TranscriptDataAudio>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.TranscriptDataAudio cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataAudio();
            int temp = 1; 
            bool foundDuration = false;
            for (int i = 0; i < 100; i++)
            {
                if (nextSib == null) { break; }
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[0].InnerText;
                        string def = CleanUpText(tdNodes[1].InnerHtml, imagesDir, true, cmp, page);
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower() == "transcript")|| (term.Trim().ToLower() == "duration"))
                        {
                            if (term.Trim().ToLower().Contains("transcript"))
                            {
                                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataAudio();
                                cPanel.title = "Read Transcript";
                                cPanel.id = temp;
                                cPanel.Transcript_txt = def;
                                temp++;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("duration"))
                            {
                                foundDuration = true;
                                HtmlAgilityPack.HtmlDocument DDoc = new HtmlAgilityPack.HtmlDocument();
                                DDoc.LoadHtml(def);
                                cPanel.audioTime = DDoc.DocumentNode.InnerText.Trim();
                                main_Texts.Add(cPanel);
                                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataAudio();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                        else {
                            nextSib = nextSib.NextSibling;
                            //break;
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (foundDuration == false)
            {
                main_Texts.Add(cPanel);
                cPanel = new PEA_Docx_to_Widget.Page.TranscriptDataAudio();
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<Text> ListAllTexts(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.Text> main_Texts = new List<PEA_Docx_to_Widget.Page.Text>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.Text cPanel = new PEA_Docx_to_Widget.Page.Text();
            int temp = 1;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        try
                        {
                            string term = tdNodes[1].InnerHtml;
                            string def = tdNodes[2].InnerHtml;
                            string id = nextSib.Attributes["id"].Value;
                            if ((term.Trim().ToLower().Contains("text_heading")) || (term.Trim().ToLower().Contains("text_text")))
                            {
                                if (term.Trim().ToLower().Contains("text_heading"))
                                {
                                    cPanel = new PEA_Docx_to_Widget.Page.Text();
                                    def = RemoveFormatting(def);
                                    cPanel.id = temp;
                                    temp++;
                                    cPanel.Text_Heading = def;
                                    nodesToDel.Add(id, nextSib);
                                }
                                if (term.Trim().ToLower().Contains("text_text"))
                                {
                                    cPanel.Text_Text = CleanUpText(def, imagesDir, true, cmp, page);
                                    main_Texts.Add(cPanel);
                                    nodesToDel.Add(id, nextSib);
                                    cPanel = new PEA_Docx_to_Widget.Page.Text();
                                }
                                nextSib = nextSib.NextSibling;
                            }
                            else { break; }
                        }
                        catch (Exception)
                        {
                            string term = tdNodes[0].InnerHtml;
                            string def = tdNodes[1].InnerHtml;
                            string id = nextSib.Attributes["id"].Value;
                            if ((term.Trim().ToLower().Contains("text_heading")) || (term.Trim().ToLower().Contains("text_text")))
                            {
                                if (term.Trim().ToLower().Contains("text_heading"))
                                {
                                    cPanel = new PEA_Docx_to_Widget.Page.Text();
                                    def = RemoveFormatting(def);
                                    cPanel.id = temp;
                                    temp++;
                                    cPanel.Text_Heading = def;
                                    nodesToDel.Add(id, nextSib);
                                }
                                if (term.Trim().ToLower().Contains("text_text"))
                                {
                                    cPanel.Text_Text = CleanUpText(def, imagesDir, true, cmp, page);
                                    main_Texts.Add(cPanel);
                                    nodesToDel.Add(id, nextSib);
                                    cPanel = new PEA_Docx_to_Widget.Page.Text();
                                }
                                nextSib = nextSib.NextSibling;
                            }
                            else { break; }
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<Card> ListAllCards(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.Card> main_Texts = new List<PEA_Docx_to_Widget.Page.Card>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.Card cPanel = new PEA_Docx_to_Widget.Page.Card();
            PEA_Docx_to_Widget.Page.Card_Front_Audio Card_Front_Audio = new PEA_Docx_to_Widget.Page.Card_Front_Audio();
            PEA_Docx_to_Widget.Page.Card_Front_Video Card_Front_Video = new PEA_Docx_to_Widget.Page.Card_Front_Video();
            PEA_Docx_to_Widget.Page.Card_Back_Audio Card_Back_Audio = new PEA_Docx_to_Widget.Page.Card_Back_Audio();
            PEA_Docx_to_Widget.Page.Card_Back_Video Card_Back_Video = new PEA_Docx_to_Widget.Page.Card_Back_Video();
            int temp = 1;
            int currentIndex = 0;
            for (int i = 0; i < 150; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[1].InnerText;
                        string def = tdNodes[2].InnerHtml;
                        string id = nextSib.Attributes["id"].Value;

                        if ((term.Trim().ToLower().Contains("card_front_heading")) || (term.Trim().ToLower().Contains("card_front_content"))|| (term.Trim().ToLower().Contains("card_front_audio"))|| (term.Trim().ToLower().Contains("card_front_image"))||(term.Trim().ToLower().Contains("card_back_heading")) || (term.Trim().ToLower().Contains("card_back_content")) || (term.Trim().ToLower().Contains("card_back_audio"))|| (term.Trim().ToLower().Contains("card_back_image")) || (term.Trim().ToLower().Contains("card_front_video")) || (term.Trim().ToLower().Contains("card_back_video")) || (term.Trim().ToLower().Contains("card_front_video_transcript")) || (term.Trim().ToLower().Contains("card_front_audio_transcript")) || (term.Trim().ToLower().Contains("card_back_video_transcript")) || (term.Trim().ToLower().Contains("card_back_audio_transcript")))
                        {
                            if ((term.Trim().ToLower().Contains("card_front_heading")))
                            {
                                //if (main_Texts.Count > 0)
                                //{
                                //    if ((main_Texts[main_Texts.Count - 1].Card_Front_Video!=null) && (main_Texts[main_Texts.Count - 1].Card_Front_Video.video == "") && (main_Texts[main_Texts.Count - 1].Card_Front_Video.transcript == ""))
                                //    {
                                //        main_Texts[main_Texts.Count - 1].Card_Front_Video = null;
                                //    }
                                //    if ((main_Texts[main_Texts.Count - 1].Card_Back_Video != null) && (main_Texts[main_Texts.Count - 1].Card_Back_Video.video == "") && (main_Texts[main_Texts.Count - 1].Card_Back_Video.transcript == ""))
                                //    {
                                //        main_Texts[main_Texts.Count - 1].Card_Back_Video = null;
                                //    }
                                //    if ((main_Texts[main_Texts.Count - 1].Card_Front_Audio != null) && (main_Texts[main_Texts.Count - 1].Card_Front_Audio.audio == "") && (main_Texts[main_Texts.Count - 1].Card_Front_Audio.transcript == ""))
                                //    {
                                //        main_Texts[main_Texts.Count - 1].Card_Front_Audio = null;
                                //    }
                                //    if ((main_Texts[main_Texts.Count - 1].Card_Back_Audio != null) && (main_Texts[main_Texts.Count - 1].Card_Back_Audio.audio == "") && (main_Texts[main_Texts.Count - 1].Card_Back_Audio.transcript == ""))
                                //    {
                                //        main_Texts[main_Texts.Count - 1].Card_Back_Audio = null;
                                //    }
                                //}
                                string cardFront = cPanel.Card_Front_Heading.ToLower().Trim().Replace("na", "").Trim();
                                if ((cardFront == ""))
                                {
                                    main_Texts.Add(cPanel);
                                    currentIndex = main_Texts.Count - 1;
                                }
                                else
                                {
                                    cPanel = new PEA_Docx_to_Widget.Page.Card();
                                    Card_Front_Audio = new PEA_Docx_to_Widget.Page.Card_Front_Audio();
                                    Card_Front_Video = new PEA_Docx_to_Widget.Page.Card_Front_Video();
                                    Card_Back_Audio = new PEA_Docx_to_Widget.Page.Card_Back_Audio();
                                    Card_Back_Video = new PEA_Docx_to_Widget.Page.Card_Back_Video();
                                    cPanel.Card_Back_Audio = Card_Back_Audio;
                                    cPanel.Card_Front_Audio = Card_Front_Audio;
                                    cPanel.Card_Back_Video = Card_Back_Video;
                                    cPanel.Card_Front_Video = Card_Front_Video;
                                    main_Texts.Add(cPanel);
                                    currentIndex = main_Texts.Count - 1;
                                }
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Front_Heading = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("card_front_content"))
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Front_Content = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower().Contains("card_front_image"))
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Front_Image = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower()=="card_front_audio")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Front_Audio == null)
                                {
                                    Card_Front_Audio.audio = text;
                                    main_Texts[currentIndex].Card_Front_Audio = Card_Front_Audio;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Front_Audio.audio = text;
                                }
                                //main_Texts[currentIndex].Card_Front_Audio.audio = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_front_audio_transcript")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Front_Audio == null)
                                {
                                    Card_Front_Audio.transcript = text;
                                    main_Texts[currentIndex].Card_Front_Audio = Card_Front_Audio;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Front_Audio.transcript = text;
                                }
                                //main_Texts[currentIndex].Card_Front_Audio.transcript = text;
                                if(!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_front_video")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Front_Video == null)
                                {
                                    Card_Front_Video.video = text;
                                    main_Texts[currentIndex].Card_Front_Video = Card_Front_Video;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Front_Video.video = text;
                                }
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_front_video_transcript")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Front_Video == null)
                                {
                                    Card_Front_Video.transcript = text;
                                    main_Texts[currentIndex].Card_Front_Video = Card_Front_Video;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Front_Video.transcript = text;
                                }
                                //main_Texts[currentIndex].Card_Front_Video.transcript = text;
                                if(!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if ((term.Trim().ToLower().Contains("card_back_heading")))
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Back_Heading = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower().Contains("card_back_content"))
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Back_Content = text;
                                nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower().Contains("card_back_image"))
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                main_Texts[currentIndex].Card_Back_Image = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_back_audio")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Back_Audio == null)
                                {
                                    Card_Back_Audio.audio = text;
                                    main_Texts[currentIndex].Card_Back_Audio = Card_Back_Audio;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Back_Audio.audio = text;
                                }
                                //main_Texts[currentIndex].Card_Back_Audio.audio = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_back_audio_transcript")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Back_Audio == null)
                                {
                                    Card_Back_Audio.transcript = text;
                                    main_Texts[currentIndex].Card_Back_Audio = Card_Back_Audio;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Back_Audio.transcript = text;
                                }
                                // main_Texts[currentIndex].Card_Back_Audio.transcript = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_back_video")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Back_Video == null)
                                {
                                    Card_Back_Video.video = text;
                                    main_Texts[currentIndex].Card_Back_Video = Card_Back_Video;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Back_Video.video = text;
                                }
                                //main_Texts[currentIndex].Card_Back_Video.video = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            if (term.Trim().ToLower() == "card_back_video_transcript")
                            {
                                string text = CleanUpText(def, imagesDir, true, cmp, page).Replace("</p>", "").Replace("<p>", "").Trim();
                                if (text.ToLower() == "na")
                                {
                                    text = "";
                                }
                                if (main_Texts[currentIndex].Card_Back_Video == null)
                                {
                                    Card_Back_Video.transcript = text;
                                    main_Texts[currentIndex].Card_Back_Video = Card_Back_Video;
                                }
                                else
                                {
                                    main_Texts[currentIndex].Card_Back_Video.transcript = text;
                                }
                                //main_Texts[currentIndex].Card_Back_Video.transcript = text;
                                if (!nodesToDel.ContainsKey(id))
                                    nodesToDel.Add(id, nextSib);
                                //cPanel = new PEA_Docx_to_Widget.Page.Card();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                        else { break; }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            TranscriptClass.Transcript transcript = new TranscriptClass.Transcript();
            foreach (Card card in main_Texts)
            {
                if (card.Card_Front_Audio != null)
                {
                    if (card.Card_Front_Audio.audio != "")
                    {
                        transcript = new TranscriptClass.Transcript();
                        transcript.id = "AudioTranscriptID_" + SharedObjects.TranscriptList.Count;
                        string transcriptText = card.Card_Front_Audio.transcript;
                        transcript.Transcript_Text = transcriptText;
                        if (transcriptText.Trim().Length > 0)
                        {
                            SharedObjects.TranscriptList.Add(transcript.id, transcript);
                            card.Card_Front_Audio.transcript = "<span id='read-transcript' transcriptID='" + transcript.id + "' videoTime='4 min'></span>";
                        }
                    }
                }
                if (card.Card_Back_Audio != null)
                {
                    if (card.Card_Back_Audio.audio != "")
                    {
                        transcript = new TranscriptClass.Transcript();
                        transcript.id = "AudioTranscriptID_" + SharedObjects.TranscriptList.Count;
                        string transcriptText = card.Card_Back_Audio.transcript;
                        transcript.Transcript_Text = transcriptText;
                        if (transcriptText.Trim().Length > 0)
                        {
                            SharedObjects.TranscriptList.Add(transcript.id, transcript);
                            card.Card_Back_Audio.transcript = "<span id='read-transcript' transcriptID='" + transcript.id + "' videoTime='4 min'></span>";
                        }
                    }
                }
                if (card.Card_Front_Video != null)
                {
                    if (card.Card_Front_Video.video != "")
                    {
                        transcript = new TranscriptClass.Transcript();
                        transcript.id = "VideoTranscriptID_" + SharedObjects.TranscriptList.Count;
                        string transcriptText = card.Card_Front_Video.transcript;
                        transcript.Transcript_Text = transcriptText;
                        if (transcriptText.Trim().Length > 0)
                        {
                            SharedObjects.TranscriptList.Add(transcript.id, transcript);
                            card.Card_Front_Video.transcript = "<span id='read-transcript' transcriptID='" + transcript.id + "' videoTime='4 min'></span>";
                        }
                    }
                }
                if (card.Card_Back_Video != null)
                {
                    if (card.Card_Back_Video.video != "")
                    {
                        transcript = new TranscriptClass.Transcript();
                        transcript.id = "VideoTranscriptID_" + SharedObjects.TranscriptList.Count;
                        string transcriptText = card.Card_Back_Video.transcript;
                        transcript.Transcript_Text = transcriptText;
                        if (transcriptText.Trim().Length > 0)
                        {
                            SharedObjects.TranscriptList.Add(transcript.id, transcript);
                            card.Card_Back_Video.transcript = "<span id='read-transcript' transcriptID='" + transcript.id + "' videoTime='4 min'></span>";
                        }
                    }
                }
            }
            foreach (Card card in main_Texts)
            {
                Card_Front_Audio = new PEA_Docx_to_Widget.Page.Card_Front_Audio();
                Card_Front_Video = new PEA_Docx_to_Widget.Page.Card_Front_Video();
                Card_Back_Audio = new PEA_Docx_to_Widget.Page.Card_Back_Audio();
                Card_Back_Video = new PEA_Docx_to_Widget.Page.Card_Back_Video();
                if (card.Card_Front_Audio == null)
                {
                    Card_Front_Audio.audio = "";
                    Card_Front_Audio.transcript = "";
                    card.Card_Front_Audio = Card_Front_Audio;
                }
                if (card.Card_Back_Audio == null)
                {
                    Card_Back_Audio.audio = "";
                    Card_Back_Audio.transcript = "";
                    card.Card_Back_Audio = Card_Back_Audio;
                }
                if (card.Card_Front_Video == null)
                {
                    Card_Front_Video.video = "";
                    Card_Front_Video.transcript = "";
                    card.Card_Front_Video = Card_Front_Video;
                }
                if (card.Card_Back_Video == null)
                {
                    Card_Back_Video.video = "";
                    Card_Back_Video.transcript = "";
                    card.Card_Back_Video = Card_Back_Video;
                }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<Hotspot> ListAllHotspots(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<Hotspot> main_Texts = new List<Hotspot>();
            HtmlNode nextSib = node;
            Hotspot cPanel = new Hotspot();
            int temp = 1;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[1].InnerHtml;
                        string def = tdNodes[2].InnerHtml;
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower().Contains("hotspot_title")) || (term.Trim().ToLower().Contains("hotspot_text")) || (term.Trim().ToLower().Contains("hotspot_position")) || (term.Trim().ToLower().Contains("hotspot_reveal-title")) || (term.Trim().ToLower().Contains("hotspot_reveal-text")))
                        {
                            if (term.Trim().ToLower().Contains("hotspot_title"))
                            {
                                cPanel = new Hotspot();
                                def = RemoveFormatting(def);
                                cPanel.id = temp;
                                temp++;
                                cPanel.Hotspot_Title = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("hotspot_text"))
                            {
                                cPanel.Hotspot_Text = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("hotspot_position"))
                            {
                                cPanel.Hotspot_Position = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("hotspot_reveal-title"))
                            {
                                cPanel.Hotspot_Reveal_title = def;
                                def = RemoveFormatting(def);
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("hotspot_reveal-text"))
                            {
                                cPanel.Hotspot_Reveal_text = CleanUpText(def, imagesDir, true, cmp, page);
                                main_Texts.Add(cPanel);
                                nodesToDel.Add(id, nextSib);
                                cPanel = new Hotspot();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<Tab> ListAllTabs(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.Tab> main_Texts = new List<PEA_Docx_to_Widget.Page.Tab>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.Tab cPanel = new PEA_Docx_to_Widget.Page.Tab();
            int temp = 1;
            int videoId = 1;
            int currIndex = 0;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[1].InnerHtml;
                        string def = tdNodes[2].InnerHtml;
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower().Contains("tab_heading")) || (term.Trim().ToLower().Contains("tab_reveal-text")) || (term.Trim().ToLower().Contains("tab_reveal-learnosity")) || (term.Trim().ToLower().Contains("tab_reveal-video")) || (term.Trim().ToLower().Contains("tab_reveal-caption")) || (term.Trim().ToLower().Contains("tab_reveal-html_widget")) || (term.Trim().ToLower().Contains("tab_reveal-transcript")))
                        {
                            if (term.Trim().ToLower().Contains("tab_heading"))
                            {
                                if (cPanel.Tab_Heading != null)
                                {
                                    cPanel = new PEA_Docx_to_Widget.Page.Tab();
                                    currIndex = main_Texts.Count;
                                }
                                def = RemoveFormatting(def);
                                temp++;
                                cPanel.Tab_Heading = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("tab_reveal-text"))
                            {
                                cPanel.Tab_Reveal_text = CleanUpText(def, imagesDir, true, cmp, page);
                                main_Texts.Add(cPanel);
                                nodesToDel.Add(id, nextSib);
                            }
                            if ((term.Trim().ToLower().Contains("tab_reveal-learnosity")) || (term.Trim().ToLower().Contains("tab_reveal-html_widget")) || (term.Trim().ToLower().Contains("tab_reveal-video")))
                            {
                                String content = getUrl(CleanUpText(def, imagesDir, true, cmp, page));
                                string RText = cPanel.Tab_Reveal_text + content;
                                string linkText = RText;
                                if (main_Texts.Count > 0)
                                {
                                    if ((main_Texts.Count - 1) == currIndex)
                                    {
                                        main_Texts[main_Texts.Count - 1].Tab_Reveal_text = linkText;
                                    }
                                    else
                                    {
                                        cPanel.Tab_Reveal_text = linkText;
                                        main_Texts.Add(cPanel);
                                    }
                                }
                                else
                                {
                                    cPanel.Tab_Reveal_text = linkText; main_Texts.Add(cPanel);
                                }
                                nodesToDel.Add(id, nextSib);
                            }
                            //if (term.Trim().ToLower().Contains("tab_reveal-video"))
                            //{
                            //    cPanel.Tab_Reveal_Video = CleanUpText(def, imagesDir, true);
                            //    main_Texts.Add(cPanel);
                            //    nodesToDel.Add(id, nextSib);
                            //}
                            if (term.Trim().ToLower().Contains("tab_reveal-caption"))
                            {
                                //cPanel.Tab_Reveal_Caption = CleanUpText(def, imagesDir, true);
                                //main_Texts.Add(cPanel);
                                //nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("tab_reveal-transcript"))
                            {
                                List<TabRevealVideo> tabReveal = new List<TabRevealVideo>();
                                TabRevealVideo tabRevealVideo = new TabRevealVideo();
                                tabRevealVideo.id = videoId;
                                tabRevealVideo.Transcript = CleanUpText(def, imagesDir, true, cmp, page);
                                tabReveal.Add(tabRevealVideo);
                                cPanel.Tab_Reveal_Video = tabReveal;
                            }
                            //if (term.Trim().ToLower().Contains("tab_reveal-html_widget"))
                            //{
                            //    cPanel.Tab_Reveal_HTML_Widget = CleanUpText(def, imagesDir, true);
                            //    main_Texts.Add(cPanel);
                            //    nodesToDel.Add(id, nextSib);
                            //}
                            nextSib = nextSib.NextSibling;
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static string getUrl(string content)
        {
            var rawString = content.Replace("<p>", "").Replace("</p>", "");
            var links = rawString.Split("\t\n ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(s => s.StartsWith("http://") || s.StartsWith("www.") || s.StartsWith("https://"));
            foreach (string s in links)
            {
                rawString = rawString.Replace(s, "<iframe width='1020' height='366' src='" + s + "'></iframe>");
            }
            return rawString;
        }
        private static List<Slide> ListAllSlides(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<Slide> main_Texts = new List<Slide>();
            HtmlNode nextSib = node;
            Slide cPanel = new Slide();
            int temp = 1;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[1].InnerHtml;
                        string def = tdNodes[2].InnerHtml;
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower().Contains("slide_title")) || (term.Trim().ToLower().Contains("slide_text")) || (term.Trim().ToLower().Contains("slide_graphic")) || (term.Trim().ToLower().Contains("slide_ffn")) || (term.Trim().ToLower().Contains("slide_caption")) || (term.Trim().ToLower().Contains("slide_acknowledgements")) || (term.Trim().ToLower().Contains("slide_alt-text")) || (term.Trim().ToLower().Contains("slide_textalign")))
                        {
                            if (term.Trim().ToLower().Contains("slide_title"))
                            {
                                cPanel = new Slide();
                                def = RemoveFormatting(def);
                                cPanel.id = temp;
                                temp++;
                                cPanel.imageName = "slider" + temp + ".jpg";
                                cPanel.Slide_Title = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("slide_text"))
                            {
                                if (!term.Trim().ToLower().Contains("slide_textalign"))
                                {
                                    cPanel.Slide_Text = CleanUpText(def, imagesDir, true, cmp, page);
                                    nodesToDel.Add(id, nextSib);
                                }
                            }
                            if (term.Trim().ToLower().Contains("slide_graphic"))
                            {
                                cPanel.Slide_Graphic = Path.GetFileNameWithoutExtension(GetImageName(def, imagesDir));
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("slide_ffn"))
                            {
                                cPanel.Slide_FFN = GetImageName(def, imagesDir);
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("slide_caption"))
                            {
                                cPanel.Slide_Caption = CleanUpText(def, imagesDir, true, cmp, page);
                                nodesToDel.Add(id, nextSib);
                            }
                            
                            if (term.Trim().ToLower().Contains("slide_acknowledgements"))
                            {
                                cPanel.Slide_Acknowledgements = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("slide_alt-text"))
                            {
                                cPanel.Slide_Alt_text = RemoveFormatting(def);
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("slide_textalign"))
                            {
                                cPanel.TextAlignment = RemoveFormatting(def);
                                main_Texts.Add(cPanel);
                                if (!nodesToDel.ContainsKey(id))
                                {
                                    nodesToDel.Add(id, nextSib);
                                }
                                cPanel = new Slide();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<PEA_Docx_to_Widget.Page.Panel> ListAllPanels(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.Panel> main_Texts = new List<PEA_Docx_to_Widget.Page.Panel>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.Panel cPanel = new PEA_Docx_to_Widget.Page.Panel();
            int temp = 1;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        string term = tdNodes[1].InnerHtml;
                        string def = tdNodes[2].InnerHtml;
                        string id = nextSib.Attributes["id"].Value;
                        if ((term.Trim().ToLower().Contains("panel_heading")) || (term.Trim().ToLower().Contains("panel_reveal-text")))
                        {
                            if (term.Trim().ToLower().Contains("panel_heading"))
                            {
                                cPanel = new PEA_Docx_to_Widget.Page.Panel();
                                def = RemoveFormatting(def);
                                cPanel.id = temp;
                                temp++;
                                cPanel.Panel_Heading = def;
                                nodesToDel.Add(id, nextSib);
                            }
                            if (term.Trim().ToLower().Contains("panel_reveal-text"))
                            {
                                cPanel.PanelRevealText = CleanUpText(def, imagesDir, true, cmp, page);
                                main_Texts.Add(cPanel);
                                nodesToDel.Add(id, nextSib);
                                cPanel = new PEA_Docx_to_Widget.Page.Panel();
                            }
                            nextSib = nextSib.NextSibling;
                        }
                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static List<Main_Text> ListAllMainText(HtmlNode node, string imagesDir, Page.Root page, Component cmp)
        {
            Dictionary<string, HtmlNode> nodesToDel = new Dictionary<string, HtmlNode>();
            List<PEA_Docx_to_Widget.Page.Main_Text> main_Texts = new List<PEA_Docx_to_Widget.Page.Main_Text>();
            HtmlNode nextSib = node;
            PEA_Docx_to_Widget.Page.Main_Text cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
            int temp = 1;
            for (int i = 0; i < 100; i++)
            {
                if ((nextSib.Name != "#text") && (nextSib != null))
                {
                    string trText = nextSib.OuterHtml;
                    HtmlAgilityPack.HtmlDocument tDoc = new HtmlAgilityPack.HtmlDocument();
                    tDoc.LoadHtml(trText);
                    HtmlAgilityPack.HtmlNodeCollection tdNodes = tDoc.DocumentNode.SelectNodes("//td|//th");
                    if (tdNodes != null)
                    {
                        try
                        {
                            string term = tdNodes[1].InnerHtml;
                            string def = tdNodes[2].InnerHtml;
                            try
                            {
                                string id = null;
                                if (nextSib.Attributes.Count > 0)
                                {
                                    id = nextSib.Attributes["id"].Value;
                                }
                                if ((term.Trim().ToLower().Contains("text_heading")) || (term.Trim().ToLower().Contains("text_text")) || (term.Trim().ToLower().Contains("glossary_term")) || (term.Trim().ToLower().Contains("glossary_definition")) || (term.Trim().ToLower().Contains("glossary_image")) || (term.Trim().ToLower().Contains("glossary_media")))
                                {
                                    if (term.Trim().ToLower().Contains("text_heading"))
                                    {
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                        def = RemoveFormatting(def);
                                        cPanel.id = temp;
                                        temp++;
                                        cPanel.Text_Heading = def;
                                        nodesToDel.Add(id, nextSib);
                                    }
                                    if (term.Trim().ToLower().Contains("glossary_term"))
                                    {
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                        def = RemoveFormatting(def);
                                        cPanel.id = temp;
                                        temp++;
                                        cPanel.Glossary_Term = def;
                                        nodesToDel.Add(id, nextSib);
                                    }
                                    if (term.Trim().ToLower().Contains("text_text"))
                                    {
                                        cPanel.Text_Text = CleanUpText(def, imagesDir, true, cmp, page);
                                        main_Texts.Add(cPanel);
                                        nodesToDel.Add(id, nextSib);
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                    }
                                    if (term.Trim().ToLower().Contains("glossary_definition"))
                                    {
                                        cPanel.Glossary_Definition = CleanUpText(def, imagesDir, true, cmp, page);
                                        if (SharedObjects.glossaryList.ContainsKey(cPanel.Glossary_Term))
                                        {
                                            if (SharedObjects.glossaryList[cPanel.Glossary_Term].glossaryImage != null)
                                            {
                                                cPanel.Glossary_Definition = cPanel.Glossary_Definition + "<img id='expand-image' src='./assets/images/" + Path.GetFileName(SharedObjects.glossaryList[cPanel.Glossary_Term].glossaryImage) + "'/>";
                                            }
                                            if (SharedObjects.glossaryList[cPanel.Glossary_Term].glossaryIframe != null)
                                            {
                                                cPanel.Glossary_Definition = cPanel.Glossary_Definition + "<iframe id='expand-media' src=" + "'" + SharedObjects.glossaryList[cPanel.Glossary_Term].glossaryIframe + "'" + "/>";
                                            }
                                        }

                                        main_Texts.Add(cPanel);
                                        nodesToDel.Add(id, nextSib);
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                    }
                                    nextSib = nextSib.NextSibling;
                                }
                                else { break; }
                            }
                            catch (Exception)
                            { }
                        }
                        catch (Exception)
                        {
                            string term = tdNodes[0].InnerHtml;
                            string def = tdNodes[1].InnerHtml;
                            try
                            {
                                string id = null;
                                if (nextSib.Attributes.Count > 0)
                                {
                                    id = nextSib.Attributes["id"].Value;
                                }
                                if ((term.Trim().ToLower().Contains("text_heading")) || (term.Trim().ToLower().Contains("text_text")))
                                {
                                    if (term.Trim().ToLower().Contains("text_heading"))
                                    {
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                        def = RemoveFormatting(def);
                                        cPanel.id = temp;
                                        temp++;
                                        cPanel.Text_Heading = def;
                                        nodesToDel.Add(id, nextSib);
                                    }
                                    if (term.Trim().ToLower().Contains("text_text"))
                                    {
                                        cPanel.Text_Text = CleanUpText(def, imagesDir, true, cmp, page);
                                        main_Texts.Add(cPanel);
                                        nodesToDel.Add(id, nextSib);
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                    }
                                    if (term.Trim().ToLower().Contains("glossary_definition"))
                                    {
                                        cPanel.Glossary_Definition = CleanUpText(def, imagesDir, true, cmp, page);
                                        main_Texts.Add(cPanel);
                                        nodesToDel.Add(id, nextSib);
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                    }
                                    if (term.Trim().ToLower().Contains("glossary_term"))
                                    {
                                        cPanel = new PEA_Docx_to_Widget.Page.Main_Text();
                                        def = RemoveFormatting(def);
                                        cPanel.id = temp;
                                        temp++;
                                        cPanel.Glossary_Term = def;
                                        nodesToDel.Add(id, nextSib);
                                    }
                                    nextSib = nextSib.NextSibling;
                                }
                                else { break; }
                            }
                            catch (Exception)
                            { }
                        }

                    }
                }
                else { nextSib = nextSib.NextSibling; }
            }
            if (SharedObjects.idNodes == null)
            {
                SharedObjects.idNodes = nodesToDel;
            }
            else
            {
                foreach (KeyValuePair<string, HtmlNode> item in nodesToDel)
                {
                    SharedObjects.idNodes.Add(item.Key, item.Value);
                }
            }
            return main_Texts;
        }
        private static Component CreateDownloadNotes(Component cmp, string imagesDir, Page.Root page)
        {
            int i = 0;
            foreach (Main_Text textItem in cmp.Main_Texts)
            {
                if (textItem.Text_Text != null)
                {
                    string def = textItem.Text_Text;
                    if ((def.Contains("&lt;")) && (def.Contains("&gt;")) && (def.Contains("|")))
                    {
                        MatchCollection matches = Regex.Matches(def, @"&lt;(.+?)&gt;");
                        foreach (Match match in matches)
                        {
                            string text = match.Value;
                            if ((text.Contains("gloss:")) || (text.Contains("pop-up|")))
                            {
                                cmp.Main_Texts[i].Text_Text = CleanUpText(cmp.Main_Texts[i].Text_Text, imagesDir, true, cmp, page);
                            }
                            else
                            {
                                string DownloadButton = text.Split('|')[text.Split('|').Length - 1].Replace("&lt;", "").Replace("&gt;", "").Trim();
                                string notes = text.Split('|')[text.Split('|').Length - 2].Replace("&lt;", "").Replace("&gt;", "").Trim();
                                if (cmp.TemplateName.ToLower().Trim().Contains("widget and text"))
                                {
                                    cmp.Notes = "Null";
                                    cmp.DownloadPDF = notes;
                                }
                                else { cmp.Notes = notes; }
                                cmp.DownloadButtonText = DownloadButton;

                                cmp.Main_Texts[i].Text_Text = CleanUpText(cmp.Main_Texts[i].Text_Text.Replace(text, "<span id=\"file-download\" downloadItem="+"\""+ notes+"\""+ " downloadButtonText="+"\""+ DownloadButton+"\""+"></span>"), imagesDir, true, cmp, page);
                            }
                        }
                    }
                    else
                    {
                        cmp.Main_Texts[i].Text_Text = CleanUpText(def, imagesDir, true, cmp, page);
                    }
                }
                i++;
            }
            return cmp;
        }
        private static string ExecuteCommandMain(string command)
        {
            string exceptionString = null;
            try
            {
                var procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command) { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                // The following commands are needed to redirect the standard output.
                Process proc = new System.Diagnostics.Process { StartInfo = procStartInfo };
                System.Windows.Forms.Application.DoEvents();
                proc.Start();
                // Get the output into a string
                var result = proc.StandardOutput.ReadToEnd();
                int status = proc.ExitCode;
                // Display the command output.
                Console.WriteLine(result);
            }
            catch (Exception e)
            {
                exceptionString = "Process can't be executed.. ";
            }
            return exceptionString;
        }
        private static Dictionary<string, List<HtmlNode>> ImplementId(string html)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.Load(html, Encoding.UTF8);
            Dictionary<string, List<HtmlNode>> screensList = new Dictionary<string, List<HtmlNode>>();
            HtmlAgilityPack.HtmlNodeCollection tableNodes = hDoc.DocumentNode.SelectNodes("//table");
            if (tableNodes != null)
            {
                string livescreen = null;
                foreach (HtmlNode node in tableNodes)
                {
                    if (node.InnerText.Contains("templateId"))
                    {
                        HtmlNode screenNode = CheckScreenNode(node);
                        if (screenNode != null)
                        {
                            livescreen = screenNode.InnerText.Trim();
                            if (!screensList.ContainsKey(livescreen))
                            {
                                screensList.Add(livescreen, new List<HtmlNode>());
                                screensList[livescreen].Add(node);
                            }
                            else
                            {
                                screensList[livescreen].Add(node);
                            }
                        }
                    }
                    else
                    {
                        HtmlNode preSib = node.PreviousSibling;
                        if (preSib != null)
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                if (preSib != null)
                                {
                                    if (preSib.InnerText.Length < 15)
                                    {
                                        if ((preSib.Name != "#text") && (preSib.InnerText.ToLower().Contains("footer")))
                                        {
                                            if (!screensList.ContainsKey("footnote"))
                                            {
                                                screensList.Add("footnote", new List<HtmlNode>());
                                                screensList["footnote"].Add(node);
                                            }
                                            else
                                            {
                                                screensList["footnote"].Add(node);
                                            }
                                            break;
                                        }
                                        else { preSib = preSib.PreviousSibling; }
                                    }
                                }
                                else { break; }
                            }
                        }
                    }
                }
            }
            return screensList;
        }
        private static Dictionary<string, GlossaryItem> GetGlossaryList(string html)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.Load(html, Encoding.UTF8);
            Dictionary<string, GlossaryItem> screensList = new Dictionary<string, GlossaryItem>();
            HtmlAgilityPack.HtmlNodeCollection tableNodes = hDoc.DocumentNode.SelectNodes("//table");
            if (tableNodes != null)
            {
                GlossaryItem glterm = new GlossaryItem();
                foreach (HtmlNode node in tableNodes)
                {
                    if (node.InnerText.Contains("Glossary List"))
                    {
                        HtmlNode screenNode = CheckScreenNode(node);
                        if (screenNode != null)
                        {
                            HtmlNode SibNode = screenNode.NextSibling;
                            for (int i = 0; i < 10; i++)
                            {
                                if (SibNode.Name != "#text")
                                {
                                    if (SibNode.Name == "table")
                                    {
                                        string tableNodeTag = SibNode.OuterHtml;
                                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                        doc.Load(html, Encoding.UTF8);
                                        HtmlAgilityPack.HtmlNodeCollection trNodes = doc.DocumentNode.SelectNodes("//tr");
                                        foreach (HtmlNode htmlNode in trNodes)
                                        {
                                            string htmlText = htmlNode.InnerText;
                                            string term = null;
                                            string def = null;
                                            if (htmlText.ToLower().Contains("glossary_term"))
                                            {

                                                foreach (HtmlNode chNode in htmlNode.ChildNodes)
                                                {

                                                    if ((chNode.Name != "#text") && (chNode.InnerText.ToLower().Trim() == "glossary_term"))
                                                    {
                                                        HtmlNode htmlNode1 = chNode.NextSibling;
                                                        for (int j = 0; j < 2; j++)
                                                        {

                                                            if ((htmlNode1.Name != "#text") && (htmlNode1.InnerText.ToLower().Trim() != "glossary_term"))
                                                            {
                                                                if (glterm.Glossary_Term != null)
                                                                {
                                                                    if (!screensList.ContainsKey(glterm.Glossary_Term))
                                                                        screensList.Add(glterm.Glossary_Term, glterm);
                                                                }
                                                                glterm = new GlossaryItem();
                                                                term = htmlNode1.InnerText;
                                                                def = htmlNode1.InnerText;
                                                                glterm.Glossary_Term = def;
                                                            }
                                                            else
                                                            {
                                                                htmlNode1 = htmlNode1.NextSibling;
                                                            }
                                                        }
                                                    }
                                                }
                                                //screensList.Add()
                                            }
                                            if (htmlText.ToLower().Contains("glossary_definition"))
                                            {
                                                foreach (HtmlNode chNode in htmlNode.ChildNodes)
                                                {
                                                    if ((chNode.Name != "#text") && (chNode.InnerText.ToLower().Trim() == "glossary_definition"))
                                                    {
                                                        HtmlNode htmlNode1 = chNode.NextSibling;
                                                        for (int j = 0; j < 2; j++)
                                                        {

                                                            if ((htmlNode1.Name != "#text") && (htmlNode1.InnerText.ToLower().Trim() != "glossary_definition"))
                                                            {
                                                                term = htmlNode1.InnerHtml;
                                                                glterm.Glossary_definition = term;
                                                                //if (screensList.Count > 0)
                                                                //{
                                                                //    //var item=screensList.ElementAt(screensList.Count - 1);
                                                                //    string key = screensList.ElementAt(screensList.Count - 1).Key;

                                                                //}
                                                            }
                                                            else
                                                            {
                                                                htmlNode1 = htmlNode1.NextSibling;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (htmlText.ToLower().Contains("glossary_image"))
                                            {
                                                foreach (HtmlNode chNode in htmlNode.ChildNodes)
                                                {
                                                    if ((chNode.Name != "#text") && (chNode.InnerText.ToLower().Trim() == "glossary_image"))
                                                    {
                                                        HtmlNode htmlNode1 = chNode.NextSibling;
                                                        for (int j = 0; j < 2; j++)
                                                        {

                                                            if ((htmlNode1.Name != "#text") && (htmlNode1.InnerText.ToLower().Trim() != "glossary_image"))
                                                            {
                                                                term = htmlNode1.InnerHtml;
                                                                glterm.glossaryImage = term;
                                                                //if (screensList.Count > 0)
                                                                //{
                                                                //    //var item=screensList.ElementAt(screensList.Count - 1);
                                                                //    string key = screensList.ElementAt(screensList.Count - 1).Key;
                                                                //    screensList[key] = term;
                                                                //}
                                                            }
                                                            else
                                                            {
                                                                htmlNode1 = htmlNode1.NextSibling;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (htmlText.ToLower().Contains("glossary_media"))
                                            {
                                                foreach (HtmlNode chNode in htmlNode.ChildNodes)
                                                {
                                                    if ((chNode.Name != "#text") && (chNode.InnerText.ToLower().Trim() == "glossary_media"))
                                                    {
                                                        HtmlNode htmlNode1 = chNode.NextSibling;
                                                        for (int j = 0; j < 2; j++)
                                                        {

                                                            if ((htmlNode1.Name != "#text") && (htmlNode1.InnerText.ToLower().Trim() != "glossary_media"))
                                                            {
                                                                term = htmlNode1.InnerHtml;
                                                                glterm.glossaryIframe = term;
                                                                //if (screensList.Count > 0)
                                                                //{
                                                                //    //var item=screensList.ElementAt(screensList.Count - 1);
                                                                //    string key = screensList.ElementAt(screensList.Count - 1).Key;
                                                                //    screensList[key] = term;
                                                                //}
                                                            }
                                                            else
                                                            {
                                                                htmlNode1 = htmlNode1.NextSibling;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else { SibNode = SibNode.NextSibling; }
                            }
                        }
                    }
                }
                if (glterm.Glossary_Term != null)
                {
                    if (!screensList.ContainsKey(glterm.Glossary_Term))
                        screensList.Add(glterm.Glossary_Term, glterm);
                }
            }
            return screensList;
        }
        private static Dictionary<string, string> GetPopupList(string html)
        {
            string htmlText = File.ReadAllText(html);
            Dictionary<string, string> popupList = new Dictionary<string, string>();
            MatchCollection matches = Regex.Matches(htmlText, @"&lt;pop-up(.+?)&gt;");
            foreach (Match match in matches)
            {
                string text = match.Value.Replace("&lt;", "").Replace("&gt;", "");
                string popupUrl_data = text.Split('|')[1].Trim().Replace("url=", "").Replace("url =", "").Trim();
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(popupUrl_data);
                popupUrl_data = doc.DocumentNode.InnerText;
                bool urltype = false;
                if ((popupUrl_data.Contains("http")) || (popupUrl_data.Trim().StartsWith("www")))
                {
                    urltype = true;
                }

                string popupBtn = text.Split('|')[text.Split('|').Length - 1].Trim();
                if (!popupList.ContainsKey(text))
                    popupList.Add(text, popupUrl_data);
            }
            return popupList;
        }
        private static HtmlNode GetNavNode(string html)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.Load(html, Encoding.UTF8);
            string xpath = ConfigurationManager.AppSettings.Get("StoryBoardXPath");
            HtmlAgilityPack.HtmlNodeCollection headNodes = hDoc.DocumentNode.SelectNodes(xpath);
            HtmlNode nodeMain = null;
            if (headNodes != null)
            {
                foreach (HtmlNode node in headNodes)
                {
                    HtmlNode nextSibling = node.NextSibling;
                    for (int i = 0; i < 10; i++)
                    {
                        if (nextSibling.Name == "table")
                        {
                            nodeMain = nextSibling;
                            break;
                        }
                        else
                        {
                            nextSibling = nextSibling.NextSibling;
                        }
                    }
                }
            }
            return nodeMain;
        }
        private static HtmlAgilityPack.HtmlNodeCollection ImplementIdFunc(HtmlAgilityPack.HtmlDocument hDoc, HtmlAgilityPack.HtmlNodeCollection trNodes)
        {
            if (trNodes != null)
            {
                foreach (HtmlNode node in trNodes)
                {
                    string id = Guid.NewGuid().ToString();
                    node.Attributes.Remove("id");
                    HtmlAttribute idAttr = hDoc.CreateAttribute("id");
                    idAttr.Value = id;
                    node.Attributes.Append(idAttr);
                }
            }
            return trNodes;
        }
        private static HtmlNode CheckScreenNode(HtmlNode node)
        {
            string nodeouter = node.OuterHtml;
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(nodeouter);
            HtmlAgilityPack.HtmlNodeCollection thNodes = hDoc.DocumentNode.SelectNodes("//th");
            HtmlNode screenNode = null;
            int j = 0;
            foreach (HtmlNode htmlNode in thNodes)
            {
                if ((htmlNode.Name!="#text")&&(htmlNode.InnerText.Contains("templateId")))
                {
                    HtmlNode valueNode = thNodes[j + 1];
                    screenNode = valueNode;
                    break;
                    j++;
                }
            }
            return screenNode;
        }
        private static string GetPageName(int page)
        {
            string pageName = null;
            if (page < 10) { pageName = "page_00" + page; }
            if ((page > 9) && (page < 100)) { pageName = "page_0" + page; }
            if ((page > 99) && (page < 1000)) { pageName = "page_" + page; }
            if (page > 999) { pageName = "page_" + page; }
            return pageName;
        }
        private static string CleanUpText(string def, string imagesDir, bool addPara, Component cmp, Page.Root page)
        {
            if ((def.Contains("insert|")) || (def.Contains("gloss:")) || (def.Contains("pop-up|") || (def.Contains("image|"))))
            {
                MatchCollection matches = Regex.Matches(def, @"&lt;(.+?)&gt;");
                foreach (Match match in matches)
                {
                    string text = match.Value.Replace("&lt;", "").Replace("&gt;", "");

                    if ((text.Contains(":")) && (text.Contains("gloss:")))
                    {
                        string termDef = text.Split('|')[0].Split(':')[1].Trim();

                        HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                        htmlDocument.OptionWriteEmptyNodes = true;
                        htmlDocument.LoadHtml(termDef);

                        termDef = htmlDocument.DocumentNode.InnerText;

                        string termText = text.Split('|')[1].Trim();
                        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.OptionWriteEmptyNodes = true;
                        htmlDoc.LoadHtml(termText);
                        string termTxt = htmlDoc.DocumentNode.InnerText;
                        string glossDef = "<a data-glossary='' glossaryTerm='" + termDef + "' class='global-glossary'>" + termTxt + "</a>";
                        if (SharedObjects.glossaryList.ContainsKey(termDef))
                        {
                            glossDef = "<a data-glossary='" + SharedObjects.glossaryList[termDef].Glossary_definition.Replace("<", "&lt;").Replace(">", "&gt;").Replace("="+"\"", "=[PIPESTART]").Replace("\"&gt;", "[PIPEEND]&gt;") + "' glossaryTerm='" + termDef + "' class='global-glossary'";
                            if (SharedObjects.glossaryList[termDef].glossaryImage != null)
                            {
                                glossDef = glossDef + "' glossaryImage='./assets/images/" + SharedObjects.glossaryList[termDef].glossaryImage;
                            }
                            if (SharedObjects.glossaryList[termDef].glossaryIframe != null)
                            {
                                glossDef = glossDef + " glossaryIframe='" + SharedObjects.glossaryList[termDef].glossaryIframe;
                            }
                            glossDef = glossDef + "'>" + termTxt + "</a>";
                        }
                        else
                        {
                            string textStr = "[Page Id: " + page.id + "]\t[Template Id: " + cmp.TemplateID + "]\t[Template Name: " + cmp.TemplateName + "]\tTerm: [" + termTxt + "]";
                            SharedObjects.missingglossaries.Add(textStr);
                        }

                        def = def.Replace(match.Value, glossDef);
                    }

                    if (text.Contains("pop-up|"))
                    {
                        string popupUrl_data = text.Split('|')[1].Trim().Replace("url=", "").Replace("url =", "").Trim();
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(popupUrl_data);
                        popupUrl_data = doc.DocumentNode.InnerText;
                        bool urltype = false;
                        if ((popupUrl_data.Contains("http")) || (popupUrl_data.Trim().StartsWith("www")))
                        {
                            urltype = true;
                        }

                        string popupBtn = text.Split('|')[text.Split('|').Length - 1].Trim();
                        string popupTerm = "<span id='renderPopup' popupUrl='" + popupUrl_data + "' popupBtn='" + popupBtn + "'></span>";
                        if (urltype == false)
                        {
                            popupTerm = "<span id='renderPopup' popupData='" + popupUrl_data + "' popupBtn='" + popupBtn + "'></span>";
                        }

                        def = def.Replace(match.Value, popupTerm);
                    }
                    if (text.Contains("image|"))
                    {
                        string alt = text.Split('|')[text.Split('|').Length - 2].Split('=')[1].Trim();
                        string imageName = text.Split('|')[text.Split('|').Length - 1].Trim().Replace("FFN","").Replace(":", "").Trim();
                        //<img id='expand-image' src=""/>
                        string imageTag = "<img id='expand-image' src='./assets/images/"+ imageName + "' alt='" + alt + "'/>";
                        def = def.Replace(match.Value, imageTag);
                    }
                    if (text.Contains("insert|media"))
                    {
                        PEA_Docx_to_Widget.TranscriptClass.Transcript transcript = new PEA_Docx_to_Widget.TranscriptClass.Transcript();
                        string transcriptText = text.Split('|')[text.Split('|').Length - 2].Split('=')[1].Trim();
                        string assetName = text.Split('|')[text.Split('|').Length - 1].Trim().Replace("FFN", "").Replace(":", "").Trim();
                        string id = "AudioTranscriptID_" + SharedObjects.TranscriptList.Count;
                        if (Path.GetExtension(assetName).ToLower() == ".mp3")
                        {
                            string audioTag = "<audio src='./assets/audio/"+Path.GetFileName(assetName)+"' controls controlsList='nodownload' class='audio-container'></audio><span id='read-transcript' transcriptID='"+ id + "'></span>";
                            def = def.Replace(match.Value, audioTag);
                        }
                        transcript.id = id;
                        transcript.Transcript_Text = transcriptText;
                        SharedObjects.TranscriptList.Add(id, transcript);
                    }
                }
            }
            
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.OptionWriteEmptyNodes = true;
            hDoc.LoadHtml(def);
            if (def.Contains("math"))
            {
                HtmlAgilityPack.HtmlNodeCollection mathNodes = hDoc.DocumentNode.SelectNodes("//span[contains(@class,'math')]");
                if (mathNodes != null)
                {
                    foreach (HtmlNode node in mathNodes.ToList())
                    {
                        string latex = node.InnerHtml.Replace("\\", "\\\\").Replace("$$", "");
                        if (node.Attributes.Contains("id"))
                        {
                            node.Attributes["id"].Value = "mathjax";
                        }
                        else { node.Attributes.Add("id", "mathjax"); }
                        node.Attributes.Add("latex", latex);
                        node.InnerHtml = "";
                        //node.Attributes.Add()
                    }
                }
            }

            HtmlAgilityPack.HtmlNodeCollection tableNodes = hDoc.DocumentNode.SelectNodes("//table");
            if (tableNodes != null)
            {
                foreach (HtmlNode node in tableNodes.ToList())
                {
                    string id = node.Attributes["id"].Value;

                    if (SharedObjects.TablesStyled.ContainsKey(id))
                    {
                        HtmlNode newNode = hDoc.CreateElement("div");
                        newNode.Attributes.Add("class", "tblcenter topbot");
                        HtmlNode htmlNode = SharedObjects.TablesStyled[id];

                        HtmlAgilityPack.HtmlDocument hDocTable = new HtmlAgilityPack.HtmlDocument();
                        hDocTable.OptionWriteEmptyNodes = true;
                        hDocTable.LoadHtml(htmlNode.OuterHtml);
                        
                        HtmlAgilityPack.HtmlNodeCollection styleNodes = hDocTable.DocumentNode.SelectNodes("//*[@style]");
                        if (styleNodes != null)
                        {
                            foreach (HtmlNode StNode in styleNodes.ToList())
                            {
                                string skipProp = ConfigurationManager.AppSettings.Get("SkipTableStyle");
                                string newStyle = "";
                                string[] styles = StNode.Attributes["style"].Value.Split(';');
                                foreach (string style in styles)
                                {
                                    if (style.Length > 0)
                                    {
                                        if (style.Trim().ToLower().StartsWith("color"))
                                        {
                                            System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml(style.Trim().ToLower().Split(':')[1].Trim());
                                            string colorName = GetColorName(color);
                                            if (colorName.ToLower() != "black")
                                            {
                                                newStyle = newStyle + ";" + style;
                                            }
                                        }
                                        else 
                                        {
                                            if (skipProp.Contains("|"+style.Trim().ToLower().Split(':')[0]+"|"))
                                            {
                                                newStyle = newStyle+";"+style;
                                            }
                                        }
                                    }
                                }
                                if (newStyle == "")
                                {
                                    StNode.Attributes.Remove("style");
                                }
                                else {
                                    StNode.Attributes.Remove("style");
                                    StNode.Attributes.Add("style",newStyle.Trim(';'));
                                }
                            }
                        }

                        //HtmlAgilityPack.HtmlNodeCollection spans = hDocTable.DocumentNode.SelectNodes("//span");
                        //if (spans != null)
                        //{
                        //    foreach (HtmlNode StNode in spans.ToList())
                        //    {
                        //        if (StNode.Attributes.Count == 0)
                        //        {
                        //            StNode.Name = "temp";
                        //        }
                        //        if (StNode.InnerHtml.Trim().Length == 0)
                        //        {
                        //            StNode.Attributes.RemoveAll();
                        //            StNode.Name = "temp";
                        //        }
                        //    }
                        //}

                        node.InnerHtml = hDocTable.DocumentNode.FirstChild.InnerHtml.Replace("<temp>","").Replace("</temp>", "").Replace("\r","").Replace("\n", "").Replace("\t", "");
                        newNode.InnerHtml = node.OuterHtml;
                        node.ParentNode.ReplaceChild(newNode, node);
                    }
                }
            }

            if (def.Contains("<math"))
            {
                HtmlAgilityPack.HtmlNodeCollection mathNodes = hDoc.DocumentNode.SelectNodes("//math");
                if (mathNodes != null)
                {
                    foreach (HtmlNode node in mathNodes.ToList())
                    {

                        string mathml = node.OuterHtml;
                        HtmlAgilityPack.HtmlDocument hDoc1 = new HtmlAgilityPack.HtmlDocument();
                        hDoc1.OptionWriteEmptyNodes = true;
                        hDoc1.LoadHtml(mathml);

                        HtmlNode annot = hDoc1.DocumentNode.SelectSingleNode("//annotation");
                        if (annot != null)
                        {
                            annot.ParentNode.RemoveChild(annot);
                        }
                        mathml = hDoc1.DocumentNode.InnerHtml;
                        //File.WriteAllText(System.Windows.Forms.Application.StartupPath + "\\math.mml", mathml);
                        //XslCompiledTransform myXslTrans = new XslCompiledTransform();
                        //myXslTrans.Load(AppDomain.CurrentDomain.BaseDirectory + "\\xsltml_1.0\\mmltex.xsl"); //this is the stylesheet
                        //myXslTrans.Transform(System.Windows.Forms.Application.StartupPath + "\\math.mml", System.Windows.Forms.Application.StartupPath + "\\math.tex");


                        System.Xml.Xsl.XslCompiledTransform transform = new XslCompiledTransform();
                        XsltSettings settings = new XsltSettings(true, false);
                        XmlUrlResolver resolver = new XmlUrlResolver();
                        transform.Load(AppDomain.CurrentDomain.BaseDirectory + "\\xsltml_1.0\\mmltex.xsl", settings, resolver);

                        StringBuilder sb = new StringBuilder();
                        using (StringReader sr = new StringReader(mathml))
                        using (XmlReader reader = XmlReader.Create(sr))
                        {
                            using (StringWriter writer = new StringWriter(sb))
                            {
                                XsltArgumentList xsltArgumentList = new XsltArgumentList();
                                transform.Transform(reader, null, writer);
                            }
                        }
                        string latex = sb.ToString().Trim().Trim('$');
                        //SharedObjects.latexes.Add("[LATEXCODE-" + SharedObjects.latexes.Count + 1 + "]", latex);
                        HtmlNode newNode = hDoc.CreateElement("span");
                        newNode.Attributes.Add("id", "mathjax");
                        newNode.Attributes.Add("latex", latex);
                        if (node.ParentNode.Name == "strong")
                        {
                            HtmlNode strongNode = node.ParentNode;
                            strongNode.ParentNode.ReplaceChild(newNode, strongNode);
                        }
                        else
                        {
                            node.ParentNode.ReplaceChild(newNode, node);
                        }
                    }
                }
            }
            HtmlAgilityPack.HtmlNodeCollection liNodes = hDoc.DocumentNode.SelectNodes("//li[@style]|//ul[@style]|//ol[@style]");
            if (liNodes != null)
            {
                foreach (HtmlNode node in liNodes.ToList())
                {
                    node.Attributes.Remove("style");
                }
            }
            HtmlAgilityPack.HtmlNodeCollection alinkNodes = hDoc.DocumentNode.SelectNodes("//a");
            if (alinkNodes != null)
            {
                foreach (HtmlNode anode in alinkNodes)
                {
                    if (anode.Attributes.Contains("href"))
                    {
                        if (!anode.Attributes.Contains("target"))
                            anode.Attributes.Add("target", "_blank");
                    }
                }
            }

            HtmlAgilityPack.HtmlNodeCollection imgNodes = hDoc.DocumentNode.SelectNodes("//img");
            if (imgNodes != null)
            {
                foreach (HtmlNode img in imgNodes.ToList())
                {
                    string src = img.Attributes["src"].Value;
                    string id = "";
                    if ((img.Attributes.Contains("id")) && (img.Attributes["id"].Value == "expand-image"))
                    { id = img.Attributes["id"].Value; }
                    foreach (HtmlAttribute att in img.Attributes.ToList())
                    {
                        if ((att.Name != "src"))
                        {
                            if ((att.Name != "alt"))
                            {
                                att.Remove();
                            }
                        }
                        else
                        {
                            string fileName = Path.GetFileName(src);
                            string newsrc = "./assets/images/" + fileName.Replace("jpeg", "jpg");
                            //if (!Directory.Exists(imagesDir))
                            //   Directory.CreateDirectory(imagesDir);

                            //File.Copy(src, imagesDir + "\\" + fileName, true);
                            img.Attributes[att.Name].Value = newsrc;
                            img.Attributes.Add("class", "topbot");
                            //if (img.Attributes.Contains("style"))
                            //    img.Attributes.Remove("style");
                        }
                    }
                    if (id != "")
                    {
                        img.Attributes.Add("id", id);
                    }
                    if (img.ParentNode.Name != "figure")
                    {
                        HtmlNode newNode = hDoc.CreateElement("figure");
                        HtmlAttribute att1 = hDoc.CreateAttribute("class");
                        att1.Value = "imgcenter";
                        string innXml = img.OuterHtml;
                        newNode.InnerHtml = innXml;
                        newNode.Attributes.Append(att1);
                        img.ParentNode.ReplaceChild(newNode, img);
                    }
                    //if(img.ParentNode!=null)
                    //{
                    //    HtmlNode newNode = hDoc.CreateElement("br");
                    //    newNode.InnerHtml = img.OuterHtml;
                    //    img.ParentNode.ReplaceChild(newNode, img);
                    //}
                }
            }
            HtmlAgilityPack.HtmlNodeCollection spanNodes = hDoc.DocumentNode.SelectNodes("//span[@class='underline']");
            if (spanNodes != null)
            {
                foreach (HtmlNode node in spanNodes.ToList())
                {
                    node.Name = "u";
                    node.Attributes.RemoveAll();
                }
                def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'");
            }
            HtmlAgilityPack.HtmlNodeCollection blankNodes = hDoc.DocumentNode.SelectNodes("//span[not(@*)]");
            if (blankNodes != null)
            {
                foreach (HtmlNode node in blankNodes.ToList())
                {
                    if (node.InnerText.Trim().Length == 0)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
                def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'");
            }
            HtmlAgilityPack.HtmlNodeCollection h2Nodes = hDoc.DocumentNode.SelectNodes("//h2");

            if (h2Nodes != null)
            {
                foreach (HtmlNode node in h2Nodes)
                {
                    node.Attributes.RemoveAll();
                    node.Name = "temp";
                }
                def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'").Replace("</br>", "").Replace("<br>", "<br/>");
            }
            else
            {
                def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'").Replace("</br>", "").Replace("<br>", "<br/>");
            }
            //HtmlNode firstNode = hDoc.DocumentNode.SelectSingleNode("/*//parent");
            //if (firstNode != null)
            //{
            //    if (firstNode.Name != "img")
            //    {
            //        if (firstNode.Name == "table") { def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'"); }
            //        else
            //        {
            //            def = firstNode.InnerHtml.Replace("\"", "'");
            //        }
            //    }
            //    else { def = hDoc.DocumentNode.InnerHtml.Replace("\"", "'"); }
            //}

            if (addPara == true)
            {
                if (!def.Trim().StartsWith("<p"))
                    def = "<p>" + def + "</p>";
            }
            if (def.Contains("<a"))
            {
                HtmlAgilityPack.HtmlNodeCollection aNodes = hDoc.DocumentNode.SelectNodes("//a");
                if (aNodes != null)
                {
                    if (aNodes[0].Attributes.Contains("href"))
                    {
                        string link = aNodes[0].Attributes["href"].Value;
                        if (aNodes[0].InnerText.Trim().Length == hDoc.DocumentNode.InnerText.Trim().Length)
                        {
                            def = link;
                        }
                    }
                }
            }

            return def.Replace("\"", "'");
        }
        private static string GetImageName(string def, string imagesDir)
        {
            HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
            hDoc.LoadHtml(def);
            HtmlAgilityPack.HtmlNodeCollection imgNodes = hDoc.DocumentNode.SelectNodes("//img");
            if (imgNodes != null)
            {
                foreach (HtmlNode img in imgNodes)
                {
                    string src = img.Attributes["src"].Value;
                    foreach (HtmlAttribute att in img.Attributes.ToList())
                    {
                        if (att.Name != "src")
                        {
                            att.Remove();
                        }
                        else
                        {
                            string fileName = Path.GetFileName(src);
                            def = fileName;
                        }
                    }
                }
            }
            return def.Replace("\"", "'").Replace("\r", "").Replace("\n", "").Replace("\t", "");
        }
        private static string CleanNodes(string json)
        {
            string updatedjson = json.Replace("\"Panels\": null,", "")
                                     .Replace("\"Slides\": null,", "")
                                     .Replace("\"Main_Texts\": null,", "")
                                     .Replace("\"Tabs\": null,", "")
                                     .Replace("\"Hotspots\": null,", "")
                                     .Replace("\"Texts\": null,", "")
                                     .Replace("\"Cards\": null,", "")
                                     .Replace("\"Notes\": null,", "")
                                     .Replace("\"Transcript\": null,", "")
                                     .Replace("\"SRT_VTT\": null,", "")
                                     .Replace("\"Media_position\": null,", "")
                                     .Replace("\"Media_position\": null,", "");
            return json;
        }
        private static string GetTitleName(string htmlPath)
        {
            string title = null;
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(File.ReadAllText(htmlPath));
            HtmlNode bodyTag = document.DocumentNode.SelectSingleNode("//body");
            HtmlAgilityPack.HtmlNodeCollection Nodes = bodyTag.SelectNodes("//body//*");
            if (Nodes != null)
            {
                foreach (HtmlNode node in Nodes)
                {
                    if (node.InnerText.Contains("Title:"))
                    {
                        title = node.InnerText.Replace("Title:", "").Trim();
                        break;
                    }
                }
            }
            if(title!=null)
            {
                title = title.Replace("\"", "'");
            }
            return title;
        }
        private static string validation(Page.Root jsonRoot, string pageName)
        {
            StringBuilder strErrors = new StringBuilder();
            StringBuilder strWarnings = new StringBuilder();
            StringBuilder Main = new StringBuilder();

            Main.Append("[Screen Name]: " + pageName + Environment.NewLine);
            if (jsonRoot != null)
            {
                var components = jsonRoot.components;
                if (components != null)
                {
                    foreach (Component cmp in components)
                    {
                        Main.Append("[Template Name]: " + cmp.TemplateName + Environment.NewLine);
                        Dictionary<string, string> logs = ApplyValidationChecks(cmp);
                        if (logs.Count > 0)
                        {
                            if (logs.ContainsKey("Errors"))
                            {
                                int errorcount = logs["Errors"].Split('\n').Length - 1;
                                Main.Append("Errors: [" + errorcount.ToString() + "]");
                                Main.Append(logs["Errors"] + Environment.NewLine);
                            }
                            if (logs.ContainsKey("Warnings"))
                            {
                                int warningcount = logs["Warnings"].Split('\n').Length - 1;
                                Main.Append(Environment.NewLine + "Warnings: [" + warningcount.ToString() + "]");
                                Main.Append(logs["Warnings"] + Environment.NewLine);
                            }
                            Main.Append("*********************************************************************************************************************************" + Environment.NewLine + Environment.NewLine);
                        }
                        else
                        {
                            Main.Append("No Error and warnings." + Environment.NewLine);
                            Main.Append("*********************************************************************************************************************************" + Environment.NewLine + Environment.NewLine);
                        }
                    }
                }
            }

            return Main.ToString();
        }
        private static Dictionary<string, string> ApplyValidationChecks(Component cmp)
        {
            string errors = "";
            string warns = "";
            Dictionary<string, string> validationCheck = new Dictionary<string, string>();
            string json = JsonConvert.SerializeObject(cmp, Newtonsoft.Json.Formatting.None,
new JsonSerializerSettings
{
    NullValueHandling = NullValueHandling.Ignore
});
            foreach (var key in ConfigurationManager.AppSettings)
            {
                string value = ConfigurationManager.AppSettings.Get(key.ToString());
                if (key.ToString().ToLower() == cmp.TemplateName.Replace("&amp;", "and").Replace("&", "and").ToLower())
                {
                    string[] mandatory = value.Split('|')[0].Split(';');
                    string[] optional = value.Split('|')[1].Split(';');
                    if (mandatory.Length > 0)
                    {
                        foreach (string mand in mandatory)
                        {
                            string field = mand.Replace("Mandatory:", "").Replace("mandatory:", "").Trim();
                            if (field.Length > 0)
                            {
                                if (!json.Contains("\"" + field + "\""))
                                {
                                    errors = errors + "\n" + "Error: '" + mand + "' property is missing from the component.";
                                }
                                else
                                {
                                    if ((json.Replace(": ", ":").Contains("\"" + field + "\"" + ":" + "\"" + "\"")) || (json.Contains("\"" + field + "\"" + ":" + "'" + "'")))
                                    {
                                        errors = errors + "\n" + "Error: '" + mand + "' value is missing from the component.";
                                    }
                                }
                            }
                        }
                    }
                    if (optional.Length > 0)
                    {
                        foreach (string opt in optional)
                        {
                            string field = opt.Replace("Optional:", "").Replace("optional:", "").Trim();
                            if (field.Length > 0)
                            {
                                if (!json.Contains("\"" + field + "\""))
                                {
                                    warns = warns + "\n" + "Warning: '" + opt + "' property is missing from the component.";
                                }
                                else
                                {
                                    if ((json.Replace(": ", ":").Contains("\"" + field + "\"" + ":" + "\"" + "\"")) || (json.Contains("\"" + field + "\"" + ":" + "'" + "'")))
                                    {
                                        warns = warns + "\n" + "Warning: '" + opt + "' value is missing from the component.";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (errors.Length > 0)
            {
                validationCheck.Add("Errors", errors);
            }
            if (warns.Length > 0)
            {
                validationCheck.Add("Warnings", warns);
            }
            return validationCheck;
        }
        public static string GetColorName(System.Drawing.Color color)
        {
            var colorProperties = typeof(System.Drawing.Color)
                .GetProperties(BindingFlags.Public | BindingFlags.Static)
                .Where(p => p.PropertyType == typeof(System.Drawing.Color));
            foreach (var colorProperty in colorProperties)
            {
                var colorPropertyValue = (System.Drawing.Color)colorProperty.GetValue(null, null);
                if (colorPropertyValue.R == color.R
                       && colorPropertyValue.G == color.G
                       && colorPropertyValue.B == color.B)
                {
                    return colorPropertyValue.Name;
                }
            }

            //If unknown color, fallback to the hex value
            //(or you could return null, "Unkown" or whatever you want)
            return ColorTranslator.ToHtml(color);
        }
        private System.Drawing.Color GetSystemDrawingColorFromHexString(string hexString)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(hexString, @"[#]([0-9]|[a-f]|[A-F]){6}\b"))
                throw new ArgumentException();
            int red = int.Parse(hexString.Substring(1, 2), NumberStyles.HexNumber);
            int green = int.Parse(hexString.Substring(3, 2), NumberStyles.HexNumber);
            int blue = int.Parse(hexString.Substring(5, 2), NumberStyles.HexNumber);
            return System.Drawing.Color.FromArgb(red, green, blue);
        }
    }
}
