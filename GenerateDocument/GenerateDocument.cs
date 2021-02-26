using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Xsl;

namespace GenerateDocument
{
    class GenerateDocument
    {
        private string utilPath { get; set; }
        private string sourcePath { get; set; }
        public GenerateDocument(string sourcePath, string utilPath)
        {
            this.sourcePath = sourcePath;
            this.utilPath = utilPath;
        }
        public void RunCommand(string command, string filename)
        {
            string name = filename.Substring(0, filename.IndexOf("."));

            string resultPath = "";

            #region generateXSLT
            XmlDocument xsl = new XmlDocument();
            if (filename.EndsWith(".xslt"))
            {
                xsl.Load(filename);
            }
            else
            {
                if (command.StartsWith("pdf") || command.StartsWith("docx") || command.StartsWith("xslt"))
                {
                    xsl = GenerateXSLT(filename);

                    #region save xsl
                    //if (command.Contains(":s"))
                    if (true)
                    {
                        string xslPath = sourcePath + name + ".xslt";
                        SaveXml(xsl, xslPath);
                    }
                    #endregion
                }
            }
            #endregion

            #region generateDOCX
            string docx = "";
            if (command.StartsWith("pdf") || command.StartsWith("docx"))
            {
                string xmlPath = GenerateXML("EmptyXML");
                docx = GenerateFile(xsl, xmlPath, name);
                resultPath = docx;
            }
            #endregion

            #region generatePDF
            if (command.StartsWith("pdf"))
            {
                resultPath = Convert(docx, name, WdSaveFormat.wdFormatPDF);
            }
            #endregion

            OpenFile(command, resultPath);
        }
        public string GenerateXML(string name)
        {
            string xmlPath = name + ".xml";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<Root />");
            CreateDir();
            SaveXml(xmlDoc, sourcePath + xmlPath, false);
            return sourcePath + xmlPath;
        }
        private string GenerateFile(string xsltFile, string xmlDataFile, string name)
        {
            XmlDocument xsl = new XmlDocument();
            xsl.Load(xsltFile);
            return GenerateFile(xsl, xmlDataFile, name);
        }
        private string GenerateFile(XmlDocument xslt, string xmlDataFile, string name)
        {
            //https://msdn.microsoft.com/en-us/library/ee872374(v=office.12).aspx

            string templateDocument = utilPath + name + ".docx";
            string outputDocument = sourcePath + name + ".docx";

            try
            {
                StringWriter stringWriter = new StringWriter();
                XmlWriter xmlWriter = XmlWriter.Create(stringWriter);

                XslCompiledTransform transform = new XslCompiledTransform();
                transform.Load(xslt);
                transform.Transform(xmlDataFile, xmlWriter);

                XmlDocument newWordContent = new XmlDocument();
                newWordContent.LoadXml(stringWriter.ToString());

                #region Remove Empty sdtContents

                #region XmlNamespaceManager
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(newWordContent.NameTable);
                nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                nsmgr.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");
                #endregion

                XmlNodeList sdts = newWordContent.SelectNodes("//w:sdt", nsmgr);
                foreach (XmlNode sdt in sdts)
                {
                    XmlNode sdtContent = sdt.SelectSingleNode(".//w:sdtContent", nsmgr);
                    if (sdtContent.InnerXml == "")
                    {
                        sdt.ParentNode.RemoveChild(sdt);
                    }
                }
                #endregion

                //System.IO.File.Copy(templateDocument, outputDocument, true);
                System.IO.File.Copy(Path.Combine(utilPath, name + ".docx"), Path.Combine(sourcePath, name + ".docx"), true);
                System.IO.File.SetAttributes(outputDocument, FileAttributes.Normal);

                using (WordprocessingDocument output = WordprocessingDocument.Open(outputDocument, true))
                {
                    Body updatedBodyContent = new Body(newWordContent.DocumentElement.InnerXml);
                    output.MainDocumentPart.Document.Body = updatedBodyContent;
                    output.MainDocumentPart.Document.Save();
                    //output.Dispose();
                }
            }
            catch (Exception e)
            {
                throw new Exception("GenerateFile: " + e.Message);
                //throw e;
            }
            return outputDocument;
        }
        private string Convert(string input, string output, WdSaveFormat format)
        {
            string path = output;

            try
            {
                if (format.Equals(WdSaveFormat.wdFormatPDF))
                {
                    path = path + ".pdf";
                    string[] fList = Directory.GetFiles(sourcePath, path);
                    foreach (string f in fList)
                    {
                        System.IO.File.Delete(f);
                    }
                }

                Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = false;
                object oMissing = System.Reflection.Missing.Value;
                object isVisible = true;
                object readOnly = false;
                object oInput = input;
                object oOutput = sourcePath + path;
                object oFormat = format;

                object saveOption = WdSaveOptions.wdDoNotSaveChanges;
                object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
                object routeDocument = false;

                Microsoft.Office.Interop.Word._Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                try
                {
                    oDoc.Activate();
                    oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    //Microsoft.Office.Interop.Word.WdStatistic stat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
                    //int num = oDoc.ComputeStatistics(stat, ref oMissing); ---WordPages.Count
                }
                catch (Exception e)
                {
                    throw new Exception("SaveAs: " + e.Message);
                }
                finally
                {
                    //oWord.NormalTemplate.Saved = true;
                    //oWord.Documents.Close();
                    oDoc.Application.Quit(saveOption, ref oMissing, ref oMissing);
                    //oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ConvertToPDF: " + ex.Message);
            }
            return sourcePath + path;
        }
        public string Generate(string xsl, string xml)
        {
            string ext = Path.GetExtension(xsl);
            CheckFileFormat(ext, ".xslt");
            string name = xsl.Substring(0, xsl.IndexOf(ext));

            string docx = GenerateFile(xsl, xml, name);
            string pdf = Convert(docx, name, WdSaveFormat.wdFormatPDF);
            return pdf;
        }

        #region Generate XSLT and WordDocumentXml

        private XmlDocument GetXMLfromXSLT(string xsl)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(xsl);

            #region XmlNamespaceManager

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xdoc.NameTable);
            nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            nsmgr.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");

            #endregion

            #region remove if condition

            XmlNodeList ifConds = xdoc.SelectNodes("//w:sdt/w:sdtContent/xsl:if", nsmgr);

            do
            {
                ifConds = xdoc.SelectNodes("//w:sdt/w:sdtContent/xsl:if", nsmgr);
                if (ifConds.Count > 0)
                {
                    ifConds[0].ParentNode.InnerXml = ifConds[0].InnerXml;
                }

            } while (ifConds.Count > 0);
            #endregion

            #region remove loop

            XmlNodeList loops = xdoc.SelectNodes("//w:sdt/w:sdtContent/xsl:for-each", nsmgr);

            do
            {
                loops = xdoc.SelectNodes("//w:sdt/w:sdtContent/xsl:for-each", nsmgr);
                if (loops.Count > 0)
                {
                    loops[0].ParentNode.InnerXml = loops[0].InnerXml;
                }

            } while (loops.Count > 0);
            #endregion

            #region set default text

            XmlNodeList texts = xdoc.SelectNodes("//w:r/w:t/xsl:value-of", nsmgr);
            foreach (XmlNode xnode in texts)
            {
                xnode.ParentNode.InnerText = xnode.Attributes["select"].Value;
            }

            #endregion

            #region set default checkbox

            XmlNodeList checkBoxes = xdoc.SelectNodes("//w:r/w:fldChar/w:ffData", nsmgr);
            foreach (XmlNode xnode in checkBoxes)
            {
                XmlNode valueof = xnode.SelectSingleNode("w:checkBox/xsl:value-of", nsmgr);

                XmlElement wdefault = xdoc.CreateElement("w", "default", nsmgr.LookupNamespace("w"));
                XmlAttribute wval = xdoc.CreateAttribute("w", "val", nsmgr.LookupNamespace("w"));
                wval.Value = "0";
                wdefault.Attributes.Append(wval);

                xnode.SelectSingleNode("w:checkBox", nsmgr).RemoveChild(valueof);
                xnode.SelectSingleNode("w:checkBox", nsmgr).AppendChild(wdefault);
            }

            #endregion

            xdoc.LoadXml(xdoc.SelectSingleNode("xsl:stylesheet/xsl:template", nsmgr).InnerXml);
            return xdoc;
        }
        public void GenerateXMLfromXSLT(string xsl)
        {
            string ext = Path.GetExtension(xsl);
            CheckFileFormat(ext, ".xslt");
            string xml = xsl.Substring(0, xsl.IndexOf(ext)) + "document.xml";

            XmlDocument xdoc = GetXMLfromXSLT(xsl);
            SaveXml(xdoc, xml);
        }
        public string GenerateDOCXfromXSLT(string xsl)
        {
            string ext = Path.GetExtension(xsl);
            CheckFileFormat(ext, ".xslt");
            string templateDocument = xsl.Substring(0, xsl.IndexOf(ext)) + ".docx";
            string outputDocument = xsl.Substring(0, xsl.IndexOf(ext)) + "Document.docx";

            XmlDocument xdoc = GetXMLfromXSLT(xsl);
            System.IO.File.Copy(templateDocument, outputDocument, true);
            using (WordprocessingDocument output = WordprocessingDocument.Open(outputDocument, true))
            {
                Body updatedBodyContent = new Body(xdoc.DocumentElement.InnerXml);
                output.MainDocumentPart.Document.Body = updatedBodyContent;
                output.MainDocumentPart.Document.Save();
            }
            return outputDocument;
        }
        public XmlDocument GenerateXSLT(string docx)
        {
            string ext = Path.GetExtension(docx);
            CheckFileFormat(ext, ".docx");
            string txt = docx.Substring(0, docx.IndexOf(ext)) + "XPath.txt";

            XmlDocument xdoc = new XmlDocument();
            List<string> xpaths = new List<string>();
            try
            {
                #region get main document template xslt

                #region XmlNamespaceManager

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xdoc.NameTable);
                nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                nsmgr.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");

                #endregion

                XmlElement stylesheet = xdoc.CreateElement("xsl", "stylesheet", nsmgr.LookupNamespace("xsl"));
                XmlAttribute version = xdoc.CreateAttribute("version");
                version.Value = "1.0";
                stylesheet.Attributes.Append(version);

                XmlElement template = xdoc.CreateElement("xsl:template", nsmgr.LookupNamespace("xsl"));
                XmlAttribute match = xdoc.CreateAttribute("match");
                match.Value = "/";
                template.Attributes.Append(match);

                stylesheet.AppendChild(template);
                xdoc.AppendChild(stylesheet);

                using (WordprocessingDocument output = WordprocessingDocument.Open(docx, false))
                {
                    DocumentFormat.OpenXml.Wordprocessing.Document bodyContent = output.MainDocumentPart.Document;
                    template.InnerXml = bodyContent.OuterXml;
                }

                #endregion

                #region remove graphics images

                XmlNodeList exts = xdoc.SelectNodes("//a:extLst", nsmgr);
                foreach (XmlNode xnode in exts)
                {
                    xnode.RemoveAll();
                }

                #endregion

                #region set value if conditions

                bool hasIfConditions = false;

                do
                {
                    hasIfConditions = false;

                    XmlNodeList sdtLoops = xdoc.SelectNodes("//w:sdt", nsmgr);
                    foreach (XmlNode sdtLoop in sdtLoops)
                    {
                        string xpath = "";
                        XmlNode sdtPrTag = sdtLoop.SelectSingleNode(".//w:sdtPr/w:tag", nsmgr);
                        XmlNode sdtContent = sdtLoop.SelectSingleNode(".//w:sdtContent", nsmgr);

                        if (sdtPrTag != null)
                        {
                            xpath = GetFieldXPath(sdtPrTag.Attributes["w:val"].Value);
                        }

                        if (!string.IsNullOrEmpty(xpath) && xpath.StartsWith("If/")
                            && sdtContent.HasChildNodes && sdtContent.FirstChild.Name != "xsl:if")
                        {
                            xpath = GetFieldsXPath(xpath);

                            XmlElement ifCond = getXslIf(xdoc, xpath, nsmgr.LookupNamespace("xsl"));
                            ifCond.InnerXml = sdtContent.InnerXml;
                            sdtContent.InnerXml = ifCond.OuterXml;
                            xpaths.Add(xpath);
                            hasIfConditions = true;
                            break;
                        }
                    }

                } while (hasIfConditions);
                #endregion

                #region set value for loops

                bool hasForeach = false;

                do
                {
                    hasForeach = false;

                    XmlNodeList sdtLoops = xdoc.SelectNodes("//w:sdt", nsmgr);
                    foreach (XmlNode sdtLoop in sdtLoops)
                    {
                        string xpath = "";
                        XmlNode sdtPrTag = sdtLoop.SelectSingleNode(".//w:sdtPr/w:tag", nsmgr);
                        XmlNode sdtContent = sdtLoop.SelectSingleNode(".//w:sdtContent", nsmgr);

                        if (sdtPrTag != null)
                        {
                            xpath = GetFieldXPath(sdtPrTag.Attributes["w:val"].Value);
                        }

                        if (!string.IsNullOrEmpty(xpath) && (xpath.StartsWith("Roots/") || xpath.StartsWith("Lists/"))
                            && sdtContent.HasChildNodes && sdtContent.FirstChild.Name != "xsl:for-each")
                        {
                            xpath = GetFieldsXPath(xpath);

                            XmlElement forEach = getXslForEach(xdoc, xpath, nsmgr.LookupNamespace("xsl"));
                            forEach.InnerXml = sdtContent.InnerXml;
                            sdtContent.InnerXml = forEach.OuterXml;
                            xpaths.Add(xpath);
                            hasForeach = true;
                            break;
                        }
                    }

                } while (hasForeach);
                #endregion

                #region set value for texts

                XmlNodeList sdtTexts = xdoc.SelectNodes("//w:sdt", nsmgr);
                foreach (XmlNode sdtText in sdtTexts)
                {
                    string xpath = "";
                    XmlNode sdtPrTag = sdtText.SelectSingleNode(".//w:sdtPr/w:tag", nsmgr);
                    XmlNode sdtContent = sdtText.SelectSingleNode(".//w:sdtContent", nsmgr);

                    if (sdtPrTag != null)
                    {
                        xpath = GetFieldXPath(sdtPrTag.Attributes["w:val"].Value);
                    }
                    if (string.IsNullOrEmpty(xpath))
                    {
                        if (!string.IsNullOrEmpty(sdtContent.InnerText.Trim()))
                        {
                            xpath = GetFieldXPath(sdtContent.InnerText.Trim());
                        }
                    }

                    if (!string.IsNullOrEmpty(xpath) && !(xpath.StartsWith("Roots/") || xpath.StartsWith("Lists/") || xpath.StartsWith("If/")))
                    {
                        XmlNodeList otherRs = sdtContent.SelectNodes(".//w:t", nsmgr);
                        if (otherRs.Count == 0)
                        {
                            XmlNode paragraph = sdtContent.SelectSingleNode(".//w:p", nsmgr);
                            if (paragraph != null)
                            {
                                paragraph.AppendChild(getTextTag(xdoc, nsmgr.LookupNamespace("w")));
                            }
                            else
                            {
                                sdtContent.AppendChild(getTextTag(xdoc, nsmgr.LookupNamespace("w")));
                            }
                            otherRs = sdtContent.SelectNodes(".//w:t", nsmgr);
                        }
                        foreach (XmlNode xnoder in otherRs)
                        {
                            xnoder.InnerXml = "";
                        }
                        otherRs[0].InnerXml = getXslValueOf(xdoc, xpath, nsmgr.LookupNamespace("xsl")).OuterXml;
                        foreach (XmlNode xnoder in otherRs)
                        {
                            if (xnoder.InnerXml == "")
                            {
                                xnoder.ParentNode.ParentNode.RemoveChild(xnoder.ParentNode);
                            }
                        }
                        xpaths.Add(xpath);
                    }
                }

                #endregion

                #region set value for checkboxes

                XmlNodeList checkBoxes = xdoc.SelectNodes("//w:r/w:fldChar/w:ffData", nsmgr);
                foreach (XmlNode xnode in checkBoxes)
                {
                    string pathname = xnode.SelectSingleNode("w:name", nsmgr).Attributes["w:val"].Value;
                    string xpath = getPathFromName(pathname);
                    xpaths.Add(xpath);

                    XmlNode wdefault = xnode.SelectSingleNode("w:checkBox/w:default", nsmgr);
                    xnode.SelectSingleNode("w:checkBox", nsmgr).RemoveChild(wdefault);
                    xnode.SelectSingleNode("w:checkBox", nsmgr).AppendChild(getXslValueOf(xdoc, xpath, nsmgr.LookupNamespace("xsl")));
                }

                #endregion

                #region Save xpath list

                List<string> list = xpaths.ToList();
                list.Sort();
                CreateDir();
                SaveList(list, sourcePath + txt);

                #endregion

                return xdoc;
            }
            catch (Exception ex)
            {
                throw new Exception("GenerateXSLT: " + ex.Message);
            }
        }
        public void SaveXml(XmlDocument xdoc, string path, bool update = true)
        {
            if (update || !System.IO.File.Exists(path))
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(path);
                file.WriteLine(xdoc.OuterXml);
                file.Close();
            }
        }
        public void SaveList(List<string> list, string path)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(path);
            try
            {
                foreach (string item in list)
                {
                    file.WriteLine(item);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                file.Close();
            }
        }
        public int CheckValidSymbol(string path)
        {
            int result = -1;
            int[] arr = new int[] { 39, 42, 43, 45, 46, 47, 48, 49, 50, 51, 52,
                                    53, 54, 55, 56, 57, 60, 61, 62, 64, 65,
                                    66, 67, 68, 69, 70, 71, 72, 73, 74, 75,
                                    76, 77, 78, 79, 80, 81, 82, 83, 84, 85,
                                    86, 87, 88, 89, 90, 95, 97, 98, 99, 100,
                                    101, 102, 103, 104, 105, 106, 107, 108,
                                    109, 110, 111, 112, 113, 114, 115, 116,
                                    117, 118, 119, 120, 121, 122, 124
            };

            bool[] bb = new bool[128]{ false, false, false, false, false, false, false, false, false, false,
                                        false, false, false, false, false, false, false, false, false, false,
                                        false, false, false, false, false, false, false, false, false, false,
                                        false, false, false, false, false, false, false, false, false, true,
                                        false, false, true, true, false, true, true, true, true, true,
                                        true, true, true, true, true, true, true, true, false, false,
                                        true, true, true, false, true, true, true, true, true, true,
                                        true, true, true, true, true, true, true, true, true, true,
                                        true, true, true, true, true, true, true, true, true, true,
                                        true, false, false, false, false, true, false, true, true, true,
                                        true, true, true, true, true, true, true, true, true, true,
                                        true, true, true, true, true, true, true, true, true, true,
                                        true, true, true, false, true, false, false, false };

            for (int i = 0; i < path.Count(); i++)
            {
                int acsiiCode = (int)path[i];
                if (acsiiCode > 0 && acsiiCode > 128 || !bb[acsiiCode])
                {
                    result = i;
                    break;
                }
            }

            return result;
        }


        private void CreateDir()
        {
            if (!System.IO.Directory.Exists(sourcePath))
            {
                System.IO.Directory.CreateDirectory(sourcePath);
            }
        }
        public XmlElement getXslValueOf(XmlDocument xdoc, string path, string namespaceUri)
        {
            GetSymbolResult(path);

            XmlElement valueof = xdoc.CreateElement("xsl", "value-of", namespaceUri);
            XmlAttribute select = xdoc.CreateAttribute("select");
            select.Value = path;
            valueof.Attributes.Append(select);
            return valueof;
        }
        public XmlElement getTextTag(XmlDocument xdoc, string namespaceUri)
        {
            XmlElement r = xdoc.CreateElement("w", "r", namespaceUri);
            XmlElement rPr = xdoc.CreateElement("w", "rPr", namespaceUri);

            XmlElement lang = xdoc.CreateElement("w", "lang", namespaceUri);
            XmlAttribute val = xdoc.CreateAttribute("w:val");
            val.Value = "en-US";
            lang.Attributes.Append(val);
            rPr.AppendChild(lang);
            r.AppendChild(rPr);

            XmlElement t = xdoc.CreateElement("w", "t", namespaceUri);
            r.AppendChild(t);

            return r;
        }

        public XmlElement getXslForEach(XmlDocument xdoc, string path, string namespaceUri)
        {
            GetSymbolResult(path);

            XmlElement valueof = xdoc.CreateElement("xsl", "for-each", namespaceUri);
            XmlAttribute select = xdoc.CreateAttribute("select");
            select.Value = path;
            valueof.Attributes.Append(select);
            return valueof;
        }
        public XmlElement getXslIf(XmlDocument xdoc, string path, string namespaceUri)
        {
            GetSymbolResult(path);

            XmlElement valueof = xdoc.CreateElement("xsl", "if", namespaceUri);
            XmlAttribute select = xdoc.CreateAttribute("test");
            select.Value = path;
            valueof.Attributes.Append(select);
            return valueof;
        }
        public void OpenFile(string command, string filepath)
        {
            if (command.Contains(":o") && !string.IsNullOrEmpty(filepath))
            {
                Process.Start(filepath);
            }
        }
        public void GetSymbolResult(string path)
        {
            int symbRes = CheckValidSymbol(path);
            if (symbRes > -1)
            {
                if (symbRes == 1)
                {
                    throw new Exception("'" + path + "' : " + symbRes + "-st symbol is invalid!");
                }
                else if (symbRes == 2)
                {
                    throw new Exception("'" + path + "' : " + symbRes + "-nd symbol is invalid!");
                }
                else if (symbRes == 3)
                {
                    throw new Exception("'" + path + "' : " + symbRes + "-rd symbol is invalid!");
                }
                else
                {
                    throw new Exception("'" + path + "' : " + symbRes + "-th symbol is invalid!");
                }
            }
        }
        #endregion

        #region DataValues
        private void CheckFileFormat(string ext, string extentsion)
        {
            if (ext != extentsion)
            {
                throw new Exception("File is invalid! Extension must be '" + extentsion + "'. See 'help'!");
            }
        }
        private string GetFieldXPath(string xpath)
        {
            xpath = xpath.Replace('\\', '/');
            if (xpath.EndsWith("/Root"))
            {
                xpath = xpath.Substring(0, xpath.LastIndexOf("/"));
            }
            return xpath;
        }
        private string GetFieldsXPath(string xpath)
        {
            if (xpath.StartsWith("Roots/"))
            {
                xpath = xpath.Replace("Roots", "Root");
            }
            if (xpath.StartsWith("Lists/"))
            {
                xpath = xpath.Substring(6);
            }
            if (xpath.StartsWith("If/"))
            {
                xpath = xpath.Substring(3);

                if (!xpath.Contains("="))
                {
                    throw new Exception("'" + xpath + "' : xpath is invalid!");
                }
                else
                {
                    string condValue = xpath.Substring(xpath.IndexOf("=") + 1);
                    if (!(condValue.Length > 1 && condValue[0] == '\'' && condValue[condValue.Length - 1] == '\''))
                    {
                        throw new Exception("'" + xpath + "' : xpath is invalid!");
                    }
                }
            }
            return xpath;
        }
        private string getBooleanValue(bool b)
        {
            return b ? "Да" : "Нет";
        }
        private string getBooleanValue(string b)
        {
            return b.Equals("true") ? "Да" : "Нет";
        }
        private string getShortDate(string date)
        {
            DateTime dateDateTime = DateTime.Parse(date, System.Globalization.CultureInfo.InvariantCulture);
            return dateDateTime.ToShortDateString();
        }
        private string getShortTime(string date)
        {
            DateTime dateDateTime = DateTime.Parse(date, System.Globalization.CultureInfo.InvariantCulture);
            return dateDateTime.ToShortTimeString();
        }
        private string getCurrency(string currency)
        {
            return currency == "0" ? "" : currency;
        }
        private string getMonthRus(int num)
        {
            string[] month = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
            //string[] month = { "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь" };
            return month[num - 1];
        }
        private string getMonthRus(string num)
        {
            int numm = Int32.Parse(num);
            return getMonthRus(numm);
        }
        private string getMonthKaz(int num)
        {
            string[] month = { "Қаңтар", "Ақпан", "Наурыз", "Сәуір", "Мамыр", "Маусым", "Шілде", "Тамыз", "Қыркүйек", "Қазан", "Қараша", "Желтоқсан" };
            //string[] month = { "қаңтар", "ақпан", "наурыз", "сәуір", "мамыр", "маусым", "шілде", "тамыз", "қыркүйек", "қазан", "қараша", "желтоқсан" };
            return month[num - 1];
        }
        private string getInnerText(XmlDocument doc, string path)
        {
            string innerText = "";
            if (doc.SelectSingleNode(path) != null)
                innerText = getInnerText(doc.SelectSingleNode(path));
            return innerText;
        }
        public string getInnerText(XmlNode node, string path)
        {
            string innerText = "";
            if (node != null)
                innerText = getInnerText(node.SelectSingleNode(path));
            return innerText;
        }
        private string getInnerText(XmlNodeList list, string path, int index)
        {
            string innerText = "";
            if (list.Count > index)
                innerText = getInnerText(list.Item(index), path);
            return innerText;
        }
        public string getInnerText(XmlNode node)
        {
            string innerText = "";
            if (node != null)
                innerText = node.InnerText;
            return innerText;
        }
        private string getPathFromName(string name)
        {
            string path = "";
            for (int i = 0; i < name.Length; i++)
            {
                if (name[i] >= 'A' && name[i] <= 'Z')
                {
                    path = path + "/" + name[i];
                }
                else
                {
                    path = path + name[i];
                }
            }
            return path.Substring(1);
        }

        private string getCheckBoxValue(int value)
        {
            string innerText = "<w:default w:val=\"" + value + "\"/>";
            return innerText;
        }
        private string getCheckBoxValue(bool value)
        {
            string innerText = getCheckBoxValue(System.Convert.ToInt32(value));
            return innerText;
        }
        #endregion

        public byte[] GenerateWord(string name, XmlDocument xData)
        {
            string templateXslt = utilPath + name + ".xslt";
            byte[] bytes;
            try
            {
                StringWriter stringWriter = new StringWriter();
                XmlWriter xmlWriter = XmlWriter.Create(stringWriter);

                XmlDocument xsl = new XmlDocument();
                xsl.Load(templateXslt);

                XslCompiledTransform transform = new XslCompiledTransform();
                transform.Load(xsl);
                transform.Transform(xData, xmlWriter);

                XmlDocument newWordContent = new XmlDocument();
                newWordContent.LoadXml(stringWriter.ToString());

                byte[] byteArray = File.ReadAllBytes(utilPath + name + ".docx");
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (WordprocessingDocument output = WordprocessingDocument.Open(stream, true))
                    {
                        Body updatedBodyContent = new Body(newWordContent.DocumentElement.InnerXml);
                        output.MainDocumentPart.Document.Body = updatedBodyContent;
                        output.MainDocumentPart.Document.Save();
                    }
                    bytes = stream.ToArray();
                }
            }
            catch (Exception e)
            {
                throw new Exception("GenerateWord: " + e.Message);
            }
            return bytes;
        }
        //public byte[] GeneratePdf(byte[] wordByte)
        //{
        //    Stream stream2 = new MemoryStream(wordByte);
        //    var asposeDocument = new Aspose.Words.Document(stream2);

        //    MemoryStream stream = new MemoryStream();
        //    asposeDocument.Save(stream, Aspose.Words.SaveFormat.Pdf);

        //    return stream.ToArray();
        //}
    }
}
