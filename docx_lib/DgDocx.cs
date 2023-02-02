namespace docx_lib;
using Markdig;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Text;
using System.IO.Compression;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System.Collections.Generic;
using Markdig.Syntax;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Vml.Office;

public class DgDocx
{
    private static IEnumerable<HyperlinkRelationship> hyperlinks;
    private static int linksCount = 0;
    public static void createWordprocessingDocument(string filepath)
    {
        // Create a document by supplying the filepath. 
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
        {
            // Add a main document part. 
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }

    }

    // stream here because anticipating zip.
    public async static Task md_to_docx(String md, Stream outputStream) //String mdFile, String docxFile, String template)
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

        var html = Markdown.ToHtml(md, pipeline);
        //edit on debug the h
        //All the document is being saved in the stream
        using (WordprocessingDocument doc = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();


            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.ParseHtml(html);
            mainPart.Document.Save();
        }
    }

    public async static Task md_to_docx(String[] mdFiles, String images, Stream[] outputStream) //String mdFile, String docxFile, String template)
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
        for (int i = 0; i < mdFiles.Length; i++)
        {
            var html = Markdown.ToHtml(mdFiles[i], pipeline);
            //edit on debug the h
            //All the document is being saved in the stream
            using (WordprocessingDocument doc = WordprocessingDocument.Create(outputStream[i], WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();


                HtmlConverter converter = new HtmlConverter(mainPart);
                converter.ParseHtml(html);
                mainPart.Document.Save();
            }
        }
    }

    public async static Task docx_to_md(Stream infile, Stream outfile, String name = "")
    {
        WordprocessingDocument wordDoc = WordprocessingDocument.Open(infile, false);
        DocumentFormat.OpenXml.Wordprocessing.Body body
        = wordDoc.MainDocumentPart.Document.Body;


        StringBuilder textBuilder = new StringBuilder();
        var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (parts != null)
        {
            //var asd = parts.Descendants<HyperlinkList>();


            foreach (var block in parts.ChildElements)
            {


                if (block is Paragraph)
                {
                    //This method is for manipulating the style of Paragraphs and text inside
                    ProcessParagraph((Paragraph)block, textBuilder);
                }

                if (block is Table) ProcessTable((Table)block, textBuilder);

            }
        }

        //This code is replacing the below one because I need to check the .md file faster
        //writing the .md file in test_result folder
        if (name != "")
        {
            using (var streamWriter = new StreamWriter(name + ".md"))
            {
                String s = textBuilder.ToString();
                streamWriter.Write(s);
            }
        }
        else
        {

            var writer = new StreamWriter(outfile);
            String s = textBuilder.ToString();
            writer.Write(s);
            writer.Flush();
        }

    }

    private static void ProcessTable(Table node, StringBuilder textBuilder)
    {
        List<string> headerDivision = new List<string>();
        int rowNumber = 0;

        foreach (var row in node.Descendants<TableRow>())
        {
            rowNumber++;

            if (rowNumber == 2)
            {
                headerDivider(headerDivision, textBuilder);
            }

            textBuilder.Append("| ");
            foreach (var cell in row.Descendants<TableCell>())
            {
                foreach (var para in cell.Descendants<Paragraph>())
                {
                    if (para.ParagraphProperties != null)
                    {
                        headerDivision.Add(para.ParagraphProperties.Justification.Val);
                    }
                    else
                    {
                        headerDivision.Add("normal");
                    }
                    textBuilder.Append(para.InnerText);
                }
                textBuilder.Append(" | ");
            }
            textBuilder.AppendLine("");
        }
    }

    private static String ProcessParagraphElements(Paragraph block)
    {
        String style = block.ParagraphProperties?.ParagraphStyleId?.Val;

        if (style == null)
        {
            style = "single";
            block.ParagraphProperties.AppendChild(new ParagraphStyleId() { Val = "single" });
        }

        int num;
        String prefix = "";
        if ("top" == block.ParagraphProperties?.ParagraphBorders?.TopBorder?.LocalName
            && null == block.ParagraphProperties?.ParagraphBorders?.BottomBorder
            && null == block.ParagraphProperties?.ParagraphBorders?.LeftBorder)
        {

            prefix += "---\n";
            return prefix;
        }
        //to find Heading Paragraphs
        if (style.Contains("Heading"))
        {
            num = int.Parse(style.Substring(style.Length - 1));


            for (int i = 0; i < num; i++)
            {
                prefix += "#";
            }
            return prefix;
        }

        //to find List Paragraphs
        if (style == "ListParagraph")
        {
            return prefix = "-";
        }

        //to find quotes Paragraphs
        if (style == "IntenseQuote")
        {
            return prefix = ">";
        }

        return prefix;
    }

    private static String ProcessRunElements(Run run)
    {
        //extract the child element of the text (Bold or Italic)
        OpenXmlElement expression = run.RunProperties.ChildElements.ElementAtOrDefault(0);

        String prefix = "";

        //to know if the propertie is Bold, Italic or both
        switch (expression)
        {
            case Bold:
                if (run.RunProperties.ChildElements.Count == 2)
                {
                    prefix = "***";
                    break;
                }
                prefix = "**";
                break;
            case Italic:
                prefix = "*";
                break;
        }
        return prefix;
    }

    private static String ProcessBlockQuote(Run block)
    {
        String text = block.InnerText;
        String[] textSliced = text.Split("\n");
        String textBack = "";

        foreach (String n in textSliced)
        {
            textBack += "> " + n + "\n";
        }

        return textBack;
    }

    private static void ProcessParagraph(Paragraph block, StringBuilder textBuilder)
    {
        String constructorBase = "";
        bool isEsc = false; //is an escape url

        //iterate along every element in the Paragraphs and childrens
        foreach (var run in block.Descendants<Run>())
        {
            String prefix = "";
            var links = block.Descendants<Hyperlink>();
            if (run.InnerText != "")
            {
                string[] escapeCharacters = new string[2];

                foreach (var text in run)
                {

                    if (text is Text)
                    {
                        escapeCharacters = ContainsEscape(text.InnerText);
                        if (isBlockQuote(block?.ParagraphProperties))
                        {
                            constructorBase += "\n";
                            constructorBase += ">" + text.InnerText;
                            constructorBase += "\n";
                            continue;
                        }
                        else
                        {

                            if (escapeCharacters[0] is not "")
                            {
                                constructorBase += "" + text.InnerText.Replace(escapeCharacters[0], escapeCharacters[1]);
                                isEsc = true;
                            }
                            else
                            {
                                constructorBase += text.InnerText;
                            }
                            if (text.InnerText == "/")
                            {
                                continue;
                            }

                        }
                    }

                    if (text is Break) { constructorBase += "\n"; continue; }
                    //checkbox
                    if (text.InnerText == "☐") { constructorBase = " [ ]"; continue; }
                    if (text.InnerText == "☒") { constructorBase = " [X]"; continue; }


                    /// 
                    /// Hyperlink 
                    /// 
                    if (links.Count() > 0 && links.Count() > linksCount)
                    {
                        var LId = links.ElementAt(linksCount).Id;
                        var result = buildHyperLink(text, LId, isEsc);
                        //is hyperlink
                        if (result != "")
                        {
                            constructorBase += result;
                            linksCount++;
                            break; //this break prevents  duplication of hyperlink description
                        }

                    }

                    //code block
                    if (isCodeBlock(block?.ParagraphProperties))
                    {
                        constructorBase = "~~~~\n" + constructorBase + "\n~~~~\n";
                        continue;
                    }

                }
            }

            //Images
            if (run.Descendants<Drawing>().Count() > 0)
            {
                string[] urlInfo = findPicUrl(run);
                constructorBase = "![" + urlInfo[0] + "](" + urlInfo[1] + ")";
            }

            // fonts, size letter, links
            if (run.RunProperties != null)
            {
                prefix = ProcessRunElements(run);
                constructorBase = prefix + constructorBase + prefix;
            }

            //general style, lists, aligment, spacing
            if (block.ParagraphProperties != null)
            {
                prefix = ProcessParagraphElements(block);

                if (prefix == null) prefix = "";

                if (prefix.Contains("#") || prefix.Contains("-"))
                {
                    constructorBase = prefix + " " + constructorBase;
                }

                if (prefix.Contains(">"))
                {
                    constructorBase = ProcessBlockQuote(run);
                }

            }


            //linksCount = 0;
            textBuilder.Append(constructorBase);
            constructorBase = "";
        }
        linksCount = 0;
        constructorBase = textBuilder.ToString();
        textBuilder.Clear();




        textBuilder.Append(constructorBase);
        //if code block

        textBuilder.Append("\n");
    }

    private static string[] ContainsEscape(string InnerText)
    {
        string[] result = new string[2];
        if (InnerText.Contains("#"))
        {
            result[0] = "#";
            result[1] = "\\#";
            return result;
        }
        else if (InnerText.Contains("-"))
        {
            result[0] = "#";
            result[1] = "\\#";
            return result;
        }
        else if (InnerText.Contains(">"))
        {
            result[0] = ">";
            result[1] = "\\>";
            return result;
        }
        else if (InnerText.Contains("["))
        {
            result[0] = "[";
            result[1] = "\\[";
            return result;
        }
        else if (InnerText.Contains("!["))
        {
            result[0] = "![";
            result[1] = "\\!\\[";
            return result;
        }
        else if (InnerText.Contains("*"))
        {
            result[0] = "![";
            result[1] = "\\!\\[";
            return result;
        }
        else
        {
            result[0] = "";
            result[1] = "";
            return result;
        }

    }

    private static string buildHyperLink(OpenXmlElement text, string id = "", bool isEsc = false) //STRING LITERAL OR OPTIONAL
    {
        string cbt = "";
        if (text is RunProperties)
        {   //get to runStyles

            //var asd = text.Descendants<RunStyle>();
            foreach (RunStyle runStyle in text.Descendants<RunStyle>())
            {
                //RunStyle runStyle = (RunStyle)text.FirstChild;
                if (runStyle.Val == "Hyperlink")
                {
                    if (isEsc)
                    {
                        cbt = hyperlinks.First(leenk => leenk.Id == id).Uri + "";

                    }
                    else
                    {
                        cbt = "[" + text.Parent.InnerText + "](" + hyperlinks.First(leenk => leenk.Id == id).Uri + ")";


                    }
                    return cbt;
                }

            }


        }
        return "";

    }

    private static bool isBlockQuote(ParagraphProperties? Properties)
    {
        if (Properties == null) return false;
        // have 4 borderlines
        bool isLines = false;
        //shade
        bool isShading = false;
        //  indentation
        bool isIndentation = false;
        foreach (var style in Properties)
        {
            if (style is Shading) isShading = true;

            if (style is ParagraphBorders) isLines = true;

            if (style is Indentation) isIndentation = true;
        }
        return (isLines && isShading == false && isIndentation);
    }

    private static bool isCodeBlock(ParagraphProperties Properties)
    {
        if (Properties == null) return false;
        // have 4 borderlines
        bool isLines = false;
        //shade
        bool isShading = false;
        //  indentation
        bool isIndentation = false;

        foreach (var style in Properties)
        {
            if (style is Shading) isShading = true;

            if (style is ParagraphBorders) isLines = true;

            if (style is Indentation) isIndentation = true;
        }

        return (isLines && isShading && isIndentation);
    }


    private static string[] findPicUrl(Run run)
    {
        string[] url = new string[2];
        foreach (
            DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties picAtr
            in
            run.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>())
        {
            url[1] = picAtr.Name;//url
            url[0] = picAtr.Description;//description
        }
        return url;
    }

    private static void headerDivider(List<String> align, StringBuilder textBuilder)
    {
        textBuilder.Append("|");
        foreach (var column in align)
        {
            switch (column)
            {
                case "left":
                    textBuilder.Append(":---|");
                    break;

                case "center":
                    textBuilder.Append(":---:|");
                    break;

                case "right":
                    textBuilder.Append("---:|");
                    break;

                case "normal":
                    textBuilder.Append("---|");
                    break;
            }
        }
        textBuilder.AppendLine("");

    }


}



