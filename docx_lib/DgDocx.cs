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
//using DocumentFormat.OpenXml.Drawing;

public class DgDocx
{
    private static IEnumerable<HyperlinkRelationship> hyperlinks;

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
    public async static Task md_to_docx(String md, Stream inputStream, bool debug = false) //String mdFile, String docxFile, String template)
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

        var html = Markdown.ToHtml(md, pipeline);
        //edit on debug the h
        //All the document is being saved in the stream
        using (WordprocessingDocument doc = WordprocessingDocument.Create(inputStream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
           

            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.ParseHtml(html);
            mainPart.Document.Save();
        }
    }

    public async static Task docx_to_md(Stream infile, Stream outfile, String name)
    {
        WordprocessingDocument wordDoc = WordprocessingDocument.Open(infile, false);
        DocumentFormat.OpenXml.Wordprocessing.Body body
        = wordDoc.MainDocumentPart.Document.Body;


        StringBuilder textBuilder = new StringBuilder();
        var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        hyperlinks = wordDoc.MainDocumentPart.HyperlinkRelationships;
        if (parts != null)
        {
            
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
        using (var streamWriter = new StreamWriter(name + ".md"))
        {
            String s = textBuilder.ToString();
            streamWriter.Write(s);
        }

        //commented code is for .zip files

        //using (var archive = new ZipArchive(outfile, ZipArchiveMode.Create, true))
        //{
        //    var demoFile = archive.CreateEntry(name);
        //    using (var entryStream = demoFile.Open())
        //    {
        //        using (var streamWriter = new StreamWriter(entryStream))
        //        {
        //            String s = textBuilder.ToString();
        //            streamWriter.Write(s);
        //        }
        //    }
        //}
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

        //iterate along every element in the Paragraphs and childrens
        foreach (var run in block.Descendants<Run>())
        {
            String prefix = "";
            
            if (run.InnerText != "")
            {
               
                foreach (var text in run)
                {

                    if (text is Text) {
                        if (isBlockQuote(block?.ParagraphProperties)) {constructorBase += ">" + text.InnerText;continue;}
                        else constructorBase += text.InnerText;
                    }

                    if (text is Break) {constructorBase += "\n"; continue;}
                    //checkbox
                    if (text.InnerText == "☐") { constructorBase = " [ ]"; continue;}
                    if (text.InnerText == "☒") { constructorBase = " [X]"; continue;}

                    
                    //Hyperlink
                    var links = block.Descendants<Hyperlink>();
                    if (links.Count() > 0)
                    {
                        var LId = links.First().Id;
                        var result = buildHyperLink(text,LId);
                        //is hyperlink
                        if (result != "" )
                        {
                            constructorBase += result.Replace("[text.Parent.InnerText]","["+run.InnerText+"}");
                            continue;
                        }

                    }


                    //code block
                    if (isCodeBlock(block?.ParagraphProperties))
                    {
                        constructorBase = "~~~~\n" + constructorBase + "\n~~~~\n";
                        continue;
                    }
                    
                    constructorBase += "\n";
                }
            }

            //Images
            if (run.FirstChild is Drawing)
            {
                constructorBase = "![](" + findPicUrl(run) + ")";
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



            textBuilder.Append(constructorBase);
            constructorBase = "";
        }
        //if code block
        constructorBase = textBuilder.ToString();
        textBuilder.Clear();


    

        textBuilder.Append(constructorBase);
        //textBuilder.Append("\n\n");
        textBuilder.Append("\n");
    }

    private static string buildHyperLink(OpenXmlElement text,string id)
    {
        string cbt = "";
        if (text is RunProperties)
        {   //get to runStyles
            if (text.FirstChild is RunStyle)
            {
                RunStyle runStyle = (RunStyle)text.FirstChild;
                if (runStyle.Val == "Hyperlink")
                {

                    cbt = "[" + "text.Parent.InnerText" + "](" + hyperlinks.First(leenk => leenk.Id == id).Uri + ")";
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
        return (isLines && isShading==false && isIndentation);
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


    private static string findPicUrl(Run run)
    {

        string ImageUrl = "";
        if (null == run?.FirstChild?.FirstChild) return "";

        foreach (var graphic in run?.FirstChild?.FirstChild)
        {
            if (graphic is DocumentFormat.OpenXml.Drawing.Graphic)
            {
                foreach (var pic in graphic?.FirstChild?.FirstChild)
                {
                    if (pic is DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties)
                    {
                        foreach (var nvdp in pic)
                        {
                            if (nvdp is DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties)
                            {
                                //listar los atributos
                                DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties picAtr = (DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties)nvdp;

                                ImageUrl = picAtr.Name;
                            }
                            break;

                        }
                        break;
                    }
                }
                break;
            }
        }

        if (ImageUrl == null) return " ";
        return ImageUrl;
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



