﻿namespace docx_lib;
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

public class DgDocx
{

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
        var html = Markdown.ToHtml(md);

        //All the document is being saved in the stream
        using (WordprocessingDocument doc = WordprocessingDocument.Create(inputStream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            if (debug) {
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Markdown: " + md));
                }

                {
                    Paragraph para2 = body.AppendChild(new Paragraph());
                    Run run2 = para2.AppendChild(new Run());
                    run2.AppendChild(new Text("Html: " + html));
                }
            }

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
        if (parts != null)
        {
            foreach (var block in parts.ChildElements)
            {
                if (block is Paragraph)
                {
                    //This method is for manipulating the style of Paragraphs and text inside
                    ProcessParagraph((Paragraph)block, textBuilder);
                }

                if (block is Table)
                {
                    ProcessTable((Table)block, textBuilder);
                }
            }
        }

        //This code is replacing the above one because I need to check the .md file faster
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
        foreach (var row in node.Descendants<TableRow>())
        {
            textBuilder.Append("| ");
            foreach (var cell in row.Descendants<TableCell>())
            {
                foreach (var para in cell.Descendants<Paragraph>())
                {
                    ProcessParagraph(para, textBuilder);
                }
                textBuilder.Append(" | ");
            }
            textBuilder.AppendLine("");
        }
    }

    private static String ProcessParagraphElements(Paragraph block)
    {
        String style = block.ParagraphProperties.ParagraphStyleId.Val;
        int num;
        String prefix = "";

        //to find which block have Heading Paragraphs
        if (style.Contains("Heading"))
        {
            num = int.Parse(style.Substring(style.Length - 1));
            

            for(int i = 0; i<num; i++)
            {
                prefix += "#";
            }
            return prefix;
        }

        //to find which block have List Paragraphs
        if (style == "ListParagraph")
        {
            return prefix = "-";
        }

        return null;
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


    private static void ProcessParagraph(Paragraph block, StringBuilder textBuilder)
    {
        String constructorBase = "";

        //iterate along every element in the Paragraphs and childrens
        foreach (var run in block.Descendants<Run>())
        {
            String prefix = "";

            if (run.InnerText != "")
            {
                constructorBase = run.InnerText;
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

                if (prefix.Contains("#") || prefix.Contains("-"))
                {
                    constructorBase = prefix + " " + constructorBase;
                }
            }
            textBuilder.Append(constructorBase);
            constructorBase = "";
        }
        
        textBuilder.Append("\n");
    }




}