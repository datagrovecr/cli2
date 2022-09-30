﻿using docx_lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

internal class Program
{
    static async Task Main(string[] args)
    {
        var outdir = "/Users/fabianvalverde/Documents/GitHub/cli2/docx_md/test_results/";
        string[] files = Directory.GetFiles("/Users/fabianvalverde/Documents/GitHub/cli2/docx_md/folder_Tests", "*.md", SearchOption.TopDirectoryOnly);

        foreach (var mdFile in files)
        {
            //Just getting the end route
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string root = outdir + fn.Replace("_md", "");
            var docxFile = root + ".docx";
            {
                try
                {
                   // markdown to docx
                    var md =  File.ReadAllText(mdFile);
                    var inputStream = new MemoryStream();
                    await DgDocx.md_to_docx(md, inputStream);
                    File.WriteAllBytes(docxFile, inputStream.ToArray());                       


                    // convert the docx back to markdown.
                    using (var instream = File.Open(docxFile, FileMode.Open)){
                        var outstream = new MemoryStream();
                        await DgDocx.docx_to_md(instream, outstream, fn.Replace("_md", ".md"));
                        using (var fileStream = new FileStream(root+".zip", FileMode.Create))
                        {
                            outstream.Seek(0, SeekOrigin.Begin);
                            outstream.CopyTo(fileStream);
                        }                        
                    }
                }catch (Exception e)
                {
                    Console.WriteLine($"{mdFile} failed {e}");
                }
                
            }
          
        }
    }

    static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
    {

        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        var errors = validator.Validate(wpDoc);

        if (!errors.GetEnumerator().MoveNext())
            return;

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

        Console.ForegroundColor = ConsoleColor.Gray;
        foreach (ValidationErrorInfo error in errors)
        {
            Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
            Console.WriteLine();
        }

        Console.ReadLine();
    }
}