using docx_lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Compression;

internal class Program
{
    static async Task Main(string[] args)
    {

        var outdir = @"./../../../../docx_md/test_results/";
        string[] files = Directory.GetFiles(@"../docx_md/folder_tests/", "*.md", SearchOption.TopDirectoryOnly);



        foreach (var mdFile in files)
        {
            //Just getting the end route
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string root = outdir + fn.Replace("_md", "");
            var docxFile = root + ".docx";
            try
            {
                // markdown to docx
                var md = File.ReadAllText(mdFile);
                var inputStream = new MemoryStream();
                await DgDocx.md_to_docx(md, inputStream);

                //inputStream is writing into the .docx file
                File.WriteAllBytes(docxFile, inputStream.ToArray());


                // convert the docx back to markdown.
                using (var instream = File.Open(docxFile, FileMode.Open))
                {
                    var outstream = new MemoryStream();
                    await DgDocx.docx_to_md(instream, outstream, root);//Previous: instream, outstream, fn.Replace("_md", "")

                    //The commented code is for .zip files

                    //using (var fileStream = new FileStream(root+".md", FileMode.Create))
                    //{
                    //    outstream.Seek(0, SeekOrigin.Begin);
                    //    outstream.CopyTo(fileStream);
                    //}                        
                }
                using (ZipArchive archive = ZipFile.OpenRead(outdir + "test.docx"))
                {
                    archive.ExtractToDirectory(outdir + "test.unzipped", true);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"{mdFile} failed {e}");
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