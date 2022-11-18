

```
mkdir cli2
cd cli2
code .
dotnet new sln

dotnet new classlib -o docx_lib
cd docx_lib
dotnet add package Markdig --version 0.30.4
dotnet add package NS.HtmlToOpenXml --version 1.1.0
cd ..

dotnet new console -o docx_md
dotnet new blazorwasm -o docx_web

dotnet sln add docx_md/docx_md.csproj
dotnet sln add docx_web/docx_web.csproj 
dotnet sln add docx_lib/docx_lib.csproj 

dotnet add docx_md/docx_md.csproj reference docx_lib/docx_lib.csproj

dotnet add docx_web/docx_web.csproj reference docx_lib/docx_lib.csproj

dotnet build
dotnet publish -c release
cd bin/Release/net6.0/publish/wwwroot
surge . mddocx2.surge.sh

```

# Getting Started

This section only wants to cover the basics to translate from markdown files to .docx format.

## Read Markdown Files

To translate markdown files to .docx format we need to read the whole md file using the `File` class followed by the method `ReadAllText`, it'll allow you to read the text inside a file and store it in a variable. After that we only have to instantiate a `MermoryStream` object that will help us to store the upcoming xml files and then we just need to call the method `md_to_docx` from the class `DgDocx` passing the markdown text and the Stream variable that will keep the conversion, please check the code below.

```
                   // markdown to docx
                    var md =  File.ReadAllText(mdFile);
                    var inputStream = new MemoryStream();
                    await DgDocx.md_to_docx(md, inputStream);
```

Now we have the .docx into the MemoryStream variable, it's just to write it into a document already located into a directory calling as before, the `File` class and the method `WriteAllBytes`, it'll write byte by byte the document into the .docx file we'll pass as a path. Check the code below.

```
                    File.WriteAllBytes(docxFile, inputStream.ToArray());
```

## Read .docx files

To translate files from .docx format to markdown we only need to have the path of the .docx file we need to translate.
With the `using` operator (recommended to close the stream automatically) we can open a .docx file and store it into a variable, then we have to call the `DgDocx` class and the method `docx_to_md`, passing the stream where is written the .docx file and the markdown path where we want to have the file located. Please check the code below.

```
                    using (var instream = File.Open(docxFile, FileMode.Open)){
                        var outstream = new MemoryStream();
                        await DgDocx.docx_to_md(instream, outstream, root)
```
