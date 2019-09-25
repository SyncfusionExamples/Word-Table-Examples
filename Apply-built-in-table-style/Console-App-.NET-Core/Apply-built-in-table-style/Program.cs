using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Apply_built_in_table_style
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the input Word document
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Table.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Gets table to apply style
                WTable table = (WTable)document.LastSection.Tables[0];
                //Applies the built-in table style (Medium Shading 1 Accent 1) to the table
                table.ApplyStyle(BuiltinTableStyle.MediumShading1Accent1);
                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
