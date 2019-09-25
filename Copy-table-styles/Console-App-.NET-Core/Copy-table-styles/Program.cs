using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Copy_table_styles
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
                //Opens the source Word document containing table style definition
                WordDocument srcDocument = new WordDocument();
                docStream = File.OpenRead(Path.GetFullPath(@"../../../TableStyles.docx"));
                srcDocument.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Gets the table style (CustomStyle) from the styles collection
                WTableStyle srcTableStyle = srcDocument.Styles.FindByName("CustomStyle", StyleType.TableStyle) as WTableStyle;
                //Creates a cloned copy of table style
                WTableStyle clonedTableStyle = srcTableStyle.Clone() as WTableStyle;
                //Adds the cloned copy of source table style to the destination document 
                document.Styles.Add(clonedTableStyle);
                //Releases the resources of source Word document instance
                srcDocument.Dispose();
                //Gets table to apply style
                WTable table = (WTable)document.LastSection.Tables[0];
                //Applies the custom table style to the table
                table.ApplyStyle("CustomStyle");
                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
