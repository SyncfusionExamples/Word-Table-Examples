using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Modify_existing_table_style
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
                //Gets the table style (Medium Shading 1 Accent 1) from the styles collection
                WTableStyle tableStyle = document.Styles.FindByName("Medium Shading 1 Accent 1", StyleType.TableStyle) as WTableStyle;
                //Gets the conditional formatting style of the first row (table headers) from the table style
                ConditionalFormattingStyle firstRowStyle = tableStyle.ConditionalFormattingStyles[ConditionalFormattingType.FirstRow];
                if (firstRowStyle != null)
                {
                    //Modifies the background color for first row (table headers)
                    firstRowStyle.CellProperties.BackColor = Syncfusion.Drawing.Color.FromArgb(255, 31, 56, 100);
                }
                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
    }
}
