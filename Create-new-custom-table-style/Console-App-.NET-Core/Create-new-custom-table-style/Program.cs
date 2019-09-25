using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Create_new_custom_table_style
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            WordDocument document = new WordDocument();
            //Opens the input Word document
            Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Table.docx"));
            document.Open(docStream, FormatType.Docx);
            docStream.Dispose();
            //Adds a new custom table style
            WTableStyle tableStyle = document.AddTableStyle("CustomStyle") as WTableStyle;
            //Applies formatting for whole table
            tableStyle.TableProperties.RowStripe = 1;
            tableStyle.TableProperties.ColumnStripe = 1;
            tableStyle.TableProperties.Paddings.Top = 0;
            tableStyle.TableProperties.Paddings.Bottom = 0;
            tableStyle.TableProperties.Paddings.Left = 5.4f;
            tableStyle.TableProperties.Paddings.Right = 5.4f;
            //Applies conditional formatting for first row
            ConditionalFormattingStyle firstRowStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstRow);
            firstRowStyle.CharacterFormat.Bold = true;
            firstRowStyle.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(255, 255, 255, 255);
            firstRowStyle.CellProperties.BackColor = Syncfusion.Drawing.Color.Blue;
            //Applies conditional formatting for first column
            ConditionalFormattingStyle firstColumnStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstColumn);
            firstColumnStyle.CharacterFormat.Bold = true;
            //Applies conditional formatting for odd row
            ConditionalFormattingStyle oddRowBandingStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.OddRowBanding);
            oddRowBandingStyle.CellProperties.BackColor = Syncfusion.Drawing.Color.WhiteSmoke;
            //Gets table to apply style
            WTable table = (WTable)document.LastSection.Tables[0];
            //Applies the custom table style to the table
            table.ApplyStyle("CustomStyle");
            //Saves the file in the given path
            docStream = File.Create(Path.GetFullPath(@"Result.docx"));
            document.Save(docStream, FormatType.Docx);
            docStream.Dispose();
            //Releases the resources of Word document instance
            document.Dispose();
        }
    }
}
