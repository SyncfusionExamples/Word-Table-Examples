using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Apply_built_in_table_style
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Adds a section to the Word document
                IWSection section = document.AddSection();
                //Sets the page margin
                section.PageSetup.Margins.All = 72;
                //Adds a paragrah to the section
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                paragraph.ParagraphFormat.AfterSpacing = 20;
                IWTextRange textRange = paragraph.AppendText("Suppliers");
                textRange.CharacterFormat.FontSize = 14;
                textRange.CharacterFormat.Bold = true;
                textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(255, 50, 62, 79);
                //Modifies the font size as 10 for default paragraph style
                WParagraphStyle style = document.Styles.FindByName("Normal") as WParagraphStyle;
                style.CharacterFormat.FontSize = 10;
                //Adds a table to the section
                WTable table = section.AddTable() as WTable;
                table.ResetCells(1, 6);
                table[0, 0].Width = 52f;
                table[0, 0].AddParagraph().AppendText("Supplier ID");
                table[0, 1].Width = 128f;
                table[0, 1].AddParagraph().AppendText("Company Name");
                table[0, 2].Width = 70f;
                table[0, 2].AddParagraph().AppendText("Contact Name");
                table[0, 3].Width = 92f;
                table[0, 3].AddParagraph().AppendText("Address");
                table[0, 4].Width = 66.5f;
                table[0, 4].AddParagraph().AppendText("City");
                table[0, 5].Width = 56f;
                table[0, 5].AddParagraph().AppendText("Country");
                //Imports data to the table.
                ImportDataToTable(table);
                //Applies the built-in table style (Medium Shading 1 Accent 1) to the table
                table.ApplyStyle(BuiltinTableStyle.MediumShading1Accent1);
                //Saves the file in the given path
                Stream docStream = File.Create(Path.GetFullPath(@"Result.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }

        /// <summary>
        /// Imports the data from XML file to the table.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void ImportDataToTable(WTable table)
        {
            FileStream fs = new FileStream(@"../../../Suppliers.xml", FileMode.Open, FileAccess.Read);
            XmlReader reader = XmlReader.Create(fs);
            if (reader == null)
                throw new Exception("reader");
            while (reader.NodeType != XmlNodeType.Element)
                reader.Read();
            if (reader.LocalName != "SuppliersList")
                throw new XmlException("Unexpected xml tag " + reader.LocalName);
            reader.Read();
            while (reader.NodeType == XmlNodeType.Whitespace)
                reader.Read();
            while (reader.LocalName != "SuppliersList")
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "Suppliers":
                            //Adds new row to the table for importing data from next record.
                            WTableRow tableRow = table.AddRow(true);
                            ImportDataToRow(reader, tableRow);
                            break;
                    }
                }
                else
                {
                    reader.Read();
                    if ((reader.LocalName == "SuppliersList") && reader.NodeType == XmlNodeType.EndElement)
                        break;
                }
            }
            reader.Dispose();
            fs.Dispose();
        }
        /// <summary>
        /// Imports the data to the table row.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void ImportDataToRow(XmlReader reader, WTableRow tableRow)
        {
            if (reader == null)
                throw new Exception("reader");
            while (reader.NodeType != XmlNodeType.Element)
                reader.Read();
            if (reader.LocalName != "Suppliers")
                throw new XmlException("Unexpected xml tag " + reader.LocalName);
            reader.Read();
            while (reader.NodeType == XmlNodeType.Whitespace)
                reader.Read();
            while (reader.LocalName != "Suppliers")
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "SupplierID":
                            tableRow.Cells[0].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "CompanyName":
                            tableRow.Cells[1].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "ContactName":
                            tableRow.Cells[2].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "Address":
                            tableRow.Cells[3].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "City":
                            tableRow.Cells[4].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        case "Country":
                            tableRow.Cells[5].AddParagraph().AppendText(reader.ReadElementContentAsString());
                            break;
                        default:
                            reader.Skip();
                            break;
                    }
                }
                else
                {
                    reader.Read();
                    if ((reader.LocalName == "Suppliers") && reader.NodeType == XmlNodeType.EndElement)
                        break;
                }
            }
        }
    }
}
