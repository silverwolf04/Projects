using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Data;
using System.Diagnostics;

namespace PdfCreator
{
    public enum PageSide
    {
        LeftSide,
        RightSide
    }
    public class PdfRender
    {
        public PdfRender ()
        {
            Console.WriteLine("No arguments passed.");
        }
        public PdfRender(string[] args)
        {
            string type = "Unknown";
            if (args.Length > 0)
                type = args[0].ToString();
            SetPdfType(type);
        }
        public enum PdfTypes
        {
            Test,
            FacultyStaff,
            Department
        }

        public PdfTypes PdfType;
        private bool Viewer = false;

        private void SetPdfType(string argStr)
        {
            // if an enum is not properly parsed, the first value will be set as the enum (Unknown)
            // Enum.TryParse<PdfTypes>(argStr, true, out PdfType);
            if(!Enum.TryParse(argStr, true, out PdfType) || !Enum.IsDefined(typeof(PdfTypes), PdfType))
                PdfType = PdfTypes.Test;
            if (PdfType == PdfTypes.Test)
                Viewer = true;
        }
        public void GenerateTest(ref Document document, string output)
        {
            document = CreateTestDocument(output);
        }

        static Document CreateTestDocument(string output)
        {
            // Create a new MigraDoc document
            Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);

            // Add some text to the paragraph
            paragraph.AddFormattedText(output, TextFormat.Bold);
            /*
            Hyperlink hyperlink = paragraph.AddHyperlink("https://mines.edu", HyperlinkType.Url);
            hyperlink.AddText("This is a test");
            */

            return document;
        }

        static Table AddCustomTable(Section section)
        {
            Table table = section.AddTable();
            /*
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            */
            table.AddColumn("8.75cm");
            table.AddColumn("8.75cm");
            table.Borders.Visible = false;
            //table.Borders.Visible = true;
            table.Borders.Width = 1;
            return table;
        }
        private static void GenerateRow(PageSide pageSide, Employee employee, ref Row row)
        {
            int cellInt = 0;
            string concatenatedStr = string.Empty;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.RightSide:
                    //cellInt = 2;
                    cellInt = 1;
                    break;
                default:
                    Console.WriteLine("Invalid Page Side provided: " + pageSide.ToString());
                    break;

            }

            /*
            if(!string.IsNullOrWhiteSpace(employee.PhoneNumber))
                row.Cells[cellInt].AddParagraph(employee.PhoneNumber);
            //cellInt++;
            */
            Paragraph paragraph;
            FormattedText text;

            if (!string.IsNullOrWhiteSpace(employee.Name))
            {
                //row.Cells[cellInt].AddParagraph(employee.Name);
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddFormattedText(employee.Name, TextFormat.Bold);
            }
            if(!string.IsNullOrWhiteSpace(employee.Title))
                row.Cells[cellInt].AddParagraph("Title: " + employee.Title);
            if(!string.IsNullOrWhiteSpace(employee.Department))
                row.Cells[cellInt].AddParagraph("Dept: " + employee.Department);
            if (!string.IsNullOrWhiteSpace(employee.Building))
                concatenatedStr = "Bldg: " + employee.Building;
            if (!string.IsNullOrWhiteSpace(employee.Building) && !string.IsNullOrWhiteSpace(employee.Office))
                concatenatedStr += " Office/Room: " + employee.Office;
            if (!string.IsNullOrEmpty(concatenatedStr))
                row.Cells[cellInt].AddParagraph(concatenatedStr);
            if (!string.IsNullOrWhiteSpace(employee.Url))

            {
                paragraph = row.Cells[cellInt].AddParagraph();
                //paragraph.Format.Font.Color = Color.FromRgb(51, 153, 255);
                paragraph.AddText("Home Page: ");
                Hyperlink hyperlink = paragraph.AddHyperlink(employee.Url, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                //text.AddText("Home Page: ");
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText("Click Here", TextFormat.Underline);
                //hyperlink.AddFormattedText("Home Page: Click Here", TextFormat.Underline);
                //row.Cells[cellInt].AddParagraph("Home Page: " + employee.Url);
            }
            if (!string.IsNullOrWhiteSpace(employee.EmailAddress))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddText("Email Address: ");
                Hyperlink hyperlink = paragraph.AddHyperlink("mailto:" + employee.EmailAddress, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                //text.AddText("Home Page: ");
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText(employee.EmailAddress, TextFormat.Underline);

                //row.Cells[cellInt].AddParagraph("Email: " + employee.EmailAddress);
            }
            if (!string.IsNullOrEmpty(employee.PhoneNumber))
                row.Cells[cellInt].AddParagraph("Phone: " + employee.PhoneNumber);
            if (!string.IsNullOrWhiteSpace(employee.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + employee.FaxNumber);
            row.Cells[cellInt].AddParagraph("");
        }

        static void AddCustomHeader(string text, ref Section section)
        {
            HeaderFooter header = new HeaderFooter();
            header.AddParagraph(text);
            header.Format.Font.Size = 30;
            header.Format.Font.Name = "Times-Roman";
            header.Format.Alignment = ParagraphAlignment.Center;
            section.Headers.Primary = header.Clone();
        }

        static void AddCustomFooter(string text, ref Section section)
        {
            HeaderFooter footer = new HeaderFooter();
            footer.AddParagraph(text);
            footer.Format.Font.Size = 10;
            footer.Format.Font.Name = "Times-Roman";
            footer.Format.Alignment = ParagraphAlignment.Center;
            section.Footers.Primary = footer.Clone();
        }

        private void CreateDocumentDepartment(ref Document document)
        {
            Section section = document.AddSection();
            //Document document = new Document();
            //return document;
        }

        private void CreateDocumentFacultyStaff(ref Document document)
        {
            // Create a new MigraDoc document
            //Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            paragraph.Format.Font.Color = Colors.Black;
            paragraph.Format.Font.Name = "Times-Roman";
            paragraph.Format.Font.Size = 32;

            // Add some text to the paragraph
            paragraph.AddText(Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine);
            paragraph.AddFormattedText("Colorado School of Mines" + Environment.NewLine + "Faculty/Staff Directory", TextFormat.Bold);

            // Add a paragraph to the section
            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            paragraph.Format.Font.Color = Colors.Black;
            paragraph.Format.Font.Name = "Times-Roman";
            paragraph.Format.Font.Size = 16;
            paragraph.AddText(Environment.NewLine + Environment.NewLine + "This document was generated at " + DateTime.Now.ToString("M/d/yyyy h:mm:ss tt"));

            section.AddPageBreak();
            section = document.AddSection();

            AddCustomHeader("A", ref section);

            string footerStr = "Need additional help?" + Environment.NewLine + "Go to " + "https://helpdesk.mines.edu";
            AddCustomFooter(footerStr, ref section);
            /*
            HeaderFooter footer = new HeaderFooter();
            footer.AddParagraph(footerStr);
            footer.Format.Font.Size = 10;
            footer.Format.Font.Name = "Times-Roman";
            footer.Format.Alignment = ParagraphAlignment.Center;
            section.Footers.Primary = footer.Clone();
            */

            section.AddPageBreak();

            // initializer for right side count
            int rowInt = 0;

            // add & set the table margins on the new page
            Table table = AddCustomTable(section);

            Employee employee = new Employee();

            DirectoryTasks directoryTasks = new DirectoryTasks();
            directoryTasks.GetData(out DataTable dataTable);
            Console.WriteLine("Rows: {0}", dataTable.Rows.Count);

            Row row = new Row();
            int currPageRow = 1;
            string currHeaderStr = string.Empty;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                PageSide pageSide;
                employee.Name = dataRow[Properties.FacStaffPdf.Default.Name].ToString();
                employee.PhoneNumber = dataRow[Properties.FacStaffPdf.Default.PhoneNumber].ToString();
                employee.Department = dataRow[Properties.FacStaffPdf.Default.Department].ToString();
                employee.FaxNumber = dataRow[Properties.FacStaffPdf.Default.FaxNumber].ToString();
                employee.EmailAddress = dataRow[Properties.FacStaffPdf.Default.EmailAddress].ToString();
                employee.Title = dataRow[Properties.FacStaffPdf.Default.Title].ToString();
                employee.Building = dataRow[Properties.FacStaffPdf.Default.Building].ToString();
                employee.Office = dataRow[Properties.FacStaffPdf.Default.Office].ToString();
                employee.Url = dataRow[Properties.FacStaffPdf.Default.URL].ToString();

                if(string.IsNullOrEmpty(currHeaderStr))
                    currHeaderStr = employee.Name.Substring(0, 1);

                if (!employee.Name.StartsWith(currHeaderStr) || currPageRow > 16)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    //section.AddPageBreak();
                    section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;
                    currHeaderStr = employee.Name.Substring(0, 1);

                    // update page header
                    AddCustomHeader(currHeaderStr, ref section);

                    // add & set the formatted on the new page
                    table = AddCustomTable(section);
                }    

                if (currPageRow == 9)
                    rowInt = 0;

                if (currPageRow <= 8)
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else // rows 9-16
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRow(pageSide, employee, ref row);
                employee.ClearAll();
                Console.WriteLine("currPageRow: " + currPageRow);
                Console.WriteLine("rowInt:" + rowInt);
                currPageRow++;
                rowInt++;
            }

            //return document;
        }

        public void RequestPdf(string arg)
        {
            string filename = string.Empty;
            // Create a MigraDoc document
            Document document = new Document();

            switch(PdfType)
            {
                case PdfTypes.FacultyStaff:
                    Console.WriteLine("This will render the Mines Faculty/Staff PDF.");
                    //document = CreateDocumentFacultyStaff():
                    CreateDocumentFacultyStaff(ref document);
                    filename = "fac_staff_dir.pdf";
                    break;
                case PdfTypes.Department:
                    Console.WriteLine("This will render the Mines Department Directory PDF");
                    CreateDocumentDepartment(ref document);
                    filename = "departmental_dir.pdf";
                    break;
                //case PdfTypes.Test:
                default:
                    Console.WriteLine("This will render the Test PDF");
                    GenerateTest(ref document, arg);
                    filename = "HelloWorld.pdf";
                    break;
            }

            //document.UseCmykColor = true;
            // A flag indicating whether to create a Unicode PDF or a WinAnsi PDF file.
            // This setting applies to all fonts used in the PDF document.
            // This setting has no effect on the RTF renderer.
            const bool unicode = false;

            // Create a renderer for the MigraDoc document.
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode)
            {
                // Associate the MigraDoc document with a renderer
                Document = document
            };

            // Layout and render document to PDF
            pdfRenderer.RenderDocument();

            // Save the document...
            pdfRenderer.PdfDocument.Save(filename);

            // ...and start a viewer if testing
            if(Viewer)
                Process.Start(filename);
        }
    }
}
