//using PdfSharp.Drawing;
//using PdfSharp.Drawing.Layout;
//using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Diagnostics;
//using Microsoft.VisualBasic;

namespace CLI_Test
{
    public enum PageSide
    {
        LeftSide,
        RightSide
    }
    class PdfRender
    {
        public static void GenerateRow(PageSide pageSide, Employee employee, ref Row row)
        {
            int cellInt = -1;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.RightSide:
                    cellInt = 2;
                    break;
                default:
                    Console.WriteLine("Invalid Page Side provided: " + pageSide.ToString());
                    break;

            }

            if(!string.IsNullOrWhiteSpace(employee.PhoneNumber))
                row.Cells[cellInt].AddParagraph(employee.PhoneNumber);
            cellInt++;
            if (!string.IsNullOrWhiteSpace(employee.Name))
                row.Cells[cellInt].AddParagraph(employee.Name);
            if(!string.IsNullOrWhiteSpace(employee.Title))
                row.Cells[cellInt].AddParagraph("Title: " + employee.Title);
            if(!string.IsNullOrWhiteSpace(employee.Department))
                row.Cells[cellInt].AddParagraph("Dept: " + employee.Department);
            if(!string.IsNullOrWhiteSpace(employee.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + employee.FaxNumber);
            if(!string.IsNullOrWhiteSpace(employee.EmailAddress))
                row.Cells[cellInt].AddParagraph("Email: " + employee.EmailAddress);
            row.Cells[cellInt].AddParagraph("");
        }
        public void GenerateTest()
        {
            GenerateTest("Hello World!");
        }
        public void GenerateTest(string output)
        {
            Console.WriteLine(output);

            // Create a MigraDoc document
            Document document = CreateTestDocument(output);
            document.UseCmykColor = true;

            // ===== Unicode encoding and font program embedding in MigraDoc is demonstrated here =====

            // A flag indicating whether to create a Unicode PDF or a WinAnsi PDF file.
            // This setting applies to all fonts used in the PDF document.
            // This setting has no effect on the RTF renderer.
            const bool unicode = false;

            // An enum indicating whether to embed fonts or not.
            // This setting applies to all font programs used in the document.
            // This setting has no effect on the RTF renderer.
            // (The term 'font program' is used by Adobe for a file containing a font. Technically a 'font file'
            // is a collection of small programs and each program renders the glyph of a character when executed.
            // Using a font in PDFsharp may lead to the embedding of one or more font programms, because each outline
            // (regular, bold, italic, bold+italic, ...) has its own fontprogram)
            //const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

            // ========================================================================================

            // Create a renderer for the MigraDoc document.
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);

            // Associate the MigraDoc document with a renderer
            pdfRenderer.Document = document;

            // Layout and render document to PDF
            pdfRenderer.RenderDocument();

            // Save the document...
            const string filename = "HelloWorld.pdf";
            pdfRenderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);

            /* // Use this code with PDFsharp NuGet
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Verdana", 20, XFontStyle.BoldItalic);

            // Draw the text
            gfx.DrawString(output, font, XBrushes.Black,
              new XRect(0, 0, page.Width, page.Height),
              XStringFormats.Center);

            // Save the document...
            const string filename = "HelloWorld.pdf";
            document.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
            */
        }

        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
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

            return document;
        }

        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
        static Document CreateDocument()
        {
            // Create a new MigraDoc document
            Document document = new Document();

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

            HeaderFooter header = new HeaderFooter();
            header.AddParagraph("A");
            header.Format.Font.Size = 30;
            header.Format.Font.Name = "Times-Roman";
            header.Format.Alignment = ParagraphAlignment.Center;
            section.Headers.Primary = header.Clone();

            HeaderFooter footer = new HeaderFooter();
            footer.AddParagraph("Need additional help?" + Environment.NewLine + "Go to https://helpdesk.mines.edu");
            footer.Format.Font.Size = 10;
            footer.Format.Font.Name = "Times-Roman";
            footer.Format.Alignment = ParagraphAlignment.Center;
            section.Footers.Primary = footer.Clone();

            section.AddPageBreak();

            // initializer for right side count
            int rowInt = 0;

            //var table = section1.AddTable();
            var table = section.AddTable();
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            table.Borders.Visible = false;
            table.Borders.Width = 1;            

            Employee employee = new Employee();

            employee.Name = "Abdul, Ramin";
            employee.PhoneNumber = "(303) 754-2251";
            employee.Department = "Arithmetic";
            employee.FaxNumber = "(303) 111-3333";
            employee.EmailAddress = "rabdul@mines.edu";

            var row = table.AddRow();
            GenerateRow(PageSide.LeftSide, employee, ref row);
            employee.ClearAll();

            employee.Name = "Alan, Gary";
            employee.PhoneNumber = "(303) 123-4567";
            employee.Title = "Associate Professor";
            employee.Department = "Engineering and Design";
            employee.FaxNumber = "(303) 555-2233";
            employee.EmailAddress = "galan@mines.edu";

            row = table.AddRow();
            GenerateRow(PageSide.LeftSide, employee, ref row);
            employee.ClearAll();

            // reset the right side
            rowInt = 0;

            employee.Name = "Azura, Monica";
            employee.PhoneNumber = "(303) 999-8877";
            employee.Department = "Engineering & Design";
            employee.EmailAddress = "mazura1@mines.edu";

            //row = table.AddRow();
            row = table.Rows[rowInt];
            GenerateRow(PageSide.RightSide, employee, ref row);
            employee.ClearAll();

            employee.Name = "Axelwu, Angel";
            employee.PhoneNumber = "(303) 541-8877";
            employee.Department = "Engineering & Design";
            employee.EmailAddress = "axel@mines.edu";

            //row = table.AddRow();
            rowInt++;
            row = table.Rows[rowInt];
            GenerateRow(PageSide.RightSide, employee, ref row);
            employee.ClearAll();

            return document;
        }

        public void FacultyStaffPDF()
        {
            Console.WriteLine("This will render the Faculty/Staff PDF.");

            // Create a MigraDoc document
            Document document = CreateDocument();
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
            const string filename = "fac_staff_dir.pdf";
            pdfRenderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }
    }
}
