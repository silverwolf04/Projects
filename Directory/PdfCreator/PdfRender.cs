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
    class PdfRender
    {
        public static void GenerateRow(PageSide pageSide, Employee employee, ref Row row)
        {
            int cellInt = -1;
            string concatenatedStr = string.Empty;

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
            if (!string.IsNullOrWhiteSpace(employee.Building))
                concatenatedStr = "Bldg: " + employee.Building;
            //row.Cells[cellInt].AddParagraph("Bldg: " + employee.Building);
            if (!string.IsNullOrWhiteSpace(employee.Building) && !string.IsNullOrWhiteSpace(employee.Office))
                concatenatedStr += " Office/Room: " + employee.Office;
            //row.Cells[cellInt].AddParagraph("Office: " + employee.Office);
            if (!string.IsNullOrEmpty(concatenatedStr))
                row.Cells[cellInt].AddParagraph(concatenatedStr);
            if(!string.IsNullOrWhiteSpace(employee.EmailAddress))
                row.Cells[cellInt].AddParagraph("Email: " + employee.EmailAddress);
            if (!string.IsNullOrWhiteSpace(employee.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + employee.FaxNumber);
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
#pragma warning disable IDE0017 // Simplify object initialization
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
#pragma warning restore IDE0017 // Simplify object initialization

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

        static Table AddCustomTable(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            table.AddColumn("2.75cm");
            table.AddColumn("6cm");
            table.Borders.Visible = false;
            table.Borders.Width = 1;
            return table;
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
            /*
            employee.Name = "Abdul, Ramin";
            employee.PhoneNumber = "(303) 754-2251";
            employee.Department = "Arithmetic";
            employee.FaxNumber = "(303) 111-3333";
            employee.EmailAddress = "rabdul@mines.edu";

            row = table.AddRow();
            GenerateRow(PageSide.LeftSide, employee, ref row);
            employee.ClearAll();
            */
            /*
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
            */
            return document;
        }

        public void FacultyStaffPDF()
        {
            /* Observations dervived from predecessor PDF generator
             * fac_staff_dir.pdf
             * select * from fac_staff_dir_view1 where releaseflag is null order by lname, fname
             if ( iIDFlag <> "1" and iIDFlag <> "2" and iIDFlag <> "E" and iIDFlag <> "F" and iIDFlag <> "S" and iIDFlag <> "G" and iIDFlag <> "H" and iIDFlag <> "U" and iIDFlag <> "W" and iIDFlag <> "Y" And iIDFlag <> "?") then
			    DisplayListing iName, iDepartment, iOffice, iBldg, iWorkTel, iHomeTel, iFAX, iBusEmail, iURL, iAddressL1, iAddressL2, iCity, iDirRestrict, iTitle, iWTitle
			    numDisplayCalls = numDisplayCalls + 1
		     end if
             * 
             if ( isNull( iName ) ) then
		        iName = result1("LName") & ", " & result1("FName") & " " & result1("MName")
		        if (result1("NName") <> "") then
			        iName = iName & " (" & result1("NName") & ")"
		        end if
	         end if 
             *
            */

            Console.WriteLine("This will render the Mines Faculty/Staff PDF.");

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
            //Process.Start(filename);
        }
    }
}
