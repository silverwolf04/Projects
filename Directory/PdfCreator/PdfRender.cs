using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Collections.Specialized;
using System.Data;
using System.Diagnostics;
using System.Text;

namespace PdfCreator
{
    public class PdfRender
    {
        public PdfRender () => Console.WriteLine("No arguments passed.");
        public PdfRender(string arg)
        {
            SetPdfType(arg);
        }
        private enum PageSide
        {
            LeftSide,
            RightSide,
            Middle
        }
        public enum PdfTypes
        {
            Test,
            FacultyStaff,
            Department
        }

        private PdfTypes _pdfType;
        public PdfTypes PdfType
        {
            get => _pdfType;
            set => _pdfType = value;
        }
        private bool Viewer = false;
        private readonly Font FontLargeBold = new Font
        {
            Name = "Times-Roman",
            Size = 14,
            Bold = true,
            Color = Colors.Black
        };
        private readonly Font FontXLarge = new Font
        {
            Name = "Times-Roman",
            Size = 16,
            Bold = true,
            Color = Colors.Black
        };
        private readonly Font FontSmall = new Font
        {
            Name = "Times-Roman",
            Size = 10,
            Color = Colors.Black
        };

        /*
        private readonly Font FontMedium = new Font
        {
            Name = "Times-Roman",
            Size = 12,
            Color = Colors.Black
        };
        */

        private void SetPdfType(string argStr)
        {
            // if an enum is not properly parsed, the first value will be set as the enum (Unknown)
            // Enum.TryParse<PdfTypes>(argStr, true, out PdfType);
            if (Enum.TryParse(argStr, true, out PdfTypes pdfTypes) || Enum.IsDefined(typeof(PdfTypes), PdfType))
            {
                //PdfType = (PdfTypes)Enum.Parse(typeof(PdfTypes), argStr);
                PdfType = pdfTypes;
                //Console.WriteLine("Type parsed:{0} From:{1}", PdfType, pdfTypes);
            }
            else
            {
                PdfType = PdfTypes.Test;
                Console.WriteLine("Undefined type passed.");
            }

            if (PdfType == PdfTypes.Test)
                Viewer = true;
        }
        private Document GenerateTest(string output)
        {
            // create a new MigraDoc
            Document document = CreateTestDocument(output);
            return document;
        }

        private Document CreateTestDocument(string output)
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

        private Table AddTableFacStaff(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("8.75cm");
            table.AddColumn("8.75cm");
            table.Borders.Visible = false;
            table.Borders.Width = 1;
            return table;
        }
        private Table AddTableDept(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("0.5cm");
            table.AddColumn("5cm");
            table.AddColumn("3cm");
            table.AddColumn("8.5cm");
            table.Borders.Visible = false;
            table.Borders.Width = 1;
            table.TopPadding = Unit.FromPoint(3);
            return table;
        }
        private Table AddTableEmergency(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("6cm");
            table.AddColumn("4cm");
            table.AddColumn("3cm");
            table.AddColumn("4.5cm");
            table.Borders.Visible = false;
            table.Borders.Width = 3;
            table.TopPadding = Unit.FromPoint(5);
            return table;
        }
        private Table AddTableBuildingDept(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("5.8cm");
            table.AddColumn("5.8cm");
            table.AddColumn("5.8cm");
            table.Borders.Visible = false;
            //table.Borders.Visible = true;
            table.Borders.Width = 3;
            table.TopPadding = Unit.FromPoint(5);
            return table;
        }
        private Entry FacStaffFillEntry(DataRow dataRow)
        {
            Entry entry = new Entry
            {
                Name = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.Name),
                Building = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.Building),
                PhoneNumber = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.PhoneNumber),
                Department = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.Department),
                EmailAddress = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.EmailAddress),
                FaxNumber = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.FaxNumber),
                Office = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.Office),
                Title = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.Title),
                Url = DataConnectorsExt.GetValue(dataRow, Properties.FacStaffPdf.Default.URL)
            };
            return entry;
        }
        private Entry DepartmentFillEntry(DataRow dataRow)
        {
            Entry entry = new Entry
            {
                Name = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Name),
                Address = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Address),
                Building = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Building),
                PhoneNumber = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.PhoneNumber),
                Category = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Category),
                CellNumber = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.CellNumber),
                Department = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Department),
                EmailAddress = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.EmailAddress),
                FaxNumber = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.FaxNumber),
                Notes = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Notes),
                Office = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Office),
                Title = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.Title),
                Url = DataConnectorsExt.GetValue(dataRow, Properties.DepartmentPdf.Default.URL)
            };
            return entry;
        }
        private void GenerateRow(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            string concatenatedStr = string.Empty;
            Paragraph paragraph;
            FormattedText text;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.RightSide:
                    cellInt = 1;
                    break;
            }

            if (!string.IsNullOrWhiteSpace(entry.Name))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddFormattedText(entry.Name, TextFormat.Bold);
            }
            if(!string.IsNullOrWhiteSpace(entry.Title))
                row.Cells[cellInt].AddParagraph("Title: " + entry.Title);
            if(!string.IsNullOrWhiteSpace(entry.Department))
                row.Cells[cellInt].AddParagraph("Dept: " + entry.Department);
            if (!string.IsNullOrWhiteSpace(entry.Building))
                concatenatedStr = "Bldg: " + entry.Building;
            if (!string.IsNullOrWhiteSpace(entry.Building) && !string.IsNullOrWhiteSpace(entry.Office))
                concatenatedStr += " Office/Room: " + entry.Office;
            if (!string.IsNullOrEmpty(concatenatedStr))
                row.Cells[cellInt].AddParagraph(concatenatedStr);
            if (!string.IsNullOrWhiteSpace(entry.Url))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddText("Home Page: ");
                Hyperlink hyperlink = paragraph.AddHyperlink(entry.Url, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText("Click Here", TextFormat.Underline);
            }
            if (!string.IsNullOrWhiteSpace(entry.EmailAddress))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddText("Email Address: ");
                Hyperlink hyperlink = paragraph.AddHyperlink("mailto:" + entry.EmailAddress, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText(entry.EmailAddress, TextFormat.Underline);
            }
            if (!string.IsNullOrEmpty(entry.PhoneNumber))
                row.Cells[cellInt].AddParagraph("Phone: " + entry.PhoneNumber);
            if (!string.IsNullOrWhiteSpace(entry.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + entry.FaxNumber);
            row.Cells[cellInt].AddParagraph("");
        }
        private void GenerateRowDeptCodes(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            string concateStr = entry.Department + " [" + entry.Notes + "]";

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.RightSide:
                    cellInt = 1;
                    break;
            }

            row.Cells[cellInt].AddParagraph(concateStr).AddLineBreak();
        }
        private void GenerateRowBuildingDept(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            Paragraph paragraph;
            FormattedText text;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.Middle:
                    cellInt = 1;
                    break;
                case PageSide.RightSide:
                    cellInt = 2;
                    break;
            }

            if (!string.IsNullOrWhiteSpace(entry.Department))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                string concatStr = entry.Department;
                if (!string.IsNullOrWhiteSpace(entry.Notes))
                    concatStr += " (" + entry.Notes + ")";
                text.AddText(concatStr);
            }

            if (!string.IsNullOrWhiteSpace(entry.Building))
                row.Cells[cellInt].AddParagraph(entry.Building);
        }
        private void GenerateRowBuildingLocate(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            Paragraph paragraph;
            FormattedText text;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.Middle:
                    cellInt = 1;
                    break;
                case PageSide.RightSide:
                    cellInt = 2;
                    break;
            }

            if (!string.IsNullOrWhiteSpace(entry.Building))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                string concatStr = entry.Building;
                if (!string.IsNullOrWhiteSpace(entry.Notes))
                    concatStr += " (" + entry.Notes + ")";
                text.AddText(concatStr);
            }

            if (!string.IsNullOrWhiteSpace(entry.Address))
                row.Cells[cellInt].AddParagraph(entry.Address);
        }
        private void GenerateRowBoardOfTrustee(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            Paragraph paragraph;
            FormattedText text;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.Middle:
                    cellInt = 1;
                    break;
                case PageSide.RightSide:
                    cellInt = 2;
                    break;
            }

            if (!string.IsNullOrWhiteSpace(entry.Name))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                text.AddText(entry.Name);
            }

            if (!string.IsNullOrWhiteSpace(entry.Title))
                row.Cells[cellInt].AddParagraph(entry.Title);

            if (!string.IsNullOrWhiteSpace(entry.Department))
                row.Cells[cellInt].AddParagraph(entry.Department);

            if (!string.IsNullOrWhiteSpace(entry.Address))
                row.Cells[cellInt].AddParagraph(entry.Address);

            if (!string.IsNullOrWhiteSpace(entry.EmailAddress))
            {
                //row.Cells[cellInt].AddParagraph(employee.EmailAddress);
                paragraph = row.Cells[cellInt].AddParagraph();
                Hyperlink hyperlink = paragraph.AddHyperlink("mailto:" + entry.EmailAddress, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText(entry.EmailAddress, TextFormat.Underline);
            }

            row.Cells[cellInt].AddParagraph("");
        }
        private void GenerateRowOfficersOfAdmin(PageSide pageSide, Entry entry, ref Row row, ref string deptStr)
        {
            int cellInt = 0;
            Paragraph paragraph;
            FormattedText text;

            switch (pageSide)
            {
                case PageSide.LeftSide:
                    cellInt = 0;
                    break;
                case PageSide.RightSide:
                    cellInt = 1;
                    break;
            }

            if(deptStr != entry.Department)
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                deptStr = entry.Department;
                text = paragraph.AddFormattedText();
                text.Underline = Underline.Single;
                text.Bold = true;
                text.AddText(entry.Department);
            }

            if (!string.IsNullOrWhiteSpace(entry.Name))
            {
                row.Cells[cellInt].AddParagraph(entry.Name);
                /*
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                //text.Bold = true;
                //text.Underline = Underline.Single;
                text.AddText(employee.Name);
                */
            }

            if (!string.IsNullOrWhiteSpace(entry.Title))
                row.Cells[cellInt].AddParagraph(entry.Title);

            if (!string.IsNullOrWhiteSpace(entry.Url))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                Hyperlink hyperlink = paragraph.AddHyperlink(entry.Url, HyperlinkType.Url);
                FormattedText formattedText = hyperlink.AddFormattedText();
                formattedText.Font.Color = Color.FromRgb(5, 99, 193);
                formattedText.AddFormattedText(entry.Url, TextFormat.Underline);
                //row.Cells[cellInt].AddParagraph(employee.Url);
            }

            /*
            if (!string.IsNullOrWhiteSpace(employee.Department))
                row.Cells[cellInt].AddParagraph(employee.Department);
            */
            if (!string.IsNullOrWhiteSpace(entry.PhoneNumber))
                row.Cells[cellInt].AddParagraph(entry.PhoneNumber);

            if (!string.IsNullOrWhiteSpace(entry.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + entry.FaxNumber);

            if (!string.IsNullOrWhiteSpace(entry.Address))
                row.Cells[cellInt].AddParagraph(entry.Address);

            if (!string.IsNullOrWhiteSpace(entry.EmailAddress))
                row.Cells[cellInt].AddParagraph(entry.EmailAddress);

            row.Cells[cellInt].AddParagraph("");
        }
        private void AddHeaderCustom(string text, ref Section section)
        {
            HeaderFooter header = new HeaderFooter();
            _ = header.AddParagraph(text);
            Font font = new Font
            {
                Size = 30,
                Name = "Times-Roman"
            };
            header.Format.Font = font;
            header.Format.Alignment = ParagraphAlignment.Center;
            section.Headers.Primary = header.Clone();
        }

        private void AddFooterHelpDesk(ref Section section)
        {
            string text = "Need additional help?" + Environment.NewLine + "Go to ";
            string link = "https://helpdesk.mines.edu";
            HeaderFooter footer = new HeaderFooter();
            Paragraph paragraph = footer.AddParagraph();
            _ = paragraph.AddText(text);
            Hyperlink hyperlink = paragraph.AddHyperlink(link, HyperlinkType.Url);
            FormattedText formattedText = hyperlink.AddFormattedText();
            formattedText.Font.Color = Color.FromRgb(5, 99, 193);
            formattedText.AddFormattedText(link, TextFormat.Underline);
            footer.Format.Font.Size = 10;
            footer.Format.Font.Name = "Times-Roman";
            footer.Format.Alignment = ParagraphAlignment.Center;
            section.Footers.Primary = footer.Clone();
        }
        private void DepartmentCoverPage(ref Section section)
        {
            // add paragraph
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font.Color = Colors.Black;
            paragraph.Format.Font.Name = "Times-Roman";
            paragraph.Format.Font.Size = 32;
            // Add some text to the paragraph
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("Colorado School of Mines", TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("Department Directory", TextFormat.Bold);
            // Add a paragraph to the section
            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font = FontXLarge;
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddText("This document was generated at " + DateTime.Now.ToString("M/d/yyyy h:mm:ss tt"));
            section.AddPageBreak();
        }
        private void DepartmentIntroPage(ref Section section)
        {
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.Format.Font = FontSmall;
            StringCollection stringCollection = Properties.DepartmentPdf.Default.InfoIntro;
            StringBuilder sb = new StringBuilder();
            foreach (string line in stringCollection)
            {
                if (string.IsNullOrEmpty(line))
                    //empty line between text
                    sb.Append(Environment.NewLine + Environment.NewLine);
                else
                    sb.Append(line);
            }
            _ = paragraph.AddText(sb.ToString());
        }
        private void DepartmentServiceCalls(ref Section section)
        {
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.LineSpacingRule = LineSpacingRule.Double;
            //AddFooterHelpDesk(ref section);

            // service calls
            _ = paragraph.AddFormattedText("SERVICE CALLS", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            _ = paragraph.AddFormattedText("Facilities Management", FontLargeBold);
            paragraph.AddLineBreak();
            _ = paragraph.AddFormattedText("Normal Business hours: (303) 273-3330", FontLargeBold);
            paragraph.AddLineBreak();
            _ = paragraph.AddFormattedText("After Hours Emergency Notifications:", FontLargeBold);

            //section = document.AddSection();
            Entry entry = new Entry();
            // add & set the table margins on the new page
            Table table = AddTableDept(section);
            string categoryStr = string.Empty;
            Row row;
            // the class sets the DataProvider
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.ServiceQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            foreach (DataRow dataRow in dataTable.Rows)
            {
                // DepartmentPdf properties
                entry.Category = dataRow[Properties.DepartmentPdf.Default.Category].ToString();
                entry.Name = dataRow[Properties.DepartmentPdf.Default.Name].ToString();
                entry.PhoneNumber = dataRow[Properties.DepartmentPdf.Default.PhoneNumber].ToString();
                entry.Title = dataRow[Properties.DepartmentPdf.Default.Title].ToString();
                entry.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();
                entry.Url = dataRow[Properties.DepartmentPdf.Default.URL].ToString();

                if (categoryStr != entry.Category)
                {
                    paragraph = section.AddParagraph();
                    categoryStr = entry.Category;
                    _ = paragraph.AddFormattedText(Environment.NewLine + entry.Category, TextFormat.Underline);
                    table = AddTableDept(section);
                }

                row = table.AddRow();
                if (!string.IsNullOrEmpty(entry.Title))
                    _ = row.Cells[1].AddParagraph(entry.Title);
                _ = row.Cells[1].AddParagraph(entry.Name);
                if(!string.IsNullOrEmpty(entry.PhoneNumber))
                    _ = row.Cells[2].AddParagraph(entry.PhoneNumber);
                if(!string.IsNullOrEmpty(entry.Url))
                {
                    //_ = row.Cells[2].AddParagraph(employee.Url);
                    paragraph = row.Cells[2].AddParagraph();
                    Hyperlink hyperlink = paragraph.AddHyperlink(entry.Url, HyperlinkType.Url);
                    FormattedText text = hyperlink.AddFormattedText();
                    text.Font.Color = Color.FromRgb(5, 99, 193);
                    text.AddFormattedText(entry.Url, TextFormat.Underline);
                }
                _ = row.Cells[3].AddParagraph(entry.Notes);
            }
        }
        private void DepartmentEmergencyNumbers(ref Section section)
        {
            // Emergency phone numbers
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.LineSpacingRule = LineSpacingRule.Double;
            _ = paragraph.AddFormattedText("EMERGENCY INFORMATION", FontLargeBold);
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.EmergencyQuery
            };
            string categoryStr = string.Empty;
            DataTable dataTable = directoryTasks.GetData();
            Table table = new Table();
            foreach (DataRow dataRow in dataTable.Rows)
            {
                // DepartmentPdf properties
                Entry entry = DepartmentFillEntry(dataRow);

                // Print the category section one time
                if (categoryStr != entry.Category)
                {
                    paragraph = section.AddParagraph();
                    categoryStr = entry.Category;
                    paragraph.AddLineBreak();
                    _ = paragraph.AddFormattedText(entry.Category, TextFormat.Underline);
                    table = AddTableEmergency(section);
                }

                // create a row in the table
                Row row = table.AddRow();
                if (!string.IsNullOrEmpty(entry.Title))
                    _ = row.Cells[0].AddParagraph(entry.Title);
                _ = row.Cells[0].AddParagraph(entry.Name);
                _ = row.Cells[1].AddParagraph(entry.PhoneNumber);
                if (!string.IsNullOrEmpty(entry.CellNumber))
                    _ = row.Cells[2].AddParagraph(entry.CellNumber);
                _ = row.Cells[3].AddParagraph(entry.Notes);
            }
        }
        private void DepartmentGenCampusInfo(ref Section section)
        {
            // General campus info
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("GENERAL CAMPUS AND TELEPHONE INFORMATION", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph = section.AddParagraph();
            StringCollection stringCollection = Properties.DepartmentPdf.Default.InfoCampus;
            StringBuilder sb = new StringBuilder();
            foreach (string line in stringCollection)
            {
                if (string.IsNullOrEmpty(line))
                {
                    //empty line between text
                    sb.Append(Environment.NewLine + Environment.NewLine);
                }
                else
                {
                    if (line.Contains("<newline>"))
                        sb.Append(line.Replace("<newline>", Environment.NewLine));
                    else
                        sb.Append(line);
                }
            }
            _ = paragraph.AddText(sb.ToString());
        }
        private void DepartmentPublicSafety(ref Section section)
        {
            // Public safety information
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("MINES DEPARTMENT OF PUBLIC SAFETY", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph = section.AddParagraph();
            StringCollection stringCollection = Properties.DepartmentPdf.Default.InfoSafety;
            StringBuilder sb = new StringBuilder();
            foreach (string line in stringCollection)
            {
                if (string.IsNullOrEmpty(line))
                {
                    //empty line between text
                    sb.Append(Environment.NewLine + Environment.NewLine);
                }
                else
                {
                    if (line.Contains("<newline>"))
                        sb.Append(line.Replace("<newline>", Environment.NewLine));
                    else
                        sb.Append(line);
                }
            }
            _ = paragraph.AddText(sb.ToString());
        }
        private void DepartmentLocations(ref Section section)
        {
            // Department locations
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("DEPARTMENT LOCATIONS & MAIL CODES", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.DepartmentQuery
            };
            /*
             * "select BuildingName, DeptID, Department from tbl_aux_department inner join tbl_aux_building on 
             * tbl_aux_department.BID = tbl_aux_building.BID order by tbl_aux_building.BuildingName, tbl_aux_department.Department"
             */

            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Row row = new Row();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = DepartmentFillEntry(dataRow);

                // the last row allowed in the table
                if (currPageRow > 45)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    section.AddPageBreak();
                    //section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;

                    // add & set the formatted on the new page
                    table = AddTableBuildingDept(section);
                }

                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 16 || currPageRow == 31)
                    rowInt = 0;

                if (currPageRow <= 15) // rows 1-15
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else if (currPageRow <= 30) // rows 16-30
                {
                    pageSide = PageSide.Middle;
                    row = table.Rows[rowInt];
                }
                else // rows 31-45
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowBuildingDept(pageSide, entry, ref row);
                currPageRow++;
                rowInt++;
            }
        }
        private void DepartmentBuildingLocations(ref Section section)
        {
            // Building locations
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("BUILDING LOCATIONS & ABBREVIATIONS", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            /*
             * "select buildingName, address_line1 from tbl_building where address_line1 is not null order by buildingName"
             */
            /*
             * "select buildingName, address_line1, BuildingCode from tbl_building where address_line1 is not null order by buildingName"
             */

            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.BuildingQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Row row = new Row();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = DepartmentFillEntry(dataRow);

                // the last row allowed in the table
                if (currPageRow > 45)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    section.AddPageBreak();
                    //section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;

                    // add & set the formatted on the new page
                    table = AddTableBuildingDept(section);
                }

                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 16 || currPageRow == 31)
                    rowInt = 0;

                if (currPageRow <= 15) // rows 1-15
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else if (currPageRow <= 30) // rows 16-30
                {
                    pageSide = PageSide.Middle;
                    row = table.Rows[rowInt];
                }
                else // rows 31-45
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowBuildingLocate(pageSide, entry, ref row);
                currPageRow++;
                rowInt++;
            }
        }
        private void DepartmentMailCodes(ref Section section)
        {
            // Department mail codes
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("DEPARMENTAL MAIL CODES", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            // "select department, mailcode from tbl_department where mailcode is not null order by department "
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = "select department, notes from [deptcodes$] order by department"
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableFacStaff(section);
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            Row row = new Row();
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = DepartmentFillEntry(dataRow);

                if (currPageRow > 50)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    section.AddPageBreak();
                    //section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;

                    // add & set the formatted on the new page
                    table = AddTableFacStaff(section);
                }

                if (currPageRow == 26)
                    rowInt = 0;

                if (currPageRow <= 25) // rows 1-25
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else // rows 26-50
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowDeptCodes(pageSide, entry, ref row);
                currPageRow++;
                rowInt++;

            }
        }
        private void DepartmentBoardOfTrustees(ref Section section)
        {
            // Building locations
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("BOARD OF TRUSTEES", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            /*
             * "select buildingName, address_line1 from tbl_building where address_line1 is not null order by buildingName"
             */
            /*
             * "select buildingName, address_line1, BuildingCode from tbl_building where address_line1 is not null order by buildingName"
             */

            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.TrusteesQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableFacStaff(section);
            Row row = new Row();
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = DepartmentFillEntry(dataRow);

                // the last row allowed in the table
                if (currPageRow > 20)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    section.AddPageBreak();
                    //section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;

                    // add & set the formatted on the new page
                    table = AddTableFacStaff(section);
                }

                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 11)
                    rowInt = 0;

                if (currPageRow <= 10) // rows 1-10
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else // rows 11-20
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowBoardOfTrustee(pageSide, entry, ref row);
            }
        }
        private void DepartmentOfficersOfAdmin(ref Section section)
        {
            // Building locations
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("OFFICERS OF ADMINISTRATION", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            /*
             * "select buildingName, address_line1 from tbl_building where address_line1 is not null order by buildingName"
             */
            /*
             * "select buildingName, address_line1, BuildingCode from tbl_building where address_line1 is not null order by buildingName"
             */

            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.OfficersQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableFacStaff(section);
            Row row = new Row();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;
            string deptStr = string.Empty;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = DepartmentFillEntry(dataRow);

                // the last row allowed in the table
                if (currPageRow > 24)
                {
                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    section.AddPageBreak();
                    //section = document.AddSection();

                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;

                    // add & set the formatted on the new page
                    table = AddTableFacStaff(section);
                }

                // 2nd columns beginning value
                if (currPageRow == 13)
                    rowInt = 0;

                if (currPageRow <= 12) // rows 1-12
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else // rows 13-24
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowOfficersOfAdmin(pageSide, entry, ref row, ref deptStr);
                currPageRow++;
                rowInt++;
            }
        }

        private Document CreateDocumentDepartment()
        {
            // create document
            Document document = new Document();
            // create section
            Section section = document.AddSection();

            // Cover page
            DepartmentCoverPage(ref section);
            // Introduction/preamble page
            DepartmentIntroPage(ref section);

            // new section is needed to exclude footers and headers from previous section
            section = document.AddSection();
            AddFooterHelpDesk(ref section);

            // Service calls pages
            DepartmentServiceCalls(ref section);

            // Emergency phone numbers pages
            DepartmentEmergencyNumbers(ref section);

            // General campus info pages
            DepartmentGenCampusInfo(ref section);

            // Public safety info pages
            DepartmentPublicSafety(ref section);

            // Department locations & mail codes pages
            DepartmentLocations(ref section);

            // Building locations & mail codes pages
            DepartmentBuildingLocations(ref section);

            // This is being refactored into DepartmentLocations();
            // Department mail code pages
            //DepartmentMailCodes(ref section);

            // Board of trustees
            DepartmentBoardOfTrustees(ref section);

            /*
             "select buildingName, BuildingCode from tbl_building where BuildingCode is not null "
            "and BuildingCode not like '%x%' and BuildingCode not like '%y%' "
            "and BuildingCode not in ('MO','RA','TH','WT') and BuildingCode not in ('1S','2S','3S') "
            "and BuildingCode not in ('4C','5C') order by buildingName "
             */

            // Officers of Administration
            DepartmentOfficersOfAdmin(ref section);

            return document;
        }

        private Document CreateDocumentFacultyStaff()
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
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            _ = paragraph.AddFormattedText("Colorado School of Mines" + Environment.NewLine + "Faculty/Staff Directory", TextFormat.Bold);

            // Add a paragraph to the section
            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            paragraph.Format.Font.Color = Colors.Black;
            paragraph.Format.Font.Name = "Times-Roman";
            paragraph.Format.Font.Size = 16;
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            _ = paragraph.AddText("This document was generated at " + DateTime.Now.ToString("M/d/yyyy h:mm:ss tt"));
            section.AddPageBreak();

            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            Table table = new Table();
            // The class sets the DataProvider, ConnectionString and QueryString
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType);
            DataTable dataTable = directoryTasks.GetData();
            Console.WriteLine("Rows: {0}", dataTable.Rows.Count);
            string currHeaderStr = " ";

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Entry entry = FacStaffFillEntry(dataRow);

                if (!entry.Name.StartsWith(currHeaderStr) || currPageRow > 16)
                {
                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;
                    currHeaderStr = entry.Name.Substring(0, 1);

                    // usually you would add a page break to the section
                    // however, because we need different headers a new section is added to the document instead
                    //section.AddPageBreak();
                    section = document.AddSection();
                    AddFooterHelpDesk(ref section);

                    // update page header
                    AddHeaderCustom(currHeaderStr, ref section);
                    // add & set the formatted on the new page
                    table = AddTableFacStaff(section);
                }

                PageSide pageSide;
                Row row = new Row();

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

                GenerateRow(pageSide, entry, ref row);
                currPageRow++;
                rowInt++;
            }
            return document;
        }

        public void RequestPdf(string arg)
        {
            string filename;
            // Create a MigraDoc document
            Document document;

            switch(PdfType)
            {
                case PdfTypes.FacultyStaff:
                    Console.WriteLine("This will render the Mines Faculty/Staff PDF.");
                    document = CreateDocumentFacultyStaff();
                    filename = "fac_staff_dir.pdf";
                    break;
                case PdfTypes.Department:
                    Console.WriteLine("This will render the Mines Department Directory PDF");
                    document = CreateDocumentDepartment();
                    filename = "departmental_dir.pdf";
                    break;
                default:
                    Console.WriteLine("This will render the Test PDF");
                    document = GenerateTest(arg);
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
