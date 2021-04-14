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
                Console.WriteLine("Type parsed:{0} From:{1}", PdfType, pdfTypes);
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
        private void GenerateRow(PageSide pageSide, Employee employee, ref Row row)
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

            if (!string.IsNullOrWhiteSpace(employee.Name))
            {
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
                paragraph.AddText("Home Page: ");
                Hyperlink hyperlink = paragraph.AddHyperlink(employee.Url, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText("Click Here", TextFormat.Underline);
            }
            if (!string.IsNullOrWhiteSpace(employee.EmailAddress))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddText("Email Address: ");
                Hyperlink hyperlink = paragraph.AddHyperlink("mailto:" + employee.EmailAddress, HyperlinkType.Url);
                text = hyperlink.AddFormattedText();
                text.Font.Color = Color.FromRgb(5, 99, 193);
                text.AddFormattedText(employee.EmailAddress, TextFormat.Underline);
            }
            if (!string.IsNullOrEmpty(employee.PhoneNumber))
                row.Cells[cellInt].AddParagraph("Phone: " + employee.PhoneNumber);
            if (!string.IsNullOrWhiteSpace(employee.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + employee.FaxNumber);
            row.Cells[cellInt].AddParagraph("");
        }
        private void GenerateRowDeptCodes(PageSide pageSide, Employee employee, ref Row row)
        {
            int cellInt = 0;
            string concateStr = employee.Department + " [" + employee.Notes + "]";

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
        private void GenerateRowBuildingDept(PageSide pageSide, Employee employee, ref Row row)
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

            if (!string.IsNullOrWhiteSpace(employee.Department))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                string concatStr = employee.Department;
                if (!string.IsNullOrWhiteSpace(employee.Notes))
                    concatStr += " (" + employee.Notes + ")";
                text.AddText(concatStr);
            }

            if (!string.IsNullOrWhiteSpace(employee.Building))
                row.Cells[cellInt].AddParagraph(employee.Building);
        }
        private void GenerateRowBuildingLocate(PageSide pageSide, Employee employee, ref Row row)
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

            if (!string.IsNullOrWhiteSpace(employee.Building))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                string concatStr = employee.Building;
                if (!string.IsNullOrWhiteSpace(employee.Notes))
                    concatStr += " (" + employee.Notes + ")";
                text.AddText(concatStr);
            }

            if (!string.IsNullOrWhiteSpace(employee.Address))
                row.Cells[cellInt].AddParagraph(employee.Address);
        }
        private void GenerateRowBoardOfTrustee(PageSide pageSide, Employee employee, ref Row row)
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

            if (!string.IsNullOrWhiteSpace(employee.Name))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                text.AddText(employee.Name);
            }

            if (!string.IsNullOrWhiteSpace(employee.Title))
                row.Cells[cellInt].AddParagraph(employee.Title);

            if (!string.IsNullOrWhiteSpace(employee.Department))
                row.Cells[cellInt].AddParagraph(employee.Department);

            if (!string.IsNullOrWhiteSpace(employee.Address))
                row.Cells[cellInt].AddParagraph(employee.Address);

            if (!string.IsNullOrWhiteSpace(employee.EmailAddress))
                row.Cells[cellInt].AddParagraph(employee.EmailAddress);
        }
        private void GenerateRowOfficersOfAdmin(PageSide pageSide, Employee employee, ref Row row)
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

            if (!string.IsNullOrWhiteSpace(employee.Name))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                text = paragraph.AddFormattedText();
                text.Bold = true;
                //text.Underline = Underline.Single;
                text.AddText(employee.Name);
            }

            if (!string.IsNullOrWhiteSpace(employee.Title))
                row.Cells[cellInt].AddParagraph(employee.Title);

            if (!string.IsNullOrWhiteSpace(employee.Department))
                row.Cells[cellInt].AddParagraph(employee.Department);

            if (!string.IsNullOrWhiteSpace(employee.Address))
                row.Cells[cellInt].AddParagraph(employee.Address);

            if (!string.IsNullOrWhiteSpace(employee.EmailAddress))
                row.Cells[cellInt].AddParagraph(employee.EmailAddress);
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
            Employee employee = new Employee();
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
                employee.Category = dataRow[Properties.DepartmentPdf.Default.Category].ToString();
                employee.Name = dataRow[Properties.DepartmentPdf.Default.Name].ToString();
                employee.PhoneNumber = dataRow[Properties.DepartmentPdf.Default.PhoneNumber].ToString();
                employee.Title = dataRow[Properties.DepartmentPdf.Default.Title].ToString();
                employee.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();

                if (categoryStr != employee.Category)
                {
                    paragraph = section.AddParagraph();
                    categoryStr = employee.Category;
                    _ = paragraph.AddFormattedText(Environment.NewLine + employee.Category, TextFormat.Underline);
                    table = AddTableDept(section);
                }

                row = table.AddRow();
                if (!string.IsNullOrEmpty(employee.Title))
                    _ = row.Cells[1].AddParagraph(employee.Title);
                _ = row.Cells[1].AddParagraph(employee.Name);
                _ = row.Cells[2].AddParagraph(employee.PhoneNumber);
                _ = row.Cells[3].AddParagraph(employee.Notes);
            }
        }
        private void DepartmentEmergencyNumbers(ref Section section)
        {
            // Emergency phone numbers
            section.AddPageBreak();
            Employee employee = new Employee();
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
                employee.Category = dataRow[Properties.DepartmentPdf.Default.Category].ToString();
                employee.Name = dataRow[Properties.DepartmentPdf.Default.Name].ToString();
                employee.PhoneNumber = dataRow[Properties.DepartmentPdf.Default.PhoneNumber].ToString();
                employee.Title = dataRow[Properties.DepartmentPdf.Default.Title].ToString();
                employee.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();
                employee.CellNumber = dataRow[Properties.DepartmentPdf.Default.CellNumber].ToString();

                // Print the category section one time
                if (categoryStr != employee.Category)
                {
                    paragraph = section.AddParagraph();
                    categoryStr = employee.Category;
                    paragraph.AddLineBreak();
                    _ = paragraph.AddFormattedText(employee.Category, TextFormat.Underline);
                    table = AddTableEmergency(section);
                }

                // create a row in the table
                Row row = table.AddRow();
                if (!string.IsNullOrEmpty(employee.Title))
                    _ = row.Cells[0].AddParagraph(employee.Title);
                _ = row.Cells[0].AddParagraph(employee.Name);
                _ = row.Cells[1].AddParagraph(employee.PhoneNumber);
                if (!string.IsNullOrEmpty(employee.CellNumber))
                    _ = row.Cells[2].AddParagraph(employee.CellNumber);
                _ = row.Cells[3].AddParagraph(employee.Notes);
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
                QueryString = "select Campusbuilding, department, notes from [departments$] order by department, campusbuilding"
            };
            /*
             * "select BuildingName, DeptID, Department from tbl_aux_department inner join tbl_aux_building on 
             * tbl_aux_department.BID = tbl_aux_building.BID order by tbl_aux_building.BuildingName, tbl_aux_department.Department"
             */

            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Row row = new Row();
            Employee employee = new Employee();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Building = dataRow[Properties.DepartmentPdf.Default.Building].ToString();
                employee.Department = dataRow[Properties.DepartmentPdf.Default.Department].ToString();
                employee.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();

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

                GenerateRowBuildingDept(pageSide, employee, ref row);
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
                QueryString = "select campusbuilding, address, notes from [buildings$] order by campusBuilding"
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Employee employee = new Employee();
            Row row = new Row();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Building = dataRow[Properties.DepartmentPdf.Default.Building].ToString();
                employee.Address = dataRow[Properties.DepartmentPdf.Default.Address].ToString();
                employee.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();

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

                GenerateRowBuildingLocate(pageSide, employee, ref row);
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
            Employee employee = new Employee();
            Row row = new Row();
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Department = dataRow[Properties.DepartmentPdf.Default.Department].ToString();
                employee.Notes = dataRow[Properties.DepartmentPdf.Default.Notes].ToString();

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

                GenerateRowDeptCodes(pageSide, employee, ref row);
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
                QueryString = "select empname, title, department, address, email from [trustees$] order by catOrder"
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Employee employee = new Employee();
            Row row = new Row();
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Name = dataRow[Properties.DepartmentPdf.Default.Name].ToString();
                employee.Title = dataRow[Properties.DepartmentPdf.Default.Title].ToString();
                employee.Department = dataRow[Properties.DepartmentPdf.Default.Department].ToString();
                employee.Address = dataRow[Properties.DepartmentPdf.Default.Address].ToString();
                employee.EmailAddress = dataRow[Properties.DepartmentPdf.Default.EmailAddress].ToString();

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

                GenerateRowBoardOfTrustee(pageSide, employee, ref row);
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
                QueryString = "select empname, title, department, address, email, workPhone from [officers$]"
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableBuildingDept(section);
            Employee employee = new Employee();
            Row row = new Row();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            PageSide pageSide;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Name = dataRow[Properties.DepartmentPdf.Default.Name].ToString();
                employee.Title = dataRow[Properties.DepartmentPdf.Default.Title].ToString();
                employee.Department = dataRow[Properties.DepartmentPdf.Default.Department].ToString();
                employee.Address = dataRow[Properties.DepartmentPdf.Default.Address].ToString();
                employee.EmailAddress = dataRow[Properties.DepartmentPdf.Default.EmailAddress].ToString();
                employee.PhoneNumber = dataRow[Properties.DepartmentPdf.Default.PhoneNumber].ToString();

                // the last row allowed in the table
                if (currPageRow > 36)
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
                if (currPageRow == 13 || currPageRow == 27)
                    rowInt = 0;

                if (currPageRow <= 12) // rows 1-12
                {
                    pageSide = PageSide.LeftSide;
                    row = table.AddRow();
                }
                else if (currPageRow <= 24) // rows 13-24
                {
                    pageSide = PageSide.Middle;
                    row = table.Rows[rowInt];
                }
                else // rows 25-36
                {
                    pageSide = PageSide.RightSide;
                    row = table.Rows[rowInt];
                }

                GenerateRowOfficersOfAdmin(pageSide, employee, ref row);
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
            Employee employee = new Employee();
            // The class sets the DataProvider, ConnectionString and QueryString
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType);
            DataTable dataTable = directoryTasks.GetData();
            Console.WriteLine("Rows: {0}", dataTable.Rows.Count);
            string currHeaderStr = "X";

            foreach (DataRow dataRow in dataTable.Rows)
            {
                employee.ClearAll();
                employee.Name = dataRow[Properties.FacStaffPdf.Default.Name].ToString();
                employee.PhoneNumber = dataRow[Properties.FacStaffPdf.Default.PhoneNumber].ToString();
                employee.Department = dataRow[Properties.FacStaffPdf.Default.Department].ToString();
                employee.FaxNumber = dataRow[Properties.FacStaffPdf.Default.FaxNumber].ToString();
                employee.EmailAddress = dataRow[Properties.FacStaffPdf.Default.EmailAddress].ToString();
                employee.Title = dataRow[Properties.FacStaffPdf.Default.Title].ToString();
                employee.Building = dataRow[Properties.FacStaffPdf.Default.Building].ToString();
                employee.Office = dataRow[Properties.FacStaffPdf.Default.Office].ToString();
                employee.Url = dataRow[Properties.FacStaffPdf.Default.URL].ToString();

                if (!employee.Name.StartsWith(currHeaderStr) || currPageRow > 16)
                {
                    // reset variables
                    currPageRow = 1;
                    rowInt = 0;
                    currHeaderStr = employee.Name.Substring(0, 1);

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

                GenerateRow(pageSide, employee, ref row);
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
