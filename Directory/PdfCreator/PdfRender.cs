﻿using MigraDoc.DocumentObjectModel;
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
        /// <summary>
        /// Page sides; useful in setting position within table structures on a page
        /// </summary>
        private enum PageSide
        {
            LeftSide,
            RightSide,
            Middle
        }
        /// <summary>
        /// Types that are in use within the class
        /// </summary>
        public enum PdfTypes
        {
            Test,
            FacultyStaff,
            Department
        }

        public PdfTypes PdfType { get; set; }
        // launches a PDF viewer when running a test
        private bool Viewer = false;
        /// <summary>
        /// Times-Roman, 14pt, black, bold
        /// </summary>
        private readonly Font FontLargeBold = new Font
        {
            Name = "Times-Roman",
            Size = 14,
            Color = Colors.Black,
            Bold = true
        };
        /// <summary>
        /// Times-Roman, 16pt, black
        /// </summary>
        private readonly Font FontXLarge = new Font
        {
            Name = "Times-Roman",
            Size = 16,
            Bold = true,
            Color = Colors.Black
        };
        /// <summary>
        /// Times-Roman, 10pt, black
        /// </summary>
        private readonly Font FontSmall = new Font
        {
            Name = "Times-Roman",
            Size = 10,
            Color = Colors.Black
        };
        /// <summary>
        /// Times-Roman, 30pt, black
        /// </summary>
        private readonly Font FontHeader = new Font
        {
            Name = "Times-Roman",
            Size = 30,
            Color = Colors.Black
        };
        /// <summary>
        /// Times-Roman, 32pt, black
        /// </summary>
        private readonly Font FontCoverPage = new Font
        {
            Name = "Times-Roman",
            Size = 32,
            Color = Colors.Black
        };
        private void SetPdfType(string argStr)
        {
            // if an enum is not properly parsed, the first value will be set as the enum (Unknown)
            if (Enum.TryParse(argStr, true, out PdfTypes pdfTypes) || Enum.IsDefined(typeof(PdfTypes), PdfType))
            {
                PdfType = pdfTypes;
            }
            else
            {
                PdfType = PdfTypes.Test;
                Console.WriteLine("Undefined type passed.");
            }
            if (PdfType == PdfTypes.Test)
                Viewer = true;
        }
        private StringBuilder DepartmentStringBuild(StringCollection stringCollection)
        {
            // read multi-string collection from properties; let the flag "<newline>" set a new line
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
            return sb;
        }
        private void AddEmailAddress(ref Paragraph paragraph, string email)
        {
            Hyperlink hyperlink = paragraph.AddHyperlink("mailto:" + email, HyperlinkType.Url);
            FormattedText formattedText = hyperlink.AddFormattedText();
            formattedText.Font.Color = Color.FromRgb(5, 99, 193);
            formattedText.AddFormattedText(email, TextFormat.Underline);
        }
        private void AddUrl(ref Paragraph paragraph, string url, string linkText = null)
        {
            Hyperlink hyperlink = paragraph.AddHyperlink(url, HyperlinkType.Url);
            FormattedText formattedText = hyperlink.AddFormattedText();
            formattedText.Font.Color = Color.FromRgb(5, 99, 193);
            if(linkText == null)
                linkText = url;
            formattedText.AddFormattedText(linkText, TextFormat.Underline);
        }
        /// <summary>
        /// Generates a test PDF based on input string passed
        /// </summary>
        /// <param name="output"></param>
        /// <returns></returns>
        private Document CreateDocumentTest(string output)
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
        /// Two columns each at 8.75cm
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        private Table AddTableTwoColumns(Section section)
        {
            Table table = section.AddTable();
            table.AddColumn("8.75cm");
            table.AddColumn("8.75cm");
            table.Borders.Visible = false;
            table.Borders.Width = 1;
            return table;
        }
        /// <summary>
        /// Four columns: 0.5cm, 5cm, 3cm, 8.5cm
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
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
        /// <summary>
        /// Four columns: 6cm, 4cm, 3cm, 4.5cm
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
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
        /// <summary>
        /// Three columns each at 5.8cm
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        private Table AddTableThreeColumns(Section section)
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
        /// <summary>
        /// Fill the faculty/staff entry properties from the row.
        /// Only properties within the row will be filled
        /// </summary>
        /// <param name="dataRow"></param>
        /// <returns></returns>
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
        /// <summary>
        /// Fill the department entry properties from the row.
        /// Only properties within the row will be filled
        /// </summary>
        /// <param name="dataRow"></param>
        /// <returns></returns>
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
        private void GenerateRowFacStaff(PageSide pageSide, Entry entry, ref Row row)
        {
            int cellInt = 0;
            string concatenatedStr = string.Empty;
            Paragraph paragraph;

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
                AddUrl(ref paragraph, entry.Url, "Click Here");
            }
            if (!string.IsNullOrWhiteSpace(entry.EmailAddress))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                paragraph.AddText("Email Address: ");
                AddEmailAddress(ref paragraph, entry.EmailAddress);
            }
            if (!string.IsNullOrEmpty(entry.PhoneNumber))
                row.Cells[cellInt].AddParagraph("Phone: " + entry.PhoneNumber);
            if (!string.IsNullOrWhiteSpace(entry.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + entry.FaxNumber);
            row.Cells[cellInt].AddParagraph("");
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
                Paragraph paragraph = row.Cells[cellInt].AddParagraph();
                FormattedText formattedText = paragraph.AddFormattedText();
                formattedText.Bold = true;
                string concatStr = entry.Building;
                if (!string.IsNullOrWhiteSpace(entry.Notes))
                    concatStr += " (" + entry.Notes + ")";
                formattedText.AddText(concatStr);
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
                paragraph = row.Cells[cellInt].AddParagraph();
                AddEmailAddress(ref paragraph, entry.EmailAddress);
            }

            row.Cells[cellInt].AddParagraph("");
        }
        private void GenerateRowOfficersOfAdmin(PageSide pageSide, Entry entry, ref Row row, ref string deptStr)
        {
            int cellInt = 0;
            Paragraph paragraph;
            FormattedText formattedText;

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
                formattedText = paragraph.AddFormattedText();
                formattedText.Underline = Underline.Single;
                formattedText.Bold = true;
                formattedText.AddText(entry.Department);
            }
            if (!string.IsNullOrWhiteSpace(entry.Name))
                row.Cells[cellInt].AddParagraph(entry.Name);
            if (!string.IsNullOrWhiteSpace(entry.Title))
                row.Cells[cellInt].AddParagraph(entry.Title);
            if (!string.IsNullOrWhiteSpace(entry.Url))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                AddUrl(ref paragraph, entry.Url);
            }
            if (!string.IsNullOrWhiteSpace(entry.PhoneNumber))
                row.Cells[cellInt].AddParagraph(entry.PhoneNumber);
            if (!string.IsNullOrWhiteSpace(entry.FaxNumber))
                row.Cells[cellInt].AddParagraph("Fax: " + entry.FaxNumber);
            if (!string.IsNullOrWhiteSpace(entry.Address))
                row.Cells[cellInt].AddParagraph(entry.Address);
            if (!string.IsNullOrWhiteSpace(entry.EmailAddress))
            {
                paragraph = row.Cells[cellInt].AddParagraph();
                AddEmailAddress(ref paragraph, entry.EmailAddress);
            }
            row.Cells[cellInt].AddParagraph("");
        }
        private void AddHeaderCustom(string text, ref Section section)
        {
            HeaderFooter header = new HeaderFooter();
            _ = header.AddParagraph(text);
            header.Format.Font = FontHeader.Clone();
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
            AddUrl(ref paragraph, link);
            footer.Format.Font = FontSmall.Clone();
            footer.Format.Alignment = ParagraphAlignment.Center;
            section.Footers.Primary = footer.Clone();
        }
        private void DepartmentCoverPage(ref Section section)
        {
            // add paragraph
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font = FontCoverPage.Clone();
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
        }
        private void DepartmentIntroPage(ref Section section)
        {
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.Format.Font = FontSmall;
            StringBuilder sb = DepartmentStringBuild(Properties.DepartmentPdf.Default.InfoIntro);
            _ = paragraph.AddText(sb.ToString());
        }
        private void DepartmentServiceCalls(ref Section section)
        {
            // start new page
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.LineSpacingRule = LineSpacingRule.Double;
            // service calls
            _ = paragraph.AddFormattedText("SERVICE CALLS", FontLargeBold);
            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            // add & set the table margins on the new page
            Table table = new Table();
            string categoryStr = string.Empty;
            // the class sets the DataProvider
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.ServiceQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            foreach (DataRow dataRow in dataTable.Rows)
            {
                // DepartmentPdf properties
                Entry entry = DepartmentFillEntry(dataRow);
                if (categoryStr != entry.Category)
                {
                    paragraph = section.AddParagraph();
                    categoryStr = entry.Category;
                    paragraph.AddLineBreak();
                    _ = paragraph.AddFormattedText(entry.Category, TextFormat.Underline);
                    table = AddTableDept(section);
                }
                Row row = table.AddRow();
                if (!string.IsNullOrEmpty(entry.Title))
                    _ = row.Cells[1].AddParagraph(entry.Title);
                _ = row.Cells[1].AddParagraph(entry.Name);
                if(!string.IsNullOrEmpty(entry.PhoneNumber))
                    _ = row.Cells[2].AddParagraph(entry.PhoneNumber);
                if(!string.IsNullOrEmpty(entry.Url))
                {
                    paragraph = row.Cells[2].AddParagraph();
                    AddUrl(ref paragraph, entry.Url);
                }
                if(!string.IsNullOrEmpty(entry.Notes))
                   _ = row.Cells[3].AddParagraph(entry.Notes);
            }
        }
        /// <summary>
        /// Emergency phone numbers section of the department PDF
        /// </summary>
        /// <param name="section"></param>
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
                if(!string.IsNullOrEmpty(entry.Notes))
                    _ = row.Cells[3].AddParagraph(entry.Notes);
            }
        }
        /// <summary>
        /// General campus information in the department PDF
        /// </summary>
        /// <param name="section"></param>
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
            StringBuilder sb = DepartmentStringBuild(Properties.DepartmentPdf.Default.InfoCampus);
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
            StringBuilder sb = DepartmentStringBuild(Properties.DepartmentPdf.Default.InfoSafety);
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
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableThreeColumns(section);
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
                    table = AddTableThreeColumns(section);
                }
                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 16 || currPageRow == 31)
                    rowInt = 0;
                Row row = new Row();
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
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.BuildingQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableThreeColumns(section);
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
                    table = AddTableThreeColumns(section);
                }
                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 16 || currPageRow == 31)
                    rowInt = 0;
                Row row = new Row();
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
        private void DepartmentBoardOfTrustees(ref Section section)
        {
            // Building locations
            section.AddPageBreak();
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("BOARD OF TRUSTEES", FontLargeBold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.TrusteesQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableTwoColumns(section);
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
                    table = AddTableTwoColumns(section);
                }

                // 2nd columns beginning value; 3rd columns beginning value
                if (currPageRow == 11)
                    rowInt = 0;
                Row row = new Row();
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
                currPageRow++;
                rowInt++;
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
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType)
            {
                QueryString = Properties.DepartmentPdf.Default.OfficersQuery
            };
            DataTable dataTable = directoryTasks.GetData();
            Table table = AddTableTwoColumns(section);
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
                    table = AddTableTwoColumns(section);
                }
                // 2nd columns beginning value
                if (currPageRow == 13)
                    rowInt = 0;
                Row row = new Row();

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
            // Board of trustees
            DepartmentBoardOfTrustees(ref section);
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
            paragraph.Format.Font = FontCoverPage.Clone();
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
            paragraph.Format.Font = FontXLarge.Clone();
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            _ = paragraph.AddText("This document was generated at " + DateTime.Now.ToString("M/d/yyyy h:mm:ss tt"));
            section.AddPageBreak();
            // initializer for right side count
            int rowInt = 0;
            int currPageRow = 1;
            Table table = new Table();
            DirectoryTasks directoryTasks = new DirectoryTasks(PdfType);
            DataTable dataTable = directoryTasks.GetData();
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
                    table = AddTableTwoColumns(section);
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
                GenerateRowFacStaff(pageSide, entry, ref row);
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
                    Console.WriteLine("Faculty/Staff PDF.");
                    document = CreateDocumentFacultyStaff();
                    filename = Properties.FacStaffPdf.Default.Filename;
                    break;
                case PdfTypes.Department:
                    Console.WriteLine("Department Directory PDF");
                    document = CreateDocumentDepartment();
                    filename = Properties.DepartmentPdf.Default.Filename;
                    break;
                default:
                    Console.WriteLine("Test PDF");
                    document = CreateDocumentTest(arg);
                    filename = "Test.pdf";
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
