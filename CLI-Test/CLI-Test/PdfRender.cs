using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Diagnostics;

namespace CLI_Test
{
    class PdfRender
    {
        public void GenerateTest()
        {
            GenerateTest("Hello World!");
        }
        public void GenerateTest(string output)
        {
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
        }

        public void FacultyStaffPDF()
        {
            Console.WriteLine("This will render the Faculty/Staff PDF.");

            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Mines Directory";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Times-Roman", 32, XFontStyle.Regular);

            XTextFormatter tf = new XTextFormatter(gfx);

            // Draw the text
            // X, Y, x, y
            XRect rect = new XRect(100, 200, 400, 200);
            tf.Alignment = XParagraphAlignment.Center;
            tf.DrawString("Colorado School of Mines" + Environment.NewLine + "Faculty/Staff Directory", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(80, 300, 450, 300);
            font = new XFont("Time-Roman", 16, XFontStyle.Regular);
            tf.DrawString("This document was generated at " + DateTime.Now.ToString("M/d/yyyy h:mm:ss tt"), font, XBrushes.Black, rect, XStringFormats.TopLeft);

            document.AddPage();

            // Save the document...
            const string filename = "fac_staff_dir.pdf";
            document.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }
    }
}
