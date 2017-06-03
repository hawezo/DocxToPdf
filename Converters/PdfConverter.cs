using DocxToPdf.Core;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxToPdf.Converters
{

    /// <summary>
    /// Static Html to Pdf converter.
    /// </summary>
    public static class PdfConverter
    {

        /// <summary>
        /// Converts the given HTML string into a PDF file.
        /// </summary>
        /// <param name="html">HTML string (needs to be well formatted).</param>
        /// <param name="output">Output PDF file.</param>
        public static void ConvertFromHtml(string html, string output)
        {

            // Uses HtmlRenderer of PdfSharp to convert HTML to PDF
            try
            {
                PdfDocument pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.A4);
                pdf.Save(output);
            }

            // File in-use or insufficient permission
            catch (IOException)
            {
                Logger.Error("An error occured while trying to save pdf file. Please make sure output file isn't in use or try to launch as adminitrator.");
                throw;
            }

            // Unknown error
            catch (Exception)
            {
                throw;
            }
        }

    }
}
