using DocxToPdf.Core;
using ExCSS;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
                HtmlDocument doc = new HtmlDocument(); doc.LoadHtml(html);
                HtmlNode style = doc.DocumentNode.SelectSingleNode("//style");

                html = html.Replace(style.InnerText, PdfConverter.ReplaceUnknownFonts(style.InnerText));
                PdfSharp.Pdf.PdfDocument pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.A4);
                pdf.Save(output);
                pdf.Close();
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

        private static string ReplaceUnknownFonts(string css)
        {
            List<String> allowedFonts = new List<string>
            {
                "Arial",
                "Times New Roman",
                "Serif",
                "Tahoma",
                "Segoe UI",
                "Comic Sans",
                "Calibri",
                "Cambria",
                "Consolas"
            };

            ExCSS.StyleSheet stylesheet = new Parser().Parse(css);

            foreach (StyleRule rule in stylesheet.StyleRules)
            {
                if (rule.Declarations.ToString().Contains("font-family"))
                {
                    Property font = rule.Declarations.FirstOrDefault(d => d.Name.Equals("font-family", StringComparison.InvariantCultureIgnoreCase));

                    bool allowed = false;
                    foreach (string allowedFont in allowedFonts)
                        if (font.Term.ToString().ToLower().Trim() == allowedFont.ToLower().Trim())
                            allowed = true;

                    if (!allowed)
                        font.Term = new PrimitiveTerm(UnitType.String, "Times New Roman");
                }
            }

            return css.Replace(css, stylesheet.ToString());
        }
    }
}
