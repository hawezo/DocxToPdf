using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DocxToPdf.Converters
{

    /// <summary>
    /// Static Docx to Htmlconverter.
    /// </summary>
    public static class HtmlConverter
    {

        /// <summary>
        /// Converts any DOCX file into a valid HTML page. Generate an image folder.
        /// </summary>
        /// <param name="file">Input DOCX file to be converted.</param>
        /// <returns>Returns the HTML string. Images are in the input file's directory.</returns>
        public static string ConvertFromDocx(string file)
        {

            // NOT MY WORK
            // See: https://github.com/OfficeDev/Open-Xml-PowerTools/blob/vNext/OpenXmlPowerToolsExamples/HtmlConverter01/HtmlConverter01.cs
            try
            {
                FileInfo fi = new FileInfo(file);

                byte[] bytes = File.ReadAllBytes(fi.FullName);

                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                    using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
                    {
                        string imageDirectory = file.Replace(".docx", null) + "_images";
                        int imageCount = 0;

                        string title = fi.FullName;
                        CoreFilePropertiesPart part = document.CoreFilePropertiesPart;
                        if (part != null)
                            title = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;

                        // Html settings and image handler
                        HtmlConverterSettings settings = new HtmlConverterSettings()
                        {
                            AdditionalCss = "body { margin: 0.1cm auto; max-width: 20cm; padding: 0; }",
                            PageTitle = title,
                            FabricateCssClasses = true,
                            CssClassPrefix = "pt-",
                            RestrictToSupportedLanguages = false,
                            RestrictToSupportedNumberingFormats = false,
                            ImageHandler = imageInfo =>
                            {
                                DirectoryInfo di = new DirectoryInfo(imageDirectory);
                                if (!di.Exists)
                                    di.Create();
                                ++imageCount;

                                string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                ImageFormat imageFormat = null;
                                if (extension == "png")
                                    imageFormat = ImageFormat.Png;
                                else if (extension == "gif")
                                    imageFormat = ImageFormat.Gif;
                                else if (extension == "bmp")
                                    imageFormat = ImageFormat.Bmp;
                                else if (extension == "jpeg")
                                    imageFormat = ImageFormat.Jpeg;
                                else if (extension == "tiff")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "x-wmf")
                                {
                                    extension = "wmf";
                                    imageFormat = ImageFormat.Wmf;
                                }

                                if (imageFormat == null)
                                    return null;

                                string imageFileName = Path.Combine(imageDirectory, "image" + imageCount.ToString() + "." + extension);
                                try
                                {
                                    imageInfo.Bitmap.Save(imageFileName, imageFormat);
                                }
                                catch (System.Runtime.InteropServices.ExternalException)
                                {
                                    return null;
                                }

                                string imageSource = Path.Combine(di.Name, "image" + imageCount.ToString() + "." + extension);
                                XElement img = new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource),
                                                                       imageInfo.ImgStyleAttribute,
                                                                       imageInfo.AltText != null ?
                                                                       new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                                return img;
                            }
                        };

                        XElement htmlElement = OpenXmlPowerTools.HtmlConverter.ConvertToHtml(document, settings);
                        XDocument html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);

                        return html.ToString(SaveOptions.DisableFormatting);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

        }
        
    }
}
