using PdfSharp.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RapportFinancier.Models
{
    public class MyFontResolver : IFontResolver
    {
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // Here you map the family name, bold, and italic to the filenames of your fonts
            var fontName = familyName.ToLower() switch
            {
                "roboto" when isBold && isItalic => "Roboto-BoldItalic.ttf",
                "roboto" when isBold => "Roboto-Bold.ttf",
                "roboto" => "Roboto-Regular.ttf",
                _ => "Roboto-Regular.ttf", // Default to regular Roboto if not found
            };

            // Return a FontResolverInfo with the font's file name without extension
            return new FontResolverInfo(fontName);
        }

        public byte[] GetFont(string faceName)
        {
            // You need to provide the actual font data here.
            // For example, you can read the font file from the embedded resources:
            var assembly = Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream($"RapportFinancier.Fonts.{faceName}");

            if (stream == null)
                throw new FileNotFoundException($"Font file '{faceName}' not found in resources.");

            using var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }
    }

}
