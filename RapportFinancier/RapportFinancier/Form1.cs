using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using RapportFinancier.Models;
using System.Diagnostics;

namespace RapportFinancier
{
    public partial class Form1 : Form
    {
        #region Variables
        private const int FONT_SIZE_10 = 10;
        private const int FONT_SIZE_12 = 12;
        private const int FONT_SIZE_14 = 14;
        private const int FONT_SIZE_16 = 16;
        private const int FONT_SIZE_20 = 20;

        private string directoryPath = string.Empty;
        private DateTime endOfTheYear = new DateTime(DateTime.Now.Year, 6, 30);
        private PdfDocument document;
        #endregion
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            GlobalFontSettings.FontResolver = new MyFontResolver();
        }
        #region Data simplificator
        private bool ValidateData()
        {
            if (!string.IsNullOrEmpty(textBoxCompanyName.Text) || !string.IsNullOrWhiteSpace(textBoxCompanyName.Text))
            {
                textBoxCompanyName.Text = textBoxCompanyName.Text.ToUpper();
                return true;
            }
            return false;
        }
        private void DrawHeaderFooter(XGraphics gfx, int pageNumber)
        {
            // Set up the fonts
            XFont footerFont = new("Roboto", FONT_SIZE_10, XFontStyleEx.BoldItalic);
            // Reuse the header and footer drawing methods
            DrawHeader(gfx);
            DrawFooter(gfx, footerFont, pageNumber);
        }
        private void DrawHeader(XGraphics gfx)
        {
            XImage logo = XImage.FromFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "aslen.png"));
            gfx.DrawImage(logo, 20, 20, logo.PixelWidth / 2, logo.PixelHeight / 2);
        }

        private void DrawFooter(XGraphics gfx, XFont footerFont, int pageNumber)
        {
            string footerLine1 = "ASLEN STRATÉGIES D’AFFAIRES";
            string footerLine2 = "6185 BOULEVARD TASCHEREAU SUITE 204 – BROSSARD, QC, J4Z1A6 450-904-1948";

            // Dessiner les lignes du footer
            double footerY = gfx.PageSize.Height - 100;
            gfx.DrawString(footerLine1, footerFont, XBrushes.Black, new XRect(0, footerY, gfx.PageSize.Width, 0), XStringFormats.TopCenter);
            gfx.DrawString(footerLine2, footerFont, XBrushes.Black, new XRect(0, footerY + 15, gfx.PageSize.Width, 0), XStringFormats.TopCenter);

            // Décaler la position Y pour le numéro de page
            double pageNumberY = footerY + 30; // Augmenter la valeur pour décaler le numéro de page vers le bas
            string pageNumberText = "Page " + pageNumber.ToString();
            gfx.DrawString(pageNumberText, footerFont, XBrushes.Black, new XRect(0, pageNumberY, gfx.PageSize.Width - 30, 0), XStringFormats.TopRight);
        }


        #endregion
        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (ValidateData())
                {
                    // Set the initial directory to the last user path if available
                    if (!string.IsNullOrEmpty(directoryPath))
                    {
                        folderDialog.SelectedPath = directoryPath;
                    }

                    // Show the dialog and get the result
                    DialogResult result = folderDialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                    {
                        // Update the userPath with the newly selected path
                        directoryPath = folderDialog.SelectedPath;
                        // Show a MessageBox with the selected path
                        //MessageBox.Show("The selected directory is: " + userPath, "Directory Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //Starting to generate pages
                        document = new PdfDocument();
                        CreateFirstPage();
                        CreateSecondPage();

                        //Close document and save
                        CloseDocument();
                        MessageBox.Show("Le rapport financier a été généré " + directoryPath + "\\RapportFinancier_" + textBoxCompanyName.Text + ".docx", "Generation finie", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
        #region CreateFirstPage
        private void CreateFirstPage()
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            DrawHeaderFooter(gfx, 1);

            XFont defaultFont = new("Roboto", FONT_SIZE_14, XFontStyleEx.Bold);
            // Center the company name and financial year in the middle of the page
            string companyName = textBoxCompanyName.Text.ToUpper();
            string financialYear = "ETATS FINANCIERS AU " + endOfTheYear.ToString("dd-MM-yyyy");

            // Measure the height needed for the company name and financial year
            double middleSectionHeight = gfx.MeasureString(companyName, defaultFont).Height +
                                         gfx.MeasureString(financialYear, defaultFont).Height;

            // Draw the text at the center of the page
            double centerY = (page.Height - middleSectionHeight) / 2;
            gfx.DrawString(companyName, defaultFont, XBrushes.Black, new XRect(0, centerY, page.Width, 0), XStringFormats.TopCenter);
            gfx.DrawString(financialYear, defaultFont, XBrushes.Black, new XRect(0, centerY + 40, page.Width, 0), XStringFormats.TopCenter);
        }
        #endregion
        #region CreateSecondPage
        private void CreateSecondPage()
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Reuse the header and footer from the first page
            DrawHeaderFooter(gfx, 2);

            // Define fonts for the page
            XFont companyNameFont = new ("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont dateFont = new ("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont sectionHeaderFont = new ("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont contentFont = new ("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);

            // Draw the company name at the top
            gfx.DrawString(textBoxCompanyName.Text.ToUpper(), companyNameFont, XBrushes.Black,
                new XRect(0, 100, page.Width, 0), XStringFormats.TopCenter);

            // Draw the date right below the company name
            gfx.DrawString("AU " + endOfTheYear.ToString("dd-MM-yyyy"), dateFont, XBrushes.Black, 
                new XRect(0, 120, page.Width, 0), XStringFormats.TopCenter);

            // Draw the section title "TABLE DES MATIÈRES"
            gfx.DrawString("TABLE DES MATIÈRES", sectionHeaderFont, XBrushes.Black,
                new XRect(0, 160, page.Width, 0), XStringFormats.TopCenter);

            // Set the initial Y position after the "TABLE DES MATIÈRES" title
            double startY = 350;
            double lineSpacing = 30;

            // The X position for the page numbers; adjust as needed for alignment
            double pageNumbersXPosition = page.Width - 150;

            // Draw the word "Page" above the page numbers list
            gfx.DrawString("Page", contentFont, XBrushes.Black,
                new XRect(pageNumbersXPosition, startY - lineSpacing, page.Width, 0), XStringFormats.TopLeft);

            // Define the table of contents entries with their page numbers
            var contents = new[]
            {
                new { Title = "- NOTE INTRODUCTIVE", Page = "3" },
                new { Title = "- BILAN – TABLEAU DES RÉSULTATS", Page = "4-6" },
                new { Title = "- NOTES COMPLÉMENTAIRES", Page = "7" }
            };

            foreach (var content in contents)
            {
                gfx.DrawString(content.Title, contentFont, XBrushes.Black,
                    new XRect(80, startY, page.Width, 0), XStringFormats.TopLeft);

                gfx.DrawString(content.Page, contentFont, XBrushes.Black,
                    new XRect(pageNumbersXPosition, startY, page.Width, 0), XStringFormats.TopLeft);

                startY += lineSpacing;
            }
        }
        #endregion
        private void CloseDocument()
        {
            // Save the document only after all pages have been added
            string filename = Path.Combine(directoryPath, "RapportFinancier_" + textBoxCompanyName.Text + "_.pdf");
           if (File.Exists(filename)) 
                File.Delete(filename);
            document.Save(filename);
            document.Close();
            document.Dispose();
            // Optionally open the PDF after creation
            //Process.Start(filename);
        }
    }
}
