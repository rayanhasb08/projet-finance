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
        private string directoryPath = string.Empty;
        private DateTime endOfTheYear = new DateTime(DateTime.Now.Year, 6, 30);
        private PdfDocument document = new PdfDocument();
        #endregion
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            GlobalFontSettings.FontResolver = new MyFontResolver();
        }
        private bool ValidateData()
        {
            if (!string.IsNullOrEmpty(textBoxCompanyName.Text) || !string.IsNullOrWhiteSpace(textBoxCompanyName.Text))
            {
                textBoxCompanyName.Text = textBoxCompanyName.Text.ToUpper();
                return true;
            }
            return false;
        }
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
                        CreateFirstPage();
                        CloseDocument();
                        MessageBox.Show("Le rapport financier a été généré " + directoryPath + "\\RapportFinancier_" + textBoxCompanyName.Text + ".docx", "Generation finie", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
        private void CreateFirstPage()
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tf = new XTextFormatter(gfx);

            // Set up the fonts
            XFont titleFont = new XFont("Roboto", 20, XFontStyleEx.Bold);
            XFont defaultFont = new XFont("Roboto", 12, XFontStyleEx.Bold);
            XFont footerFont = new XFont("Roboto", 10, XFontStyleEx.BoldItalic);

            // Load your image for the header
            XImage logo = XImage.FromFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "aslen.png"));

            // Draw the image at the top left corner of the page
            gfx.DrawImage(logo, 20, 20, logo.PixelWidth / 2, logo.PixelHeight / 2); // Adjust scaling as needed

            // Center the company name and financial year in the middle of the page
            string companyName = textBoxCompanyName.Text.ToUpper();
            string financialYear = "ETATS FINANCIERS AU " + endOfTheYear.ToString("dd-MM-yyyy");

            // Measure the height needed for the company name and financial year
            double middleSectionHeight = gfx.MeasureString(companyName, titleFont).Height +
                                         gfx.MeasureString(financialYear, defaultFont).Height;

            // Draw the text at the center of the page
            double centerY = (page.Height - middleSectionHeight) / 2;
            gfx.DrawString(companyName, titleFont, XBrushes.Black, new XRect(0, centerY, page.Width, 0), XStringFormats.TopCenter);
            gfx.DrawString(financialYear, defaultFont, XBrushes.Black, new XRect(0, centerY + 30, page.Width, 0), XStringFormats.TopCenter);

            // Add footer content at the bottom of the page
            string footerLine1 = "ASLEN STRATÉGIES D’AFFAIRES";
            string footerLine2 = "6185 BOULEVARD TASCHEREAU SUITE 204 – BROSSARD, QC, J4Z1A6 450-904-1948";

            // Measure the height needed for the footer
            double footerSectionHeight = gfx.MeasureString(footerLine1, footerFont).Height +
                                         gfx.MeasureString(footerLine2, footerFont).Height;

            // Draw the footer text at the bottom center of the page
            double footerY = page.Height - footerSectionHeight - 30; // Adjust the Y offset as needed for the footer
            gfx.DrawString(footerLine1, footerFont, XBrushes.Black, new XRect(0, footerY, page.Width, 0), XStringFormats.TopCenter);
            gfx.DrawString(footerLine2, footerFont, XBrushes.Black, new XRect(0, footerY + 15, page.Width, 0), XStringFormats.TopCenter);
        }

        private void CloseDocument()
        {
            // Save the document only after all pages have been added
            string filename = Path.Combine(directoryPath, "RapportFinancier_" + textBoxCompanyName.Text + "_.pdf");
            if (File.Exists(filename)) 
                File.Delete(filename);
            document.Save(filename);
            document.Close();

            // Optionally open the PDF after creation
            //Process.Start(filename);
        }
    }
}
