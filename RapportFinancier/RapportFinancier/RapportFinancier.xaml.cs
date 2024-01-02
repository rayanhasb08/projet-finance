using Newtonsoft.Json;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using RapportFinancier.Models;
using System.Globalization;
using System.IO;
using System.Windows;
using Path = System.IO.Path;

namespace RapportFinancier
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class RappotFinancier : Window
    {
        #region Variables
        private const int FONT_SIZE_10 = 10;
        private const int FONT_SIZE_12 = 12;
        private const int FONT_SIZE_14 = 14;
        private const int FONT_SIZE_16 = 16;
        private const int FONT_SIZE_20 = 20;
        private const string FILENAME = "settings.json"; // Nom du fichier JSON
        private const string DIRECTORYNAME = "Sauvegardes"; // Nom du répertoire

        private readonly string FilePath;
        private readonly string LogFilePath;

        private string directoryPath = string.Empty;
        private readonly DateTime endOfTheYear = DateTime.Now.Month <= 6 ? new DateTime(DateTime.Now.Year, 6, 30) : new DateTime(DateTime.Now.Year + 1, 6, 30);
        private readonly DateTime todayDate = DateTime.Now;
        private PdfDocument document;
        #endregion
        public RappotFinancier()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            GlobalFontSettings.FontResolver = new MyFontResolver();
            CultureInfo culture = new CultureInfo("fr-FR");
            textBoxTodayDate.Text = todayDate.ToString("Le dd MMMM, yyyy", culture);
            // Obtenez le chemin complet du répertoire
            string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DIRECTORYNAME);
            // Obtenez le chemin complet du fichier
            FilePath = Path.Combine(folderPath, FILENAME);
            LogFilePath = Path.Combine(folderPath, "log.txt");
            LoadSettingsFromFile();
        }
        #region LoadSettingsFromFile
        private void LoadSettingsFromFile()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    // Lisez le contenu du fichier JSON
                    string jsonData = File.ReadAllText(FilePath);

                    // Désérialisez les données JSON en un objet
                    var settings = JsonConvert.DeserializeObject<dynamic>(jsonData);

                    // Mettez les données dans les boîtes de texte
                    textBoxCompanyName.Text = settings.CompanyName;
                    textBoxAccountantName.Text = settings.AccountantName;
                    textBoxAutorisedSignatory.Text = settings.AutorisedSignatory;
                    // Remarque : Assurez-vous d'utiliser les mêmes noms de propriétés que ceux utilisés lors de la sérialisation
                }
                else
                {
                    // Si le fichier n'existe pas, ne faites rien ici
                }
            }
            catch (Exception ex)
            {
                // Enregistrez l'erreur dans le fichier de log
                LogError(ex.Message);
            }
        }
        #endregion
        #region Data simplificator
        private bool ValidateData()
        {
            if ((!string.IsNullOrEmpty(textBoxCompanyName.Text) || !string.IsNullOrWhiteSpace(textBoxCompanyName.Text)) &&
                (!string.IsNullOrEmpty(textBoxAccountantName.Text) || !string.IsNullOrWhiteSpace(textBoxAccountantName.Text)) &&
                (!string.IsNullOrEmpty(textBoxAutorisedSignatory.Text) || !string.IsNullOrWhiteSpace(textBoxAutorisedSignatory.Text)))
            {
                textBoxCompanyName.Text = textBoxCompanyName.Text.ToUpper();
                textBoxAutorisedSignatory.Text = ToPascalCaseWithSpaces(textBoxAutorisedSignatory.Text);
                return true;
            }
            return false;
        }
        public static string ToPascalCaseWithSpaces(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            string[] words = input.Split(new char[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                string word = words[i];
                if (word.Length > 0)
                {
                    words[i] = char.ToUpper(word[0]) + word.Substring(1).ToLower();
                }
            }

            return string.Join(" ", words);
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
        #region buttonGenerate
        private void buttonGenerate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Using OpenFileDialog from Microsoft.Win32 to select folders
                var folderDialog = new Microsoft.Win32.OpenFileDialog
                {
                    ValidateNames = false,
                    CheckFileExists = false,
                    CheckPathExists = true,
                    FileName = "Folder Selection."
                };

                if (ValidateData())
                {
                    // Set the initial directory to the last user path if available
                    if (!string.IsNullOrEmpty(directoryPath))
                    {
                        folderDialog.InitialDirectory = directoryPath;
                    }

                    // Show the dialog and get the result
                    bool? result = folderDialog.ShowDialog();

                    if (result == true && !string.IsNullOrWhiteSpace(folderDialog.FileName))
                    {
                        // Update the directoryPath with the newly selected path
                        directoryPath = Path.GetDirectoryName(folderDialog.FileName);

                        // Starting to generate pages
                        document = new PdfDocument();
                        CreateFirstPage();
                        CreateSecondPage();
                        CreateThirdPage();

                        // Close document and save
                        CloseDocument();
                        MessageBox.Show("Le rapport financier a été généré " + directoryPath + "\\RapportFinancier_" + textBoxCompanyName.Text + ".docx", "Generation finie", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                    MessageBox.Show("Veuillez remplir tout les champs", "Champs vides", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                // Log any unexpected exceptions
                LogError("Entreprise : " + textBoxCompanyName.Text + " Une erreur inattendue s'est produite : " + ex.Message);
            }
        }
        #endregion
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
            XFont companyNameFont = new("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont dateFont = new("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont sectionHeaderFont = new("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);
            XFont contentFont = new("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);

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
        #region CreateThirdPage
        private void CreateThirdPage()
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Header and footer
            DrawHeaderFooter(gfx, 3); // Pass the correct page number

            // Fonts
            XFont regularFont = new XFont("Roboto", FONT_SIZE_12, XFontStyleEx.Regular);
            XFont boldFont = new XFont("Roboto", FONT_SIZE_12, XFontStyleEx.Bold);

            // Start Y position for the text body
            double startY = 150; // Adjust based on your header size

            // Notice Header
            gfx.DrawString("AVIS AU LECTEUR", boldFont, XBrushes.Black, new XRect(0, startY, page.Width, 0), XStringFormats.TopCenter);

            // Increment Y position after the header
            startY += 40; // Adjust this value as needed based on your content

            // First Paragraph - dynamically insert company name and formatted date
            string firstParagraph = $"J’ai préparé le bilan de la compagnie {textBoxCompanyName.Text.ToUpper()} au " +
                                    $"{endOfTheYear.ToString("dd MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR"))} " +
                                    "d’après les livres de l’entreprise et suivant les renseignements obtenus auprès de son Président.";
            gfx.DrawString(firstParagraph, regularFont, XBrushes.Black, new XRect(40, startY, page.Width - 80, 0), XStringFormats.TopLeft);

            // Accountant Name and Signature Line
            startY += 100; // Adjust as needed
            gfx.DrawString(textBoxAccountantName.Text, boldFont, XBrushes.Black, new XRect(0, startY, page.Width, 0), XStringFormats.TopCenter);
            gfx.DrawLine(XPens.Black, 200, startY + 20, page.Width - 200, startY + 20); // Signature line

            // Today's Date
            startY += 30; // Adjust as needed
            gfx.DrawString(textBoxTodayDate.Text, regularFont, XBrushes.Black, new XRect(0, startY, page.Width, 0), XStringFormats.TopCenter);

            // "Lu et Approuvé" Text and Date
            startY += 60; // Adjust as needed
            gfx.DrawString("Lu et Approuvé Le " + textBoxTodayDate.Text, regularFont, XBrushes.Black, new XRect(40, startY, page.Width - 80, 0), XStringFormats.TopLeft);

            // Authorized Signatory Name and Signature Line
            startY += 20; // Adjust as needed
            gfx.DrawString(textBoxAutorisedSignatory.Text, boldFont, XBrushes.Black, new XRect(0, startY, page.Width, 0), XStringFormats.TopCenter);
            gfx.DrawLine(XPens.Black, 200, startY + 20, page.Width - 200, startY + 20); // Signature line
        }


        #endregion
        #region CloseDocument
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
        #endregion
        #region RapportFinancier_Closing & Log
        private void RapportFinancier_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveSettingsToFile();
        }
        private void SaveSettingsToFile()
        {
            string company = textBoxCompanyName.Text;
            string accountant = textBoxAccountantName.Text;
            string autorisedSignatory = textBoxAutorisedSignatory.Text;

            // Créez un objet contenant tous les champs
            var settings = new
            {
                CompanyName = company,
                AccountantName = accountant,
                AutorisedSignatory = autorisedSignatory
                // Ajoutez d'autres champs ici
            };

            try
            {
                // Créez le répertoire s'il n'existe pas
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));

                // Sérialisez les données en JSON et enregistrez-les dans le fichier
                string jsonData = JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(FilePath, jsonData);
            }
            catch (Exception ex)
            {
                // Enregistrez l'erreur dans le fichier de log
                LogError(ex.Message);
            }
        }
        private void LogError(string errorMessage)
        {
            try
            {
                // Écrivez l'erreur dans le fichier de log
                File.AppendAllText(LogFilePath, $"{DateTime.Now} - Erreur : {errorMessage}\n");
            }
            catch
            {
                // En cas d'échec de l'enregistrement de l'erreur dans le fichier de log, ne faites rien ici pour éviter une boucle infinie d'erreurs
            }
        }
        #endregion
    }
}