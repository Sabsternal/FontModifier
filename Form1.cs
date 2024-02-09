using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace FontModifier
{
    public partial class Form1 : Form
    {
        

        private ComboBox fontComboBox;
        private NumericUpDown sizeNumericUpDown;
        private Button colorButton;
        private System.Windows.Forms.CheckBox boldCheckBox;
        private System.Windows.Forms.CheckBox italicCheckBox;
        private RichTextBox inputTextBox;
        private Button saveButton;
        private Button loadButton;
        private SettingsController MainSettings;
        private int previousSelectionStart = 0;
        public Form1()
        {
            InitializeComponents();
            LoadFonts();
            Load += Form1_Load;
        }

        private void InitializeComponents()
        {


            Text = "Взаимодействие со шрифтом";
            Width = 500;
            Height = 500;


            MainSettings = new SettingsController();

            fontComboBox = new ComboBox();
            fontComboBox.Location = new Point(10, 10);
            fontComboBox.Width = 200;
            fontComboBox.SelectedIndexChanged += FontComboBox_SelectedIndexChanged;
            Controls.Add(fontComboBox);

            sizeNumericUpDown = new NumericUpDown();
            sizeNumericUpDown.Location = new Point(220, 10);
            sizeNumericUpDown.Width = 60;
            sizeNumericUpDown.Minimum = 1;
            sizeNumericUpDown.ValueChanged += SizeNumericUpDown_ValueChanged;

            colorButton = new Button();
            colorButton.Location = new Point(290, 10);
            colorButton.Width = 40;
            colorButton.Height = 40;
            colorButton.Click += ColorButton_Click;

            boldCheckBox = new System.Windows.Forms.CheckBox();
            boldCheckBox.Location = new Point(10, 40);
            boldCheckBox.Text = "Жирный";
            boldCheckBox.CheckedChanged += BoldCheckBox_CheckedChanged;

            italicCheckBox = new System.Windows.Forms.CheckBox();
            italicCheckBox.Location = new Point(150, 40);
            italicCheckBox.Text = "Косой";
            italicCheckBox.CheckedChanged += ItalicCheckBox_CheckedChanged;

            saveButton = new Button();
            saveButton.Location = new Point(340, 10);
            saveButton.AutoSize = true;
            saveButton.AutoEllipsis = true;
            saveButton.Text = "Сохранить текст";
            saveButton.Click += SaveButton_Click;
            Controls.Add(saveButton);

            loadButton = new Button();
            loadButton.Location = new Point(saveButton.Location.X, saveButton.Location.Y + saveButton.Height + 5);
            loadButton.AutoSize = true;
            loadButton.AutoEllipsis = true;
            loadButton.Width = saveButton.Width;
            loadButton.Text = "Загрузить текст";
            loadButton.Click += LoadButton_Click;
            Controls.Add(loadButton);

            Controls.Add(sizeNumericUpDown);
            Controls.Add(colorButton);
            Controls.Add(boldCheckBox);
            Controls.Add(italicCheckBox);

            Panel panel = new Panel();
            panel.Location = new Point(0, 70);
            panel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Top;
            panel.Padding = new Padding(10, 100, 10, 10);
            panel.Dock = DockStyle.Fill;
            panel.AutoScroll = true;
            panel.HorizontalScroll.Enabled = false;
            panel.HorizontalScroll.Visible = false;

            inputTextBox = new RichTextBox();
            inputTextBox.Dock = DockStyle.Fill;
            inputTextBox.Multiline = true;
            inputTextBox.SelectionChanged += InputTextBox_SelectionChanged;
            panel.Controls.Add(inputTextBox);
            Controls.Add(panel);

            

        }

        private void LoadFonts()
        {

            foreach (System.Drawing.FontFamily fontFamily in System.Drawing.FontFamily.Families)
            {
                fontComboBox.Items.Add(fontFamily.Name);
            }

            if (fontComboBox.Items.Count > 0)
            {
                fontComboBox.SelectedIndex = 0;
            }
        }

        private void UpdatePreview(String Message)
        {
            string selectedText = inputTextBox.SelectedText;
            int selectionStart = inputTextBox.SelectionStart;
            int selectionLength = inputTextBox.SelectionLength;
            MainSettings.UpdateSettings(boldCheckBox.Checked, italicCheckBox.Checked, fontComboBox.SelectedIndex, colorButton.BackColor, (float)sizeNumericUpDown.Value);
            if (MainSettings.SettingsChanged())
            {
                switch (Message)
                {
                    case "Form1_Load":
                        { 
                            inputTextBox.SelectionFont = new System.Drawing.Font(
                                fontComboBox.SelectedItem.ToString(),
                                (float)sizeNumericUpDown.Value,
                                (boldCheckBox.Checked ? FontStyle.Bold : FontStyle.Regular) | (italicCheckBox.Checked ? FontStyle.Italic : FontStyle.Regular));
                            inputTextBox.SelectionColor = colorButton.BackColor;
                            break;
                        }
                    case "FontChanged":
                        {
                            inputTextBox.SelectionFont = new System.Drawing.Font(
                            fontComboBox.SelectedItem.ToString(),
                            inputTextBox.SelectionFont?.Size ?? inputTextBox.Font.Size,
                            inputTextBox.SelectionFont?.Style ?? FontStyle.Regular);
                            break;
                        }
                    case "SizeChanged":
                        {
                            if (inputTextBox.SelectionLength > 0)
                            {
                                int selectionEnd = selectionStart + inputTextBox.SelectionLength;

                                for (int i = selectionStart; i < selectionEnd; i++)
                                {
                                    inputTextBox.Select(i, 1);
                                    inputTextBox.SelectionFont = new System.Drawing.Font(
                                        inputTextBox.SelectionFont.FontFamily,
                                        (float)sizeNumericUpDown.Value,
                                        inputTextBox.SelectionFont.Style);
                                }

                                inputTextBox.Select(selectionStart, inputTextBox.SelectionLength);
                            }
                            break;
                        }
                    case "ColorChanged":
                        {
                            inputTextBox.SelectionColor = colorButton.BackColor;
                            break;
                        }
                    case "BoldChanged":
                        {
                            if (inputTextBox.SelectionLength > 0)
                            {
                                int selectionEnd = selectionStart + inputTextBox.SelectionLength;

                                for (int i = selectionStart; i < selectionEnd; i++)
                                {
                                    inputTextBox.Select(i, 1);

                                    FontStyle currentStyle = inputTextBox.SelectionFont?.Style ?? FontStyle.Regular;

                                    if (boldCheckBox.Checked)
                                    {
                                        currentStyle |= FontStyle.Bold;
                                    }
                                    else
                                    {
                                        currentStyle &= ~FontStyle.Bold;
                                    }

                                    System.Drawing.Font currentFont = inputTextBox.SelectionFont ?? inputTextBox.Font;
                                    System.Drawing.Font newFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, currentStyle);
                                    inputTextBox.SelectionFont = newFont;
                                }

                                inputTextBox.Select(selectionStart, selectionLength);
                            }
                            break;
                        }
                    case "ItalicChanged":
                        {
                            if (inputTextBox.SelectionLength > 0)
                            {
                                int selectionEnd = selectionStart + inputTextBox.SelectionLength;

                                for (int i = selectionStart; i < selectionEnd; i++)
                                {
                                    inputTextBox.Select(i, 1); 

                                    FontStyle currentStyle = inputTextBox.SelectionFont?.Style ?? FontStyle.Regular;

                                    if (italicCheckBox.Checked)
                                    {
                                        currentStyle |= FontStyle.Italic;
                                    }
                                    else
                                    {
                                        currentStyle &= ~FontStyle.Italic;
                                    }

                                    System.Drawing.Font currentFont = inputTextBox.SelectionFont ?? inputTextBox.Font;
                                    System.Drawing.Font newFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, currentStyle);
                                    inputTextBox.SelectionFont = newFont;
                                }

                                inputTextBox.Select(selectionStart, selectionLength);
                            }
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            inputTextBox.Select(selectionStart, selectionLength);
            inputTextBox.Focus();
        }


            private void FontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePreview("FontChanged");
        }

        private void SizeNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            UpdatePreview("SizeChanged");
        }

        private void InputTextBox_SelectionChanged(object sender, EventArgs e)
        {
            if (inputTextBox.SelectionLength > 0)
            {
                UpdatePreview("SelectionChanged");
            }
        }

        private void ColorButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                colorButton.BackColor = colorDialog.Color;
                UpdatePreview("ColorChanged");
            }
        }

        private void BoldCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePreview("BoldChanged");
        }

        private void ItalicCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePreview("ItalicChanged");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadFonts();
            if (fontComboBox.Items.Count > 0)
            {
                fontComboBox.SelectedIndex = fontComboBox.FindString("Times New Roman");
                sizeNumericUpDown.Value = 14;
                colorButton.BackColor = System.Drawing.Color.Black;
                boldCheckBox.Checked = false;
                italicCheckBox.Checked = false;
                MainSettings.SetSettings(boldCheckBox.Checked, italicCheckBox.Checked, fontComboBox.SelectedIndex, colorButton.BackColor, (float)sizeNumericUpDown.Value);
            }

            UpdatePreview("Form1_Load");

            inputTextBox.Focus();
            inputTextBox.Select();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = mainPart.Document.AppendChild(new Body());

                        Paragraph paragraph = body.AppendChild(new Paragraph());

                        string inputText = inputTextBox.Text;

                        for (int i = 0; i < inputText.Length; i++)
                        {
                            char c = inputText[i];
                            inputTextBox.Select(i, 1);
                            System.Drawing.Font font = inputTextBox.SelectionFont;
                            int fontSize = (int)font.Size;
                            bool isBold = font.Bold;
                            bool isItalic = font.Italic;
                            System.Drawing.Color textColor = inputTextBox.SelectionColor;

                            Run run = paragraph.AppendChild(new Run());
                            RunProperties runProperties = run.AppendChild(new RunProperties());
                            if (isBold)
                                runProperties.AppendChild(new Bold());
                            if (isItalic)
                                runProperties.AppendChild(new Italic());
                            runProperties.AppendChild(new FontSize() { Val = new StringValue(fontSize.ToString()) });
                            runProperties.AppendChild(new RunFonts() { Ascii = font.FontFamily.Name });
                            runProperties.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = ColorTranslator.ToHtml(textColor) });
                            if (c == ' ')
                            {
                                run.AppendChild(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
                            }
                            else if (c == '\t')
                            {
                                run.AppendChild(new Text("\t") { Space = SpaceProcessingModeValues.Preserve });
                            }
                            else if (c == '\n')
                            {
                                run.AppendChild(new Break());
                            }
                            else
                            {
                                run.AppendChild(new Text(c.ToString()));
                            }
                        }

                        wordDocument.Save();

                        MessageBox.Show("Результаты успешно сохранены");
                    }
                }
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Document (*.docx)|*.docx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
                    {
                        MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                        Body body = mainPart.Document.Body;

                        // Clear the existing text in the inputTextBox
                        inputTextBox.Clear();

                        // Iterate through the paragraphs in the Word document
                        foreach (Paragraph paragraph in body.Elements<Paragraph>())
                        {
                            // Iterate through the runs in the paragraph
                            foreach (Run run in paragraph.Elements<Run>())
                            {
                                // Get the font properties for the run
                                RunProperties runProperties = run.GetFirstChild<RunProperties>();

                                // Get the text of the run
                                Text textElement = run.GetFirstChild<Text>();
                                if (textElement == null)
                                {
                                    continue;

                                }
                                // Get the text of the run
                                string text = textElement.Text;

                                // Append the text to the inputTextBox with the corresponding formatting
                                inputTextBox.AppendText(text);

                                int startIndex = inputTextBox.Text.Length - text.Length;
                                int length = text.Length;

                                // Apply the formatting to the appended text
                                if (runProperties != null)
                                {
                                    // Get the font family
                                    string fontFamily = runProperties.GetFirstChild<RunFonts>()?.Ascii;

                                    // Get the font size
                                    int? fontSizeValue;
                                    if (int.TryParse(runProperties.GetFirstChild<FontSize>()?.Val, out int parsedFontSize))
                                    {
                                        fontSizeValue = parsedFontSize;
                                    }
                                    else
                                    {
                                        fontSizeValue = null;
                                    }

                                    int fontSize = fontSizeValue ?? 12;

                                    // Get the font color
                                    string colorHex = runProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val;
                                    System.Drawing.Color fontColor = ParseFontColor(colorHex);

                                    // Get the bold and italic styles
                                    bool isBold = runProperties.GetFirstChild<Bold>() != null;
                                    bool isItalic = runProperties.GetFirstChild<Italic>() != null;

                                    // Apply the formatting to the appended text
                                    inputTextBox.Select(startIndex, length);
                                    inputTextBox.SelectionFont = new System.Drawing.Font(fontFamily ?? "Arial", fontSize, (isBold ? FontStyle.Bold : FontStyle.Regular) | (isItalic ? FontStyle.Italic : FontStyle.Regular));
                                    inputTextBox.SelectionColor = fontColor;
                                }
                            }

                            // Append a new line after each paragraph
                            inputTextBox.AppendText(Environment.NewLine);
                        }
                    }
                }
            }
        }

        private System.Drawing.Color ParseFontColor(string colorHex)
        {
            if (!string.IsNullOrEmpty(colorHex) && colorHex.Length == 6)
            {
                // Prepend the '#' character to the color hex code
                colorHex = "#" + colorHex;

                // Convert the color hex code to a Color object
                return System.Drawing.ColorTranslator.FromHtml(colorHex);
            }

            // Return a default color if parsing fails
            return System.Drawing.Color.Black;
        }
    }
}
