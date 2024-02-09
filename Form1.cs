using System;
using System.Drawing;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        private Panel panel;
        private int previousSelectionStart = 0;
        public Form1()
        {
            InitializeComponents();
            LoadFonts();
            Load += Form1_Load;
            Resize += Form1_Resize;
        }

        private void InitializeComponents()
        {


            Text = "Взаимодействие со шрифтом";
            Width = 500;
            Height = 300;


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

            panel = new Panel();
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

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
            {
                if (Height > 500)
                {
                    panel.Padding = new Padding(50, panel.Padding.Top, 50, 50);
                }
                else if (Width > 1000)
                {
                    panel.Padding = new Padding(50, panel.Padding.Top, 50, panel.Padding.Bottom);
                }
                else
                {
                    panel.Padding = new Padding(10, panel.Padding.Top, 10, 10);
                }
            }
            else if (WindowState == FormWindowState.Maximized)
            {
                panel.Padding = new Padding(50, panel.Padding.Top, 50, 50);
            }
            else
            {
                panel.Padding = new Padding(10, panel.Padding.Top, 10, 10);
            }
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
                        inputTextBox.Clear();
                        foreach (Paragraph paragraph in body.Elements<Paragraph>())
                        {
                            foreach (Run run in paragraph.Elements<Run>())
                            {
                                RunProperties runProperties = run.GetFirstChild<RunProperties>();
                                Text textElement = run.GetFirstChild<Text>();
                                if (textElement == null)
                                {
                                    continue;

                                }
                                string text = textElement.Text;
                                inputTextBox.AppendText(text);
                                int startIndex = inputTextBox.Text.Length - text.Length;
                                int length = text.Length;
                                if (runProperties != null)
                                {
                                    string fontFamily = runProperties.GetFirstChild<RunFonts>()?.Ascii;
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
                                    string colorHex = runProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val;
                                    System.Drawing.Color fontColor = ParseFontColor(colorHex);
                                    bool isBold = runProperties.GetFirstChild<Bold>() != null;
                                    bool isItalic = runProperties.GetFirstChild<Italic>() != null;
                                    inputTextBox.Select(startIndex, length);
                                    inputTextBox.SelectionFont = new System.Drawing.Font(fontFamily ?? "Times New Romain", fontSize, (isBold ? FontStyle.Bold : FontStyle.Regular) | (isItalic ? FontStyle.Italic : FontStyle.Regular));
                                    inputTextBox.SelectionColor = fontColor;
                                }
                            }
                            inputTextBox.AppendText(Environment.NewLine);
                        }
                    }
                }
            }
        }

        private System.Drawing.Color ParseFontColor(string colorValue)
        {
            if (!string.IsNullOrEmpty(colorValue))
            {
                if (colorValue.StartsWith("#") && colorValue.Length == 7)
                {
                    return System.Drawing.ColorTranslator.FromHtml(colorValue);
                }
                else
                {   
                    if(colorValue.Length == 6)
                    {
                        int hexValue;
                        if (int.TryParse(colorValue, System.Globalization.NumberStyles.HexNumber, null, out hexValue))
                        {
                            colorValue = "#" + colorValue;
                            return System.Drawing.ColorTranslator.FromHtml(colorValue);
                        }
                    }
                    try
                    {
                        return System.Drawing.Color.FromName(colorValue);
                    }
                    catch
                    {
                        return System.Drawing.Color.Black;
                    }
                }
            }
            return System.Drawing.Color.Black;
        }
    }
}
