using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace DocxLetterGenerator
{
    public class AppendixItem
    {
        public int Number { get; set; }
        public string Title { get; set; }
        public string Pages { get; set; }
        public string Text { get; set; }
    }

    public partial class MainWindow : Window
    {
        private List<AppendixItem> appendices = new List<AppendixItem>();

        public MainWindow()
        {
            InitializeComponent();

            RecipientNameBox.TextChanged += (s, e) => UpdatePreview();
            RecipientPostBox.TextChanged += (s, e) => UpdatePreview();
            SubjectBox.TextChanged += (s, e) => UpdatePreview();
            BodyBox.TextChanged += (s, e) => UpdatePreview();
            Corporation.TextChanged += (s, e) => UpdatePreview();
            LLC.TextChanged += (s, e) => UpdatePreview();
            Telephone.TextChanged += (s, e) => UpdatePreview();
            Fax.TextChanged += (s, e) => UpdatePreview();
            Email.TextChanged += (s, e) => UpdatePreview();
            OKPO.TextChanged += (s, e) => UpdatePreview();
            OGRN.TextChanged += (s, e) => UpdatePreview();
            INN.TextChanged += (s, e) => UpdatePreview();
            KPP.TextChanged += (s, e) => UpdatePreview();
            RecipientCorp.TextChanged += (s, e) => UpdatePreview();
            SenderPosition.TextChanged += (s, e) => UpdatePreview();
            SenderName.TextChanged += (s, e) => UpdatePreview();

            UpdatePreview();
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusText.Text = "Генерация документа...";
                StatusText.Foreground = Brushes.Blue;

                if (string.IsNullOrWhiteSpace(Corporation?.Text))
                {
                    MessageBox.Show("Введите корпорацию/организацию отправителя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Corporation.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(LLC?.Text))
                {
                    MessageBox.Show("Введите ООО/Юр. лицо отправителя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    LLC.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(Telephone?.Text))
                {
                    MessageBox.Show("Введите номер телефона!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Telephone.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(Email?.Text))
                {
                    MessageBox.Show("Введите Email!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Email.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(OKPO?.Text))
                {
                    MessageBox.Show("Введите ОКПО!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    OKPO.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(OGRN?.Text))
                {
                    MessageBox.Show("Введите ОГРН!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    OGRN.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(INN?.Text))
                {
                    MessageBox.Show("Введите ИНН!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    INN.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(KPP?.Text))
                {
                    MessageBox.Show("Введите КПП!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    KPP.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(RecipientCorp?.Text))
                {
                    MessageBox.Show("Введите корпорацию получателя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    RecipientCorp.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(RecipientPostBox.Text))
                {
                    MessageBox.Show("Введите должность получателя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    RecipientPostBox.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(RecipientNameBox.Text))
                {
                    MessageBox.Show("Введите имя получателя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    RecipientNameBox.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(SubjectBox.Text))
                {
                    MessageBox.Show("Введите тему письма!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    SubjectBox.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(BodyBox.Text))
                {
                    if (MessageBox.Show("Текст письма пуст. Продолжить генерацию?", "Предупреждение",
                        MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    {
                        BodyBox.Focus();
                        return;
                    }
                }

                if (string.IsNullOrWhiteSpace(SenderPosition?.Text))
                {
                    MessageBox.Show("Введите должность отправителя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    SenderPosition.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(SenderName?.Text))
                {
                    MessageBox.Show("Введите имя отправителя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    SenderName.Focus();
                    return;
                }

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "template.docx");
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string outputPath = Path.Combine(desktopPath, $"letter_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"Шаблон не найден:\n{templatePath}\n\nПоложите template.docx в папку Templates", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "Ошибка: шаблон не найден";
                    StatusText.Foreground = Brushes.Red;
                    return;
                }

                File.Copy(templatePath, outputPath, true);

                var replacements = new Dictionary<string, string>
                {
                    ["{ CORPORATION }"] = Corporation?.Text ?? "",
                    ["({ LLC })"] = $"({LLC?.Text ?? ""})",
                    ["{ TELEPHONE }"] = Telephone?.Text ?? "",
                    ["{ FAX }"] = Fax?.Text ?? "",
                    ["{ EMAIL }"] = Email?.Text ?? "",
                    ["{ OKPO }"] = OKPO?.Text ?? "",
                    ["{ OGRN }"] = OGRN?.Text ?? "",
                    ["{ INN }"] = INN?.Text ?? "",
                    ["{ KPP }"] = KPP?.Text ?? "",
                    ["{ RECIPIENT_POST }"] = RecipientPostBox.Text,
                    ["{ RECIPIENT_NAME }"] = RecipientNameBox.Text,
                    ["{ RECIPIENT_CORP }"] = RecipientCorp?.Text ?? "",
                    ["{ SUBJECT }"] = SubjectBox.Text,
                    ["{ CONTENT }"] = BodyBox.Text,
                    ["{ DATE }"] = DateTime.Now.ToString("dd.MM.yyyy"),
                    ["{ SENDER_POSITION }"] = SenderPosition?.Text ?? "",
                    ["{ SENDER_NAME }"] = SenderName?.Text ?? ""
                };

                DocxConverter.ReplacePlaceholders(outputPath, replacements);

                if (appendices.Count > 0)
                {
                    var listBlocks = DocxConverter.BuildAppendixListBlocks(appendices);
                    DocxConverter.ReplacePlaceholderWithBlocks(outputPath, "{ APPENDIX_LIST }", listBlocks);

                    var fullBlocks = DocxConverter.BuildAppendixBlocks(appendices);
                    DocxConverter.ReplacePlaceholderWithBlocks(outputPath, "{ APPENDIX_BLOCKS }", fullBlocks);
                }
                else
                {
                    DocxConverter.RemovePlaceholder(outputPath, "{ APPENDIX_LIST }");
                    DocxConverter.RemovePlaceholder(outputPath, "{ APPENDIX_BLOCKS }");
                }

                StatusText.Text = $"Документ создан: {outputPath}";
                StatusText.Foreground = Brushes.Green;

                if (File.Exists(outputPath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "Ошибка генерации";
                StatusText.Foreground = Brushes.Red;
            }
        }
        private void AddAppendixButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(AppendixTitleBox.Text))
            {
                MessageBox.Show("Введите заголовок приложения!", "Ошибка");
                return;
            }

            appendices.Add(new AppendixItem
            {
                Number = appendices.Count + 1,
                Title = AppendixTitleBox.Text,
                Pages = string.IsNullOrWhiteSpace(AppendixPagesBox.Text) ? "" : AppendixPagesBox.Text,
                Text = AppendixTextBox.Text ?? ""
            });

            AppendixTitleBox.Text = "";
            AppendixPagesBox.Text = "";
            AppendixTextBox.Text = "";

            RefreshAppendicesGrid();
            UpdatePreview();
        }

        private void RemoveAppendixButton_Click(object sender, RoutedEventArgs e)
        {
            if (AppendicesGrid.SelectedItem is AppendixItem selected)
            {
                appendices.Remove(selected);
                for (int i = 0; i < appendices.Count; i++)
                    appendices[i].Number = i + 1;

                RefreshAppendicesGrid();
                UpdatePreview();
            }
            else
            {
                MessageBox.Show("Выберите приложение для удаления!", "Ошибка");
            }
        }

        private void RefreshAppendicesGrid()
        {
            AppendicesGrid.ItemsSource = null;
            AppendicesGrid.ItemsSource = appendices;
        }

        private void UpdatePreview()
        {
            try
            {
                string tempPreviewPath = Path.Combine(Path.GetTempPath(), "preview_temp.docx");
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "template.docx");

                if (!File.Exists(templatePath)) return;

                File.Copy(templatePath, tempPreviewPath, true);

                var replacements = new Dictionary<string, string>
                {
                    ["{ CORPORATION }"] = Corporation?.Text ?? "",
                    ["({ LLC })"] = $"({LLC?.Text ?? ""})",
                    ["{ TELEPHONE }"] = Telephone?.Text ?? "",
                    ["{ FAX }"] = Fax?.Text ?? "",
                    ["{ EMAIL }"] = Email?.Text ?? "",
                    ["{ OKPO }"] = OKPO?.Text ?? "",
                    ["{ OGRN }"] = OGRN?.Text ?? "",
                    ["{ INN }"] = INN?.Text ?? "",
                    ["{ KPP }"] = KPP?.Text ?? "",
                    ["{ RECIPIENT_POST }"] = RecipientPostBox.Text,
                    ["{ RECIPIENT_NAME }"] = RecipientNameBox.Text,
                    ["{ RECIPIENT_CORP }"] = RecipientCorp?.Text ?? "",
                    ["{ SUBJECT }"] = SubjectBox.Text,
                    ["{ CONTENT }"] = BodyBox.Text,
                    ["{ DATE }"] = DateTime.Now.ToString("dd.MM.yyyy"),
                    ["{ SENDER_POSITION }"] = SenderPosition?.Text ?? "",
                    ["{ SENDER_NAME }"] = SenderName?.Text ?? ""
                };

                DocxConverter.ReplacePlaceholders(tempPreviewPath, replacements);

                if (appendices.Count > 0)
                {
                    var listBlocks = DocxConverter.BuildAppendixListBlocks(appendices);
                    DocxConverter.ReplacePlaceholderWithBlocks(tempPreviewPath, "{ APPENDIX_LIST }", listBlocks);

                    var fullBlocks = DocxConverter.BuildAppendixBlocks(appendices);
                    DocxConverter.ReplacePlaceholderWithBlocks(tempPreviewPath, "{ APPENDIX_BLOCKS }", fullBlocks);
                }
                else
                {
                    DocxConverter.RemovePlaceholder(tempPreviewPath, "{ APPENDIX_LIST }");
                    DocxConverter.RemovePlaceholder(tempPreviewPath, "{ APPENDIX_BLOCKS }");
                }

                DocPreview.Load(tempPreviewPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Preview error: {ex.Message}");
            }
        }
    }
}
