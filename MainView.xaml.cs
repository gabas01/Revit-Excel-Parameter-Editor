using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using OfficeOpenXml; // Biblioteca EPPlus
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace RevitParameterEditor
{
    public partial class MainView : Window
    {
        private readonly UIApplication _uiapp;
        private readonly Document _doc;

        public MainView(UIApplication uiapp)
        {
            InitializeComponent();
            _uiapp = uiapp;
            _doc = uiapp.ActiveUIDocument.Document;

            // Define o contexto da licença para o EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Mude para Commercial se aplicável

            LoadSchedules();
        }

        // Carrega todas as tabelas do projeto no ComboBox
        private void LoadSchedules()
        {
            var schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate && s.Definition.GetFieldCount() > 0)
                .OrderBy(s => s.Name)
                .ToList();

            ScheduleComboBox.ItemsSource = schedules;
            ScheduleComboBox.DisplayMemberPath = "Name";
        }

        // Lógica do botão de EXPORTAR
        private async void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (ScheduleComboBox.SelectedItem == null)
            {
                StatusTextBlock.Text = "Erro: Nenhuma tabela selecionada.";
                TaskDialog.Show("Erro", "Por favor, selecione uma tabela da lista.");
                return;
            }

            var schedule = ScheduleComboBox.SelectedItem as ViewSchedule;

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Arquivo Excel (*.xlsx)|*.xlsx",
                FileName = $"{schedule.Name}.xlsx",
                Title = "Salvar Tabela"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                StatusTextBlock.Text = "Exportando dados...";
                try
                {
                    await ExportScheduleAsync(schedule, saveFileDialog.FileName);
                    StatusTextBlock.Text = "Exportação concluída com sucesso!";
                    TaskDialog.Show("Sucesso", $"A tabela '{schedule.Name}' foi exportada.");
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = "Erro na exportação.";
                    TaskDialog.Show("Erro de Exportação", ex.Message);
                }
            }
        }

        // Lógica do botão de IMPORTAR
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Arquivo Excel (*.xlsx)|*.xlsx",
                Title = "Selecionar Arquivo para Importação"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                StatusTextBlock.Text = "Importando e atualizando...";
                try
                {
                    int updatedCount = ImportParameters(openFileDialog.FileName);
                    StatusTextBlock.Text = $"Importação concluída. {updatedCount} elementos atualizados.";
                    TaskDialog.Show("Sucesso", $"{updatedCount} elementos foram atualizados no projeto Revit.");
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = "Erro na importação.";
                    TaskDialog.Show("Erro de Importação", $"Ocorreu um erro: {ex.Message}");
                }
            }
        }

        // Método assíncrono para exportar
        private System.Threading.Tasks.Task ExportScheduleAsync(ViewSchedule schedule, string filePath)
        {
            return System.Threading.Tasks.Task.Run(() =>
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Dados da Tabela");
                    var collector = new FilteredElementCollector(_doc, schedule.Id);

                    // Cabeçalhos
                    var fields = schedule.Definition.GetSchedulableFields();
                    var headers = new List<string> { "ElementID" }; // Coluna A será o ID
                    headers.AddRange(schedule.Definition.GetSchedulableFields().Select(f => f.GetName(_doc)));
                    
                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                    }

                    // Dados
                    int row = 2;
                    foreach (Element elem in collector)
                    {
                        worksheet.Cells[row, 1].Value = elem.Id.ToString(); // Adiciona o ID do elemento
                        for (int col = 0; col < schedule.Definition.GetFieldCount(); col++)
                        {
                            var field = schedule.Definition.GetField(col);
                            worksheet.Cells[row, col + 2].Value = elem.get_Parameter(field.ParameterId)?.AsValueString() ?? elem.get_Parameter(field.ParameterId)?.AsString();
                        }
                        row++;
                    }

                    worksheet.Cells.AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }
            });
        }

        // Método para importar
        private int ImportParameters(string filePath)
        {
            int updatedCount = 0;
            var fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null) throw new Exception("Nenhuma planilha encontrada no arquivo.");

                var headers = new List<string>();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

                int elementIdCol = headers.IndexOf("ElementID") + 1;
                if (elementIdCol == 0) throw new Exception("A coluna 'ElementID' não foi encontrada. Ela é essencial para a importação.");

                using (var trans = new Transaction(_doc, "Atualizar Parâmetros via Excel"))
                {
                    trans.Start();

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var idText = worksheet.Cells[row, elementIdCol].Text;
                        if (!long.TryParse(idText, out long idAsLong)) continue;
                        
                        var elementId = new ElementId(idAsLong);
                        var element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            bool elementUpdated = false;
                            for (int col = 1; col <= headers.Count; col++)
                            {
                                if (col == elementIdCol) continue; // Pula a coluna do ID

                                var paramName = headers[col - 1];
                                var parameter = element.LookupParameter(paramName);

                                if (parameter != null && !parameter.IsReadOnly)
                                {
                                    var newValue = worksheet.Cells[row, col].Text;
                                    // Compara o valor antigo com o novo para evitar transações desnecessárias
                                    string oldValue = parameter.AsValueString() ?? parameter.AsString();
                                    if(oldValue != newValue)
                                    {
                                        parameter.Set(newValue);
                                        elementUpdated = true;
                                    }
                                }
                            }
                            if (elementUpdated) updatedCount++;
                        }
                    }
                    trans.Commit();
                }
            }
            return updatedCount;
        }
    }
}
