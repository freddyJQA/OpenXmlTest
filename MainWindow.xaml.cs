using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace OpenXmlTest
{
    public partial class MainWindow : Window
    {
        private DataSet ds;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnEjecutar_Click(object sender, RoutedEventArgs e)
        {
            string connStr = ConfigurationManager.ConnectionStrings["BD"].ConnectionString;
            string query = txtPrueba.Text.Trim();
            decimal rows = decimal.Zero;

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                ds = new DataSet();

                using (SqlCommand command = new SqlCommand(query, conn))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);
                            ds.Tables.Add(dt);
                            rows += dt.Rows.Count;
                        }
                    };
                };
            };

            btnExportar.IsEnabled = rows > decimal.Zero;
        }

        private void BtnExportar_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                FileName = "Libro1",
                Filter = "Excel documents (.xlsx)|*.xlsx",
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result is true)
            {
                string filename = saveFileDialog.FileName;
                ExportToExcel(filename);
            }
        }

        private void ExportToExcel(string filepath)
        {
            // Por defecto, AutoSave = true, Editable = true y Type = xlsx.
            using (SpreadsheetDocument ssd = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                // Agrega un WorkbookPart al documento.
                WorkbookPart workbookPart = ssd.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Agrega Sheets al Workbook.
                Sheets sheets = ssd.WorkbookPart.Workbook.AppendChild(new Sheets());

                for (int tableIndex = 0; tableIndex < ds.Tables.Count; tableIndex++)
                {
                    // Agrega un WorksheetPart al WorkbookPart.
                    WorksheetPart wsp = workbookPart.AddNewPart<WorksheetPart>();
                    wsp.Worksheet = new Worksheet(new SheetData());

                    // Crea una Sheet y la asocia con el WorkbookPart.
                    Sheet sheet = new Sheet()
                    {
                        Id = ssd.WorkbookPart.GetIdOfPart(wsp),
                        SheetId = (uint)tableIndex + 1,
                        Name = "Hoja" + tableIndex
                    };
                    sheets.Append(sheet);

                    WriteData(wsp, tableIndex);                    
                }                

                workbookPart.Workbook.Save();
            };
        }

        private void WriteData(WorksheetPart worksheetPart, int tableIndex)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            WriteColums(sheetData, tableIndex);
            WriteRows(sheetData, tableIndex);
        }

        private void WriteColums(SheetData sheetData, int tableIndex)
        {
            Row row = new Row() { RowIndex = 1 };
            sheetData.Append(row);

            foreach (DataColumn column in ds.Tables[tableIndex].Columns)
            {
                RunProperties runProperties = new RunProperties();
                runProperties.Append(new Bold());
                runProperties.Append(new Color() { Rgb = "aaaaaa" });
                runProperties.Append(new FontSize() { Val = 20 });

                Cell cell = CreateTextCell(ds.Tables[tableIndex].Columns.IndexOf(column) + 1, 1, column.ColumnName, runProperties);
                row.AppendChild(cell);
            }
        }

        private void WriteRows(SheetData sheetData, int tableIndex)
        {
            for (int i = 0; i < ds.Tables[tableIndex].Rows.Count; i++)
            {
                DataRow dataRow = ds.Tables[tableIndex].Rows[i];
                Row row = CreateRow(dataRow, i + 2);
                sheetData.AppendChild(row);
            }
        }

        private Row CreateRow(DataRow dataRow, int rowIndex)
        {
            Row row = new Row
            {
                RowIndex = (uint)rowIndex
            };

            for (int i = 0; i < dataRow.Table.Columns.Count; i++)
            {
                Cell dataCell = CreateTextCell(i + 1, rowIndex, dataRow[i]);
                row.AppendChild(dataCell);
            }

            return row;
        }

        private Cell CreateTextCell(int columnIndex, int rowIndex, object cellValue, RunProperties runProperties = null)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = GetColumnName(columnIndex) + rowIndex,
            };           

            Text text = new Text
            {
                Text = cellValue.ToString()
            };

            Run run = new Run();
            run.Append(text);

            if (runProperties != null)
                run.RunProperties = runProperties;

            InlineString inlineString = new InlineString();
            inlineString.Append(run);
            cell.AppendChild(inlineString);

            return cell;
        }

        private string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (dividend - modifier) / 26;
            }

            return columnName;
        }
    }
}
