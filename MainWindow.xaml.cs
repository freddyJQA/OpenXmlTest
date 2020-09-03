using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using Bold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using RunSpreadsheet = DocumentFormat.OpenXml.Spreadsheet.Run;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Spreadsheet.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Linq;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using System.IO;

namespace OpenXmlTest
{
    public partial class MainWindow : Window
    {
        private DataSet ds;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Query

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
            btnWord.IsEnabled = rows > decimal.Zero;
        }

        #endregion

        #region Excel

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

            RunSpreadsheet run = new RunSpreadsheet();
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

        #endregion

        #region Word

        private void BtnWord_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Word documents (.docx)|*.docx",
            };

            if (!(bool)openFileDialog.ShowDialog())
                return;

            using (FileStream templateFile = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    // Crea una copia de la plantilla en memoria.
                    templateFile.CopyTo(stream);

                    // Inserta la data en la copia.
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                    {
                        ChangeTextWord(doc);
                        InsertTableWord(doc);
                    };

                    stream.Seek(0, SeekOrigin.Begin);

                    // Guarda la copia en la ruta especificada.
                    SaveFileDialog saveFileDialog = new SaveFileDialog()
                    {
                        FileName = "Word1",
                        Filter = "Word documents (.docx)|*.docx",
                    };

                    if ((bool)saveFileDialog.ShowDialog())
                    {
                        using (FileStream fileStream = File.Create(saveFileDialog.FileName))
                        {
                            stream.CopyTo(fileStream);
                        };
                    }
                };
            };
        }

        private void ChangeTextWord(WordprocessingDocument doc)
        {
            // Encuentra la primera tabla en el documento.
            Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

            // Encuentra la segunda y tercera fila en la tabla.
            TableRow row1 = table.Elements<TableRow>().ElementAt(1);
            TableRow row2 = table.Elements<TableRow>().ElementAt(2);

            // Encuentra las celdas a modificar.
            TableCell cellNombre = row1.Elements<TableCell>().ElementAt(1);
            TableCell cellApellido = row1.Elements<TableCell>().ElementAt(3);
            TableCell cellEdad = row2.Elements<TableCell>().ElementAt(1);
            TableCell cellDireccion = row2.Elements<TableCell>().ElementAt(3);

            // Llena las celdas con los datos de la primera fila de la primera tabla del dataset.
            cellNombre.AppendChild(new Paragraph(new Run(new Text("Freddy"))));
            cellApellido.AppendChild(new Paragraph(new Run(new Text("Quintero"))));
            cellEdad.AppendChild(new Paragraph(new Run(new Text("29"))));
            cellDireccion.AppendChild(new Paragraph(new Run(new Text("Porlamar"))));
        }

        private void InsertTableWord(WordprocessingDocument doc)
        {
            // Encuentra la segunda tabla en el documento.
            Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(1);

            // Encuentra la segunda fila en la tabla.
            TableRow row = table.Elements<TableRow>().ElementAt(1);

            // Encuentra la celda a modificar.
            TableCell cell = row.Elements<TableCell>().First();

            // Crea la tabla.
            Table tbl = new Table();

            // Establece estiloy y anchura a la tabla.
            TableProperties tableProp = new TableProperties();
            TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

            // Hace que la tabla ocupe el 100% de la pagina.
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

            // Aplicar propiedades a la tabla.
            tableProp.Append(tableStyle, tableWidth);
            tbl.AppendChild(tableProp);

            // Define las columnas de la tabla.
            TableGrid tg = new TableGrid();
            foreach (DataColumn column in ds.Tables[0].Columns)
                tg.AppendChild(new GridColumn());
            tbl.AppendChild(tg);

            // Fila para las columnas de la tabla.
            TableRow tblRowColumns = new TableRow();
            tbl.AppendChild(tblRowColumns);

            // Obtiene y asigna nombres a las columnas de la tabla.
            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                TableCell tblCell = new TableCell(new Paragraph(new Run(new Text(column.ColumnName))));
                tblRowColumns.AppendChild(tblCell);
            }

            // Agrega el resto de las filas a la tabla.
            foreach (DataRow dtRow in ds.Tables[0].Rows)
            {
                TableRow tblRow = new TableRow();

                for (int i = 0; i < dtRow.Table.Columns.Count; i++)
                {
                    TableCell tblCell = new TableCell(new Paragraph(new Run(new Text(dtRow[i].ToString()))));
                    tblRow.AppendChild(tblCell);
                }

                tbl.AppendChild(tblRow);
            }

            // Agrega la tabla al placeholder correspondiente.
            cell.AppendChild(new Paragraph(new Run(tbl)));
        }

        #endregion
    }
}
