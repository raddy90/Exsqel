using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace Exsqel
{
    public partial class Exsqel : Form
    {

        private List<string> excelColumnNames = new List<string>();
        private List<string> columnNamesDDBB = new List<string>();
        Dictionary<string, string> columnNamesDDBBtipos = new Dictionary<string, string>();


        private bool conectado1vez = false;
        public Exsqel()
        {
            InitializeComponent();
        }



        DataTableCollection tableCollection;

        private DataTable dataTable;
        private string filePath;

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel .xlsx, .xls o .xlsm|*.xlsx;*.xls;*.xlsm";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    txtArchivo.Text = filePath;

                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            dataTable = result.Tables[0];
                            dataGridView1.DataSource = dataTable;
                            tableCollection = result.Tables;
                            comboTablaExcel.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                comboTablaExcel.Items.Add(table.TableName);
                                
                            }
                            comboTablaExcel.SelectedIndex = 0;


                            btnVerModificar.Visible = true;
                            
                            btnVerConectar.Visible = true;
                        }
                    }
                }
            }
        }

        
    private void NombreValorColumnaNueva(object sender, EventArgs e)
        {
            if(textNombreNcolumna.Text.Length != 0)
            {
                btnAgregarColumna.Enabled = true;
            } else
            {
                btnAgregarColumna.Enabled=false;
            }
        }

        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            if (dataTable != null)
            {
                string columnName = textNombreNcolumna.Text; 
                dataTable.Columns.Add(columnName, typeof(string)); 

                foreach (DataRow row in dataTable.Rows)
                {
                    row[columnName] = textValorNcolumna.Text; 
                }

                dataGridView1.DataSource = null;
                dataGridView1.DataSource = dataTable;
                textValorNcolumna.Text = null;
                textNombreNcolumna.Text = null;
                btnAgregarColumna.Enabled = false;
            }
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            if (dataTable != null)
            {
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.Add(newRow);
            }
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            if (dataTable != null && !string.IsNullOrEmpty(filePath))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                    package.Save();
                    MessageBox.Show("Archivo de Excel guardado correctamente.");
                }
            }
        }


        private DataTable GetExcelData()
        {
            return (DataTable)dataGridView1.DataSource;
        }

       

        private void comboTablaExcel_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[comboTablaExcel.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;

            excelColumnNames.Clear();
            foreach (DataColumn column in dt.Columns)
            {
                excelColumnNames.Add(column.ColumnName);
            }

            GetTableColumnNames();
            DisplayTableSchema();
           
        }


        private void GetTableColumnNames()
        {
            
            if (conectado1vez)
            {
                columnNamesDDBB.Clear();
                using (var connection = new SqlConnection(txtConnectionString.Text))
                {
                    try
                    {
                        connection.Open();
                        var schemaTable = connection.GetSchema("Columns", new string[] { null, null, txtTablaSQL.Text, null });

                        foreach (DataRow row in schemaTable.Rows)
                        {

                            string columnName = row["COLUMN_NAME"].ToString();
                            string columnType = row["DATA_TYPE"].ToString();

                            columnNamesDDBB.Add(columnName);
                            
                            columnNamesDDBBtipos[columnName] = columnType;

                        }
                    }
                    catch (SqlException sqlEx)
                    {

                        MessageBox.Show($"Error de SQL: {sqlEx.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (ArgumentException argEx)
                    {

                        MessageBox.Show($"Error de validación: {argEx.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show($"Error: {ex.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

        }

        private void DisplayTableSchema()
        {

            flowLayoutPanelFields.Controls.Clear();

            var x = 0;

            if(columnNamesDDBB != null) { 

            foreach (string columnName in columnNamesDDBB)
            {

                var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight };

                var comboBoxExcel = new ComboBox
                {
                    Width = 180, 
                    Name = "ComboBoxExcel" + flowLayoutPanelFields.Controls.Count 
                }; 
                
                comboBoxExcel.Items.AddRange(excelColumnNames.ToArray());
                    if (excelColumnNames.Count > x)
                    {
                        comboBoxExcel.SelectedIndex = x;
                    } else
                    {
                        comboBoxExcel.SelectedIndex = 0;
                    }

                panel.Controls.Add(comboBoxExcel);


                var comboBoxDb = new ComboBox
                {
                    Width = 180, 
                    Name = "ComboBoxBD" + flowLayoutPanelFields.Controls.Count
                };
                comboBoxDb.Items.AddRange(columnNamesDDBB.ToArray());
                comboBoxDb.SelectedIndex = x;
                panel.Controls.Add(comboBoxDb);


                var textBoxCond = new TextBox
                {
                    Width = 80, 
                    Name = "textBoxCond" + flowLayoutPanelFields.Controls.Count
                };

                textBoxCond.Tag = " Yes ";

                textBoxCond.ForeColor = System.Drawing.Color.Gray;
                textBoxCond.Text = " Yes ";
                textBoxCond.Enter += (s, e) =>
                {
                    if (textBoxCond.Text == " Yes ")
                    {
                        textBoxCond.Text = "";
                        textBoxCond.ForeColor = System.Drawing.Color.Black;
                    }
                };
                textBoxCond.Leave += (s, e) =>
                {
                    if (string.IsNullOrWhiteSpace(textBoxCond.Text))
                    {
                        textBoxCond.Text = " Yes ";
                        textBoxCond.ForeColor = System.Drawing.Color.Gray;
                    }
                };
                panel.Controls.Add(textBoxCond);


                    //////////
                    /////////

                var textBoxCond2 = new TextBox
                {
                    Width = 80, 
                    Name = "textBoxCond2" + flowLayoutPanelFields.Controls.Count 
                };

                textBoxCond2.Tag = " true ";

                textBoxCond2.ForeColor = System.Drawing.Color.Gray;
                textBoxCond2.Text = " true ";
                textBoxCond2.Enter += (s, e) =>
                {
                    if (textBoxCond2.Text == " true ")
                    {
                        textBoxCond2.Text = "";
                        textBoxCond2.ForeColor = System.Drawing.Color.Black;
                    }
                };
                textBoxCond2.Leave += (s, e) =>
                {
                    if (string.IsNullOrWhiteSpace(textBoxCond2.Text))
                    {
                        textBoxCond2.Text = " true ";
                        textBoxCond2.ForeColor = System.Drawing.Color.Gray;
                    }
                };
                panel.Controls.Add(textBoxCond2);

               

                    ///////////
                    //////////


                    var textBoxCondOR = new TextBox
                    {
                        Width = 80, 
                        Name = "textBoxCondOR" + flowLayoutPanelFields.Controls.Count 
                    };

                    textBoxCondOR.Tag = " No ";

                    textBoxCondOR.ForeColor = System.Drawing.Color.Gray;
                    textBoxCondOR.Text = " No ";
                    textBoxCondOR.Enter += (s, e) =>
                    {
                        if (textBoxCondOR.Text == " No ")
                        {
                            textBoxCondOR.Text = "";
                            textBoxCondOR.ForeColor = System.Drawing.Color.Black;
                        }
                    };
                    textBoxCondOR.Leave += (s, e) =>
                    {
                        if (string.IsNullOrWhiteSpace(textBoxCondOR.Text))
                        {
                            textBoxCondOR.Text = " No ";
                            textBoxCondOR.ForeColor = System.Drawing.Color.Gray;
                        }
                    };
                    panel.Controls.Add(textBoxCondOR);


                    /////////////
                    ////////////


                    var textBoxCondOR2 = new TextBox
                    {
                        Width = 80, 
                        Name = "textBoxCondOR2" + flowLayoutPanelFields.Controls.Count
                    };

                    textBoxCondOR2.Tag = " false ";

                    textBoxCondOR2.ForeColor = System.Drawing.Color.Gray;
                    textBoxCondOR2.Text = " false ";
                    textBoxCondOR2.Enter += (s, e) =>
                    {
                        if (textBoxCondOR2.Text == " false ")
                        {
                            textBoxCondOR2.Text = "";
                            textBoxCondOR2.ForeColor = System.Drawing.Color.Black;
                        }
                    };
                    textBoxCondOR2.Leave += (s, e) =>
                    {
                        if (string.IsNullOrWhiteSpace(textBoxCondOR2.Text))
                        {
                            textBoxCondOR2.Text = " false ";
                            textBoxCondOR2.ForeColor = System.Drawing.Color.Gray;
                        }
                    };
                    panel.Controls.Add(textBoxCondOR2);

                    x++;

                    panel.Width = 750; 
                    panel.Height = 24;


                    flowLayoutPanelFields.Controls.Add(panel);

            }


            }
        }


        private void btnImportar_Click(object sender, EventArgs e)
        {

            txtGuardando.Visible = true;
            btnImportar.Enabled = false;
            btnConectar.Enabled = false;
            btnSeleccionar.Enabled = false;

            Cursor = Cursors.WaitCursor;

            var data = new List<Dictionary<string, object>>();
            DataTable excelData = GetExcelData(); 

            foreach (DataRow row in excelData.Rows)
            {
                var record = new Dictionary<string, object>();
                foreach (Control control in flowLayoutPanelFields.Controls)
                {
                    if (control is FlowLayoutPanel panel)
                    {
                        ComboBox comboBoxExcel = panel.Controls.OfType<ComboBox>().FirstOrDefault();
                        ComboBox comboBoxDb = panel.Controls.OfType<ComboBox>().LastOrDefault();

                        var textBoxes = panel.Controls.OfType<TextBox>().ToList();

                        TextBox textBoxCond = textBoxes.ElementAtOrDefault(0);
                        TextBox textBoxCond2 = textBoxes.ElementAtOrDefault(1);
                        TextBox textBoxCondOR = textBoxes.ElementAtOrDefault(2);
                        TextBox textBoxCondOR2 = textBoxes.ElementAtOrDefault(3);

                        if (comboBoxDb != null && comboBoxExcel != null && comboBoxDb.SelectedItem != null && comboBoxExcel.SelectedItem != null)
                        {

                            string dbField = comboBoxDb.SelectedItem.ToString();
                            string excelField = comboBoxExcel.SelectedItem.ToString();


                            object value = row[excelField];

                             if (!string.IsNullOrWhiteSpace(textBoxCond.Text) && !string.IsNullOrWhiteSpace(textBoxCond2.Text) && textBoxCond.Text.ToString() != " Yes " && textBoxCond2.Text.ToString() != " true " && (value.Equals(textBoxCond.Text.ToString())))
                            {
                                 if (int.TryParse(textBoxCond2.Text.ToString(), out int intValue))
                                {
                                    value = intValue;
                                }
                                else if (double.TryParse(textBoxCond2.Text.ToString(), out double doubleValue))
                                {
                                    value = doubleValue;
                                }
                                else if (bool.TryParse(textBoxCond2.Text.ToString(), out bool boolValue))
                                {
                                    value = boolValue;
                                }
                                else if (DateTime.TryParse(textBoxCond2.Text.ToString(), out DateTime dateTimeValue))
                                {
                                    value = dateTimeValue;
                                }
                                else
                                {

                                    value = textBoxCond2.Text.ToString();
                                }


                            }
                            if (!string.IsNullOrWhiteSpace(textBoxCondOR.Text) && !string.IsNullOrWhiteSpace(textBoxCondOR2.Text) && textBoxCondOR.Text.ToString()!= " No " && textBoxCondOR2.Text.ToString() != " false " && (value.Equals(textBoxCondOR.Text.ToString())))
                            {
                                 if (int.TryParse(textBoxCondOR2.Text.ToString(), out int intValue))
                                {
                                    value = intValue;
                                }
                                else if (double.TryParse(textBoxCondOR2.Text.ToString(), out double doubleValue))
                                {
                                    value = doubleValue;
                                }
                                else if (bool.TryParse(textBoxCondOR2.Text.ToString(), out bool boolValue))
                                {
                                    value = boolValue;
                                }
                                else if (DateTime.TryParse(textBoxCondOR2.Text.ToString(), out DateTime dateTimeValue))
                                {
                                    value = dateTimeValue;
                                }
                                else
                                {

                                    value = textBoxCondOR2.Text.ToString();
                                }


                            }
                           
                            
                                record[dbField] = value.ToString();
                            
                            
                        } else if(comboBoxDb != null && comboBoxExcel != null && comboBoxDb.SelectedItem != null)
                        {
                            string dbField = comboBoxDb.SelectedItem.ToString();


                            record[dbField] = comboBoxExcel.Text.ToString();
                            
                        }

                    }
                }

                data.Add(record);
            }

            guardarDatosBBDD(data);

        }

        public class MyData
        {
            public string FieldName { get; set; }
            public string FieldValue { get; set; }
        }

      
        private void guardarDatosBBDD(List<Dictionary<string, object>> data)
        {

            string connectionString = txtConnectionString.Text;
            string tableName = txtTablaSQL.Text;

            DapperPlusManager.Entity<dynamic>().Table(tableName);

                Console.WriteLine("rrrr");

            var dataList = data.Select((record, i) =>
            {
                Console.WriteLine("rrrr1", record.Values.ToString());

                dynamic expando = new ExpandoObject();
                var expandoDict = (IDictionary<string, object>)expando;

                foreach (var kvp in record)
                {
                    Console.WriteLine("rrrr2",kvp.Value.ToString());

                    if (kvp.Value.ToString() != "" && kvp.Value.ToString() != null)
                    {

                        try
                        {
                          
                            if (columnNamesDDBBtipos.ContainsKey(kvp.Key.ToString()) && kvp.Value.ToString() != null && kvp.Value.ToString() != "" && columnNamesDDBBtipos[kvp.Key.ToString()] == "bit")
                            {

                                expandoDict.Add(kvp.Key.ToString(), true);
                              
                            }
                            else if (columnNamesDDBBtipos.ContainsKey(kvp.Key.ToString()) && columnNamesDDBBtipos[kvp.Key.ToString()] == "decimal")
                            {
                                if (kvp.Value.ToString() != null && kvp.Value.ToString() !="")
                                {
                                    try
                                    {
                                        
                                        decimal? decimalValue = null;

                                        if (!string.IsNullOrWhiteSpace(kvp.Value.ToString()))
                                        {
                                            if (decimal.TryParse(kvp.Value.ToString().Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsedDecimal))
                                            {
                                                decimalValue = parsedDecimal;
                                                expandoDict.Add(kvp.Key.ToString(), parsedDecimal);
                                            }
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }

                                }
                                else
                                {
                                    expandoDict.Add(kvp.Key.ToString(), null);
                                }

                            }
                            else if (columnNamesDDBBtipos.ContainsKey(kvp.Key.ToString()) && kvp.Value.ToString() != null && kvp.Value.ToString() != "" && columnNamesDDBBtipos[kvp.Key.ToString()] == "datetime")
                            {
                                try
                                {
                                    DateTime fecha;
                                    string laFecha = kvp.Value.ToString();
                                    

                                    if (DateTime.TryParseExact(laFecha, "dd-MM-yyyyTHH:mm", null, System.Globalization.DateTimeStyles.None, out fecha))
                                    {
                                       
                                        if (fecha < new DateTime(1753, 1, 1) || fecha > new DateTime(9999, 12, 31))
                                        {
                                            expandoDict.Add(kvp.Key.ToString(), DateTime.Parse(new DateTime(2024, 12, 12, 0, 0, 0).ToString("dd-MM-yyyyThh:mm")));
                                        }
                                        else
                                        {
                                            expandoDict.Add(kvp.Key.ToString(), DateTime.Parse(fecha.ToString()));
                                        }

                                    }
                                    else
                                    {
                                        expandoDict.Add(kvp.Key.ToString(), DateTime.Parse(new DateTime(2024, 12, 12, 0, 0, 0).ToString("dd-MM-yyyyThh:mm")));
                                    }

                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                            }

                            else if (columnNamesDDBBtipos.ContainsKey(kvp.Key.ToString()) && kvp.Value.ToString() != null && kvp.Value.ToString() != "" && columnNamesDDBBtipos[kvp.Key.ToString()] == "uniqueidentifier")
                            {
                                try
                                {
                                    
                                    expandoDict.Add(kvp.Key.ToString(), Guid.Parse(kvp.Value.ToString()));
                                   
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                            }
                            
                            else
                            {

                                expandoDict.Add(kvp.Key.ToString(), kvp.Value.ToString().Replace("'",""));
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                    else
                    {

                       
                        expandoDict.Add(kvp.Key.ToString(), null);
                       

                    }
                }

                return expando;

            }).ToList();


            try
            {


                if (dataList != null)
                {
                    using (IDbConnection db = new SqlConnection(connectionString))
                    {
                        db.BulkInsert(dataList);
                    }
                }
                
                MessageBox.Show("Hecho.\n Registros insertados: " + dataList.Count());

                btnImportar.Enabled = true;
                btnConectar.Enabled = true;
                btnSeleccionar.Enabled = true;

                Cursor = Cursors.Default;
                txtGuardando.Visible = false;

            }
            catch (SqlException sqlEx)
            {
                btnImportar.Enabled = true;
                btnConectar.Enabled = true;
                btnSeleccionar.Enabled = true;

                Cursor = Cursors.Default;
                txtGuardando.Visible = false;

                MessageBox.Show($"Error de SQL: {sqlEx.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ArgumentException argEx)
            {
                btnImportar.Enabled = true;
                btnConectar.Enabled = true;
                btnSeleccionar.Enabled = true;

                Cursor = Cursors.Default;
                txtGuardando.Visible = false;

                MessageBox.Show($"Error de validación: {argEx.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                btnImportar.Enabled = true;
                btnConectar.Enabled = true;
                btnSeleccionar.Enabled = true;

                Cursor = Cursors.Default;
                txtGuardando.Visible = false;

                MessageBox.Show($"Error: {ex.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnVerConectar_Click(object sender, EventArgs e)
        {
            flowLayoutPanel2.Visible = true;
            txtConexionSQL.Visible = true;
            btnVerConectar.Visible = false;

        }

        private void btnVerModificar_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Visible = true;
            txtCambiarExcel.Visible = true;
            btnVerModificar.Visible = false;
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            txtColumnaExcel.Visible = true;
            txtColumnaSQL.Visible = true;
            txtSi.Visible = true;
            txtOsi.Visible = true;
            txtEntonces1.Visible = true;
            txtEntonces2.Visible = true;

            string connectionString = txtConnectionString.Text;
            string tableName = txtTablaSQL.Text;
            
            conectado1vez = true;

            GetTableColumnNames();

            DisplayTableSchema();


            flowLayoutPanelFields.Visible = true;
            btnImportar.Visible = true;

        }
    }
}
