using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data.OleDb;
using System.Windows.Forms.DataVisualization.Charting;

namespace ANALISI_AUTISTI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            
        }

        

   

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Apri una finestra di dialogo per selezionare un file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Ottieni il percorso del file selezionato
                string filePath = openFileDialog1.FileName;

                // Crea una connessione OLEDB per leggere il file Excel
                string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'", filePath);
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    // Apri la connessione e leggi il contenuto del file Excel in un DataSet
                    connection.Open();
                    using (DataSet ds = new DataSet())
                    {
                        // Ottieni il nome della prima tabella nel file Excel
                        DataTable dtSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();

                        // Leggi il contenuto del file Excel nella prima tabella in un DataSet
                        OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
                        adapter.Fill(ds);

                        // Assegna il DataSet come DataSource della DataGridView
                        dataGridView1.DataSource = ds.Tables[0];
                    }
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Ottieni il nome della colonna selezionata
            string columnName = comboBox2.SelectedItem.ToString();

            // Creazione di un oggetto DataView per filtrare il contenuto della DataGridView
            DataView dv = ((DataTable)dataGridView1.DataSource).DefaultView;

            // Apertura della finestra di dialogo di ricerca
            string filterText = Microsoft.VisualBasic.Interaction.InputBox("Inserisci un valore di ricerca per la colonna " + columnName + ": ", "Ricerca");

            // Se l'utente ha inserito un valore di ricerca, applicare il filtro
            if (!string.IsNullOrEmpty(filterText))
            {
                dv.RowFilter = columnName + " LIKE '%" + filterText + "%'";
            }
            else
            {
                dv.RowFilter = "";
            }

            // Assegnazione della DataView alla DataGridView
            dataGridView1.DataSource = dv;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // Recupera la riga selezionata nella DataGridView
            DataGridViewRow selectedRow = dataGridView1.CurrentRow;

            // Recupera la data dalla colonna "Data" nella riga selezionata
            DateTime date = Convert.ToDateTime(selectedRow.Cells["Data"].Value);

            // Recupera i valori dalla colonna "Consegnati" e "Usciti" nella riga selezionata
            int consegnati = Convert.ToInt32(selectedRow.Cells["Consegnati"].Value);
            int usciti = Convert.ToInt32(selectedRow.Cells["Usciti"].Value);

            // Aggiorna il grafico Chart1 con i valori recuperati
            chart1.Series["Consegnati"].Points.AddXY(date, consegnati);
            chart1.Series["Usciti"].Points.AddXY(date, usciti);
        }

        
    }
}
