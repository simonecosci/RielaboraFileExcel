using ExcelTools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RielaboraFileExcel
{
    public partial class FormRielabora : Form
    {
        string azienda;
        string schema_glamour;
        DataTable data;
        SqlConnection connection;
        Dictionary<string, string> map;

        readonly object stateLock = new object();
        int target;
        int currentCount;
        int updatedCount;
        volatile bool stop;

        string oldvalue;

        public FormRielabora()
        {
            azienda = ConfigurationManager.AppSettings["azienda"];
            schema_glamour = ConfigurationManager.AppSettings["schema_glamour"];
            string connetionString = ConfigurationManager.AppSettings["connectionString"];
            connection = new SqlConnection(connetionString);
            stop = false;
            InitializeComponent();
        }

        private void ApriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog1.Title = "Seleziona File da processare";
            OpenFileDialog1.FileName = "";
            OpenFileDialog1.Filter = "Excel File|*.xls;*.xlsx";

            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = OpenFileDialog1.FileName;
                if (sFileName.Trim() != "")
                {
                    if (sFileName == string.Empty)
                    {
                        MessageBox.Show("Selezionare un file da caricare");
                        return;
                    }
                    if (!File.Exists(sFileName))
                    {
                        MessageBox.Show("file non trovato");
                        return;
                    }
                    data = LoadFile(sFileName);
                    dataGridView1.DataSource = data;
                    if (data != null)
                    {
                        map = new Dictionary<string, string>();
                        map.Add("CodiceArticolo", data.Columns[0].ColumnName);
                        buttonColumnsMapping.Enabled = true;
                        buttonElabora.Enabled = true;
                        buttonStop.Enabled = false;
                    }
                }
            }
        }

        private DataTable LoadFile(string filename)
        {
            try
            {
                ExcelTool xls = new ExcelTool(filename);
                xls.SetDriver("NPOI");
                return xls.ToDataTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }

        private void ChiudiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            buttonColumnsMapping.Enabled = false;
        }

        private void EsciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void SalvaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Nessuna riga da salvare");
                return;
            }
            data = (DataTable) dataGridView1.DataSource;
            SaveFileDialog1.Title = "Salva File";
            SaveFileDialog1.Filter = "File Excel (*.xls)|*.xls|Tutti i file (*.*)|*.*";
            if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = SaveFileDialog1.FileName;
                if (sFileName.Trim() != "")
                {
                    try
                    {
                        ExcelTool writer = new ExcelTool(sFileName);
                        writer.SetDriver("NPOI");
                        writer.FromDataTable(data);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void ButtonColumnsMapping_Click(object sender, EventArgs e)
        {
            ColumnsMapping FormColumnMapping = new ColumnsMapping(map, this);
            DialogResult dr = FormColumnMapping.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                map = FormColumnMapping.GetMap();
                FormColumnMapping.Dispose();
            }
        }

        void StartThread(object sender, EventArgs e)
        {
            Invoke(new MethodInvoker(JobStarted));
            lock (stateLock)
            {
                target = 0;
            }
            Thread t = new Thread(new ThreadStart(ThreadJob));
            t.IsBackground = true;
            t.Start();
        }

        void ThreadJob()
        {
            int localTarget;
            lock (stateLock)
            {
                localTarget = target;
            }
            MethodInvoker updateCounterDelegate = new MethodInvoker(UpdateCount);
            Invoke(updateCounterDelegate);
            lock (stateLock)
            {
                updatedCount = 0;
                currentCount = 0;
            }
            ProcessRows();
            Invoke(new MethodInvoker(JobFinished));
        }

        void JobStarted()
        {
            progressBar1.Enabled = true;
            progressBar1.Value = 0;
            buttonColumnsMapping.Enabled = false;
            buttonStop.Enabled = true;
            stop = false;
        }

        void JobFinished()
        {
            progressBar1.Value = 0;
            progressBar1.Enabled = false;
            buttonColumnsMapping.Enabled = true;
            buttonStop.Enabled = false;
            stop = false;
            if (connection.State != ConnectionState.Closed)
                connection.Close();
            MessageBox.Show("Processo terminato\nAggiornati " + updatedCount + " articoli");
        }

        void JobInterrupted()
        {
            progressBar1.Value = 0;
            progressBar1.Enabled = false;
            buttonColumnsMapping.Enabled = true;
            buttonStop.Enabled = false;
            stop = false;
            if (connection.State != ConnectionState.Closed)
                connection.Close();
            MessageBox.Show("Processo interrotto");
        }

        void UpdateCount()
        {
            int tmpCount;
            lock (stateLock)
            {
                tmpCount = currentCount;
            }
            if (dataGridView1.Rows.Count > 0)
            {
                progressBar1.Maximum = dataGridView1.Rows.Count;
                progressBar1.Value = tmpCount;
            }
        }

        void ProcessRows()
        {
            try
            {
                MethodInvoker updateCounterDelegate = new MethodInvoker(UpdateCount);
                lock (stateLock)
                {
                    target = dataGridView1.Rows.Count;
                }
                connection.Open();
                int c = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    map.TryGetValue("CodiceArticolo", out string CodiceArticoloColumn);
                    string CodiceArticolo = row.Cells[CodiceArticoloColumn].Value.ToString().Trim();

                    map.TryGetValue("DescrizioneSupplementare", out string DescrizioneSupplementareColumn);
                    map.TryGetValue("MadeIn", out string MadeInColumn);
                    map.TryGetValue("NomenclaturaCombinata", out string NomenclaturaCombinataColumn);
                    map.TryGetValue("FornitoreAbituale", out string FornitoreAbitualeColumn);
                    map.TryGetValue("Stagione", out string StagioneColumn);
                    map.TryGetValue("Genere", out string GenereColumn);
                    map.TryGetValue("GruppoMerceologico", out string GruppoMerceologicoColumn);
                    map.TryGetValue("CategoriaOmogenea", out string CategoriaOmogeneaColumn);
                    map.TryGetValue("Marchio", out string MarchioColumn);
                    map.TryGetValue("Famiglia", out string FamigliaColumn);

                    Dictionary<string, string> updateARTICOLI = new Dictionary<string, string>();
                    Dictionary<string, string> updateTCSCHMAS = new Dictionary<string, string>();
                    if (DescrizioneSupplementareColumn != null)
                    {
                        string DescrizioneSupplementare = row.Cells[DescrizioneSupplementareColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(DescrizioneSupplementare))
                            updateARTICOLI.Add("ARDESSUP", DescrizioneSupplementare);
                    }
                    if (MadeInColumn != null)
                    {
                        string MadeIn = row.Cells[MadeInColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(MadeIn)) {
                            updateARTICOLI.Add("TCMADEIN", MadeIn);
                            updateTCSCHMAS.Add("STMADEIN", MadeIn);
                        }
                    }
                    if (NomenclaturaCombinataColumn != null)
                    {
                        string NomenclaturaCombinata = row.Cells[NomenclaturaCombinataColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(NomenclaturaCombinata))
                        {
                            updateARTICOLI.Add("ARNOMENC", NomenclaturaCombinata);
                            updateTCSCHMAS.Add("STNOMENC", NomenclaturaCombinata);
                        }
                    }
                    if (FornitoreAbitualeColumn != null)
                    {
                        string FornitoreAbituale = row.Cells[FornitoreAbitualeColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(FornitoreAbituale))
                            updateTCSCHMAS.Add("STCODFOR", FornitoreAbituale);
                    }
                    if (StagioneColumn != null)
                    {
                        string Stagione = row.Cells[StagioneColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(Stagione))
                        {
                            updateARTICOLI.Add("ARSTAGIO", Stagione);
                            updateTCSCHMAS.Add("STSTAGIO", Stagione);
                        }
                    }
                    if (GenereColumn != null)
                    {
                        string Genere = row.Cells[GenereColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(Genere))
                        {
                            updateARTICOLI.Add("TCGENERE", Genere);
                            updateTCSCHMAS.Add("STGENERE", Genere);
                        }
                    }
                    if (GruppoMerceologicoColumn != null)
                    {
                        string GruppoMerceologico = row.Cells[GruppoMerceologicoColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(GruppoMerceologico))
                        {
                            updateARTICOLI.Add("ARGRUMER", GruppoMerceologico);
                            updateTCSCHMAS.Add("STGRUMER", GruppoMerceologico);
                        }
                    }
                    if (CategoriaOmogeneaColumn != null)
                    {
                        string CategoriaOmogenea = row.Cells[CategoriaOmogeneaColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(CategoriaOmogenea)) 
                        {
                            updateARTICOLI.Add("ARCATOMO", CategoriaOmogenea);
                            updateTCSCHMAS.Add("STCATOMO", CategoriaOmogenea);
                        }
                    }
                    if (MarchioColumn != null)
                    {
                        string Marchio = row.Cells[MarchioColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(Marchio))
                        {
                            updateARTICOLI.Add("ARCODMAR", Marchio);
                            updateTCSCHMAS.Add("STMARCHI", Marchio);
                        }
                    }
                    if (FamigliaColumn != null)
                    {
                        string Famiglia = row.Cells[FamigliaColumn].Value.ToString().Trim();
                        if (!string.IsNullOrEmpty(Famiglia))
                        {
                            updateARTICOLI.Add("ARCODFAM", Famiglia);
                            updateTCSCHMAS.Add("STCODFAM", Famiglia);
                        }
                    }

                    bool res = UpdateRow(CodiceArticolo, updateARTICOLI, updateTCSCHMAS);
                    lock (stateLock)
                    {
                        currentCount = c + 1;
                        if (res)
                            updatedCount++;
                    }
                    Invoke(updateCounterDelegate);
                    c++;
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        bool CheckExists(string table, string[] fields, string[] values)
        {
            List<string> where = new List<string>();
            foreach (var field in fields)
            {
                where.Add(field + " = @" + field);
            }
            var query = "SELECT * FROM [" + schema_glamour + "].[dbo].[" + table + "] WHERE " + string.Join(" AND ", where);
            SqlCommand command = new SqlCommand(query, connection);
            for (int i = 0; i < fields.Length; i++)
            {
                command.Parameters.Add(new SqlParameter(fields[i], values[i]));
            }
            SqlDataReader reader = command.ExecuteReader();
            bool exists = reader.HasRows;
            reader.Close();
            command.Dispose();
            return exists;
        }

        bool UpdateRow(string Codice, Dictionary<string, string> updART, Dictionary<string, string> updTCS)
        {
            bool u1 = false, u2 = false;
            if (updART.Count > 0)
            {
                u1 = UpdateTable(Codice, updART, azienda + "ART_ICOL", "ARCODART");
            }
            if (updTCS.Count > 0)
            {
                u2 = UpdateTable(Codice, updTCS, azienda + "TCSCHMAS", "STCODART");
            }
            return u1 || u2;
        }

        bool UpdateTable(string Codice, Dictionary<string, string> updates, string tableName, string key)
        {

            if (!CheckExists(tableName, new string[] { key }, new string[] { Codice }))
                return false;

            List<string> fields = new List<string>();
            foreach (KeyValuePair<string, string> entry in updates)
            {
                fields.Add(entry.Key + "=@" + entry.Key);
            }
            string sql = "UPDATE [" + schema_glamour + "].[dbo].[" + tableName + "] SET " + string.Join(", ", fields) + " WHERE " + key + " = @Codice";
            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.Add(new SqlParameter("Codice", Codice));
            foreach (KeyValuePair<string, string> entry in updates)
            {
                command.Parameters.AddWithValue(entry.Key, entry.Value);
            }
            command.ExecuteNonQuery();
            command.Dispose();
            return true;
        }

        private void buttonElabora_Click(object sender, EventArgs e)
        {
            if (!map.ContainsKey("CodiceArticolo"))
            {
                MessageBox.Show("Manca il mapping per la colonna chiave CodiceArticolo");
                buttonColumnsMapping.Focus();
                return;
            }
            StartThread(sender, e);
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            stop = true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            DialogResult response = MessageBox.Show("Eliminare queste righe?", "Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if ((response == DialogResult.No))
            {
                e.Cancel = true;
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldvalue = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString().Trim();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string newvalue = ((DataGridView)sender).Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim();
            DialogResult response = MessageBox.Show("Modificare tutte righe sostituendo " + oldvalue + " con " + newvalue + " ?", "Modifica", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (oldvalue == newvalue || response == DialogResult.No)
            {
                return;
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string value = row.Cells[e.ColumnIndex].Value.ToString().Trim();
                if (value == oldvalue)
                {
                    row.Cells[e.ColumnIndex].Value = newvalue;
                }
            }
        }
    }
}
