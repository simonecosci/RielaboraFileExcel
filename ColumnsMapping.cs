using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RielaboraFileExcel
{
    public partial class ColumnsMapping : Form
    {
        private Dictionary<string, string> map;

        public List<string> Columns = new List<string>() {
            "CodiceArticolo",
            "DescrizioneSupplementare",
            "MadeIn",
            "NomenclaturaCombinata",
            "FornitoreAbituale",
            "Stagione",
            "Genere",
            "GruppoMerceologico",
            "CategoriaOmogenea",
            "Marchio",
            "Famiglia"
        };

        public ColumnsMapping(Dictionary<string, string> map, FormRielabora formRielabora)
        {
            this.map = map;
            InitializeComponent();

            DataGridViewComboBoxColumn firstColumn = (DataGridViewComboBoxColumn)dataGridView1.Columns[0];
            foreach(string item in Columns)
                firstColumn.Items.Add(item);

            DataGridViewComboBoxColumn secondColumn = (DataGridViewComboBoxColumn) dataGridView1.Columns[1];
            foreach (DataGridViewColumn column in formRielabora.dataGridView1.Columns)
            {
                secondColumn.Items.Add(column.Name);
            }
            foreach (KeyValuePair<string, string> entry in map)
            {
                dataGridView1.Rows.Add(entry.Key, entry.Value);
            }
        }

        public Dictionary<string, string> GetMap()
        {
            map = new Dictionary<string, string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    map.Add(row.Cells[0].Value.ToString().Trim(), row.Cells[1].Value.ToString().Trim());
                }
            }
            return map;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
