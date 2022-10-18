using Microsoft.VisualBasic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;

namespace ExcelLab
{
    public partial class Form1 : Form
    {
        Data data;
        string _textboxExp = String.Empty;
        string Filename;
        SaveFileDialog saveFileDialog1 = new SaveFileDialog();

        public Form1()
        {
            InitializeComponent();
            data = new Data(dataGridView1);
            this.Text = "Excel";
        }

        /*internal Data Data
        {
            get => default(Data);
            set { }
        }*/
        private void Form1_Load(object sender, EventArgs e)
        {
            data.FillData(Mode.Value);
        }

        private void dataGridView_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Value.ToString()))
            {
                data.ChangeData(e.Value.ToString(), e.RowIndex, e.ColumnIndex);
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void dataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count == 1)
            {
                var selectedCell = dataGridView1.SelectedCells[0];
                textBox1.Text = Data.cells[selectedCell.RowIndex][selectedCell.ColumnIndex].Expression;
                _textboxExp = textBox1.Text;
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (Data.cells[e.RowIndex][e.ColumnIndex].Expression != null)
                if (!String.IsNullOrEmpty(Data.cells[e.RowIndex][e.ColumnIndex].Error))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Error;
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Value.ToString();
                }
        }

        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Expression;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count == 1)
            {
                var selected = dataGridView1.SelectedCells[0];
                if (textBox1.Text == String.Empty)
                {
                    Data.cells[selected.RowIndex][selected.ColumnIndex].Expression = null;
                    Data.cells[selected.RowIndex][selected.ColumnIndex].Value = 0;
                    dataGridView1.Rows[selected.RowIndex].Cells[selected.ColumnIndex].Value = "";
                }
                else
                {
                    data.ChangeData(textBox1.Text, selected.RowIndex, selected.ColumnIndex);
                    if (!string.IsNullOrEmpty(Data.cells[selected.RowIndex][selected.ColumnIndex].Error))
                        dataGridView1.Rows[selected.RowIndex].Cells[selected.ColumnIndex].Value = Data.cells[selected.RowIndex][selected.ColumnIndex].Error;
                    else
                    {
                        dataGridView1.Rows[selected.RowIndex].Cells[selected.ColumnIndex].Value = Data.cells[selected.RowIndex][selected.ColumnIndex].Value.ToString();
                    }
                }
            }
        }
        private void dataGridView_editingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox tb = (TextBox)e.Control;
            tb.TextChanged += Tb_TextChanged;
        }
        private void Tb_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = ((TextBox)sender).Text;
            _textboxExp = textBox1.Text;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            data.AddRow();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            data.AddColumn();
        }
        private void applybutton_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            data = new Data(dataGridView1);
            data.FillData(Mode.Value);
            _textboxExp = "";
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            textBox1.Text = _textboxExp;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            data.AddColumn();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            data.RemoveColumn();
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click_1(object sender, EventArgs e)
        {
            data.RemoveRow();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            DialogResult result = MessageBox.Show("Do you want to save changes?", "Confirm Exit", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if(result == DialogResult.Yes)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    data.SaveToFile(saveFileDialog1.FileName);
                    Filename = saveFileDialog1.FileName;
                    this.Text = Filename + " - MyExcel";
                }

            }
            else if(result == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void openToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                data = new Data(dataGridView1);
                data.OpenFile(openFileDialog1.FileName);
                data.FillData(Mode.Value);
                Filename = openFileDialog1.FileName;
                this.Text = Filename + " - Excel";
            }
        }

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                data.SaveToFile(saveFileDialog1.FileName);
                Filename = saveFileDialog1.FileName;
                this.Text = Filename + " - MyExcel";
            }
        }
       
    }
}