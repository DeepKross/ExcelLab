using FileManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelLab
{
    enum Mode { Expression, Value}
    class Data
    {
        private int num_col = 10;
        private int num_row = 10;
        Parser parser = new Parser();
        DataGridView dataGridView1;

        public static List<List<Cell>> cells = new List<List<Cell>>();

        public Data(DataGridView _dataGridView)
        {
            dataGridView1 = _dataGridView;
            dataGridView1.AllowUserToAddRows = false;
            cells.Clear();
            for (int i = 0; i < num_row; i++)
            {
                cells.Add(new List<Cell>());
                for (int j = 0; j < num_col; j++)
                {
                    cells[i].Add(new Cell() { RowNumber = i + 1, ColumnLetter = Convert.ToChar('A' + j) });
                }
            }
        }

       /* public Cell Cell
        {
            get => default(Cell);
            set
            {

            }
        }*/

        /*public Parser Parser
        {
            get => default(Parser);
            set
            {

            }
        }*/

        public void AddRow()
        {
            cells.Add(new List<Cell>());
            for(int j = 0; j < num_col; j++)
            {
                cells[cells.Count - 1].Add(new Cell() { RowNumber = num_row + 1, ColumnLetter = Convert.ToChar('A' + j) });
            }
            dataGridView1.Rows.Add(1);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = (dataGridView1.Rows.Count).ToString();
            num_row++;
            //dataGridView1.Rows[num_row-1].Cells[0].Value = "a";
        }

        public void AddColumn()
        {
            //num_col = dataGridView1.Columns.Count;
            //num_row = dataGridView1.Rows.Count;
            if (num_col > 25)
            {
                num_col++;
                DataGridViewColumn col = new DataGridViewColumn();
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                col.CellTemplate = cell;
                int k = dataGridView1.ColumnCount - 1;
                string s = dataGridView1.Columns[k].Name;
                BaseSysFormat f2 = new BaseSysFormat();
                col.Name = f2.ToSys(num_col - 1);
                dataGridView1.Columns.Add(col);
                dataGridView1.Refresh();
            }
            else
            {
                for (int i = 0; i < num_row; i++) 
                {
                    cells[i].Add(new Cell() { RowNumber = i + 1, ColumnLetter = Convert.ToChar('A' + num_col) });
                }
                dataGridView1.Columns.Add(((char)('A' + num_col)).ToString(), ((char)('A' + num_col)).ToString());
                num_col++;
            }
        }
        public void RemoveRow()
        {
            num_row = dataGridView1.Rows.Count;
            //dataGridView1.Rows[num_row].Cells[0].Value = "a";
            dataGridView1.Rows.RemoveAt(num_row - 1);
            cells.RemoveAt(num_row - 1);
            for(int i = 0; i < num_row - 1; i++)
            {
                for(int j = 0; j < num_col; j++)
                {
                    if (cells[i][j].References.Where(a => a.RowNumber == num_row).Count() != 0)
                    {
                        ChangeData(cells[i][j].Expression, i, j);
                    }
                }
            }
            num_row--;
        }
        public void RemoveColumn()
        {
            //num_col = dataGridView1.ColumnCount;
            dataGridView1.Columns.RemoveAt(num_col - 1);
            for(int i = 0; i< num_row; i++)
            {
                cells[i].RemoveAt(num_col - 1);
            }
            for(int i = 0; i < num_row; i++)
            {
                for(int j = 0; j < num_col - 1; j++)
                {
                    if (cells[i][j].References.Where(a => a.ColumnLetter == 'A' + num_col - 1).Count() != 0)
                    {
                        ChangeData(cells[i][j].Expression, i, j);
                    }
                }
            }
            num_col--;
        }
        public void ChangeData(string expression, int row, int col)
        {
            try
            {
                cells[row][col].Expression = expression;
                cells[row][col].Value = parser.ExpressionStart(expression, cells[row][col]);
                cells[row][col].Error = null;
                Recalculate(cells[row][col]);
            }
            catch (ParserException ex)
            {
                cells[ex.row][ex.col].Error = ex.Message;
                dataGridView1.Rows[ex.row].Cells[ex.col].Value = cells[ex.row][ex.col].Error;
            }
        }
        void Recalculate(Cell cell)
        {
            for(int i = 0; i < num_row; i++)
            {
                for(int j = 0; j < num_col; j++)
                {
                    if (cells[i][j].Expression != null)
                    {
                        for(int  k = 0; k < cells[i][j].References.Count; k++)
                        {
                            if (cells[i][j].References[k].RowNumber == cell.RowNumber && cells[i][j].References[k].ColumnLetter == cell.ColumnLetter)
                            {
                                cells[i][j].Value = parser.ExpressionStart(cells[i][j].Expression, cells[i][j]);
                                cells[i][j].Error = null;
                                dataGridView1.Rows[i].Cells[j].Value = cells[i][j].Value.ToString();
                                Recalculate(cells[i][j]);
                            }
                        }
                    }
                }
            }
        }
        public void FillData(Mode mode)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            for (char i = 'A'; i < 'A' + num_col; i++)
                dataGridView1.Columns.Add(i.ToString(), i.ToString());

            dataGridView1.Rows.Add(num_row);

            for(int i = 0; i < num_row; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
                for(int j = 0; j < num_col; j++)
                {
                    if (cells[i][j].Expression != null)
                    {
                        if (cells[i][j].Error != null)
                            dataGridView1.Rows[i].Cells[j].Value = cells[i][j].Error.ToString();
                        else
                            dataGridView1.Rows[i].Cells[j].Value = mode == Mode.Expression ? cells[i][j].Expression.ToString() : cells[i][j].Value.ToString();   
                    }
                }
            }
        }
        public void SaveToFile(string path)
        {
            StreamWriter stream = new StreamWriter(path);
            stream.WriteLine(num_row);
            stream.WriteLine(num_col);
            for(int i = 0; i < num_row; i++)
                for(int j =0; j < num_col; j++)
                {
                    if (cells[i][j].Expression != null)
                    {
                        stream.WriteLine(i);
                        stream.WriteLine(j);
                        stream.WriteLine(cells[i][j].Expression);
                        stream.WriteLine(cells[i][j].Value);
                        if (cells[i][j].Error == null)
                            stream.WriteLine();
                        else
                            stream.WriteLine(cells[i][j].Error);
                    }
                }
            stream.Close();
        }
        public void OpenFile(string path)
        {
            StreamReader stream = new StreamReader(path);
            DataGridView file = new DataGridView();
            num_row = Convert.ToInt32(stream.ReadLine());
            num_col = Convert.ToInt32(stream.ReadLine());
            file.ColumnCount = num_col;
            file.RowCount = num_row;
            while (!stream.EndOfStream)
            {
                int i = Convert.ToInt32(stream.ReadLine());
                int j = Convert.ToInt32(stream.ReadLine());
                cells[i][j].Expression = stream.ReadLine();
                cells[i][j].Value = Convert.ToDouble(stream.ReadLine());
                string error = stream.ReadLine();
                if(!string.IsNullOrEmpty(error))
                {
                    cells[i][j].Error = error;
                }
            }
            stream.Close();
        }
    }
}
