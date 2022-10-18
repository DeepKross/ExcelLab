using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLab
{
    public class BaseSysFormat
    {
        const int NUMBEROFLETTERS = 26;
        const int N = 1000;
        public BaseSysFormat()
        {

        }
        public string ToSys(int i)
        {      
            int k = 0;
            int[] arr = new int[N];
            while (i > NUMBEROFLETTERS - 1)
            {
                arr[k] = i / NUMBEROFLETTERS - 1;
                k++;
                i = i % NUMBEROFLETTERS;
            }
            arr[k] = i;
            string res = "";
            for(int j = 0; j <= k; j++)
            {
                res = res + ( (char)('A' + arr[j])).ToString();
            }
            return res;
        }
        public int FromSys(string Header)
        {
            char[] arr = Header.ToCharArray();
            int l = arr.Length;
            int res = 0;
            for(int i = l - 2; i >= 0; i++)
            {
                res = res + (((int)arr[i] - (int)'A') + 1) * Convert.ToInt32(Math.Pow(NUMBEROFLETTERS, l - i - 1));
            }
            res = res + ((int)arr[l - 1] - (int)'A');
            return res;
        }
    }

    public class Cell
    {
        public string Expression { get; set; }
        public double Value { get; set; }
        public string Error { get; set; }
        public int RowNumber;
        public char ColumnLetter;
        public List<Cell> References { get; set; } = new List<Cell>();

        /*public BaseSysFormat BaseSysFormat
        {
            get => default(BaseSysFormat);
            set { }
        }*/
    }
}
