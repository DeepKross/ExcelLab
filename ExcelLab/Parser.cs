using ExcelLab;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace FileManager
{
    public class ParserException : ApplicationException
    {
        public int row;
        public int col;
        public ParserException(string str, int row, int col) : base(str)
        {
            this.row = row;
            this.col = col; 
        }
        public override string ToString()
        {
            return Message;
        }
    }
    public class Parser
    {
        const int NUMBEROFLETTERS = 26;
        const int ASCIIVALUEA = 65;
        enum Types { NONE, DELIMITER,VARIABLE, NUMBER };
        enum Errors { SYNTAX, UNBALPARENS, NOEXP, DIVBYZERO, RECURCELLS, REFWITHERROR, NOTEXISTCELL};
        string exp;
        int expInd;
        string token;
        Types type;

        Cell curCell;
        int Row_cur;
        int Col_cur;
        double[]vars = new double[NUMBEROFLETTERS];
        public Parser()
        {
            for(int i = 0; i < vars.Length; i++)
                vars[i] = 0.0;    
        }

        /*public Cell Cell
        {
            get => default(Cell);
            set { }
        }*/

        public double ExpressionStart(string expstr, Cell _curCell)
        {
            curCell = _curCell;
            int index = ASCIIVALUEA;
            Col_cur = _curCell.ColumnLetter - index;
            Row_cur = _curCell.RowNumber - 1;
            curCell.References.Clear();
            double result;
            exp = expstr;
            expInd = 0;
            GetToken();
            if (token == "")
            {
                SyntaxErr(Errors.NOEXP);
                return 0.0;
            }
            EvalExp1(out result);
            if (token != "")
                SyntaxErr(Errors.SYNTAX);
            if (IsRecur(curCell))
                SyntaxErr(Errors.RECURCELLS);
            return result;

        }

        bool IsRecur(Cell cell)
        {
            foreach(var i in cell.References)
            {
                if (i == curCell)
                    return true;
                if(IsRecur(i) == true)
                    return true;
            }
            return false;
        }

            void EvalExp1(out double result)
            {
                string op;
                double result1;
                EvalExp2(out result);
                while( (op = token) == ">" || op == "<" || op == "<=" || op == ">=" || op =="<>" || op == "=")
                {
                    GetToken();
                    EvalExp2(out result1);
                    switch (op)
                    {
                        case ">":
                            if (result > result1) result = 1.0;
                            else result = 0.0;
                            break;
                        case "<":
                            if (result < result1) result = 1.0;
                            else result = 0.0;
                            break;
                        case "<=":
                            if (result <= result1) result = 1.0;
                            else result = 0.0;
                            break;
                        case ">=":
                            if (result >= result1) result = 1.0;
                            else result = 0.0;
                            break;
                        case "<>":
                            if (result != result1) result = 1.0;
                            else result = 0.0;
                            break;
                        case "=":
                            if (result == result1) result = 1.0;
                            else result = 0.0;
                            break;
                    }
                }
                
            }
            void EvalExp2(out double result)
            {
                string op;
                double result1;

                EvalExp3(out result);
                while ((op = token) == "+" || op == "-")
                {
                    GetToken();
                    EvalExp3(out result1);
                    switch(op)
                    {
                        case "+":
                            result = result + result1;
                            break;
                        case "-":
                            result = result - result1;
                            break;
                    }
                }
            }

            void EvalExp3(out double result)
            {
                string op;
                double result1 = 0.0;
                EvalExp4(out result);
                while ((op = token) == "*" || op == "/" || op == "%" || op == "|")
                {
                    GetToken();
                    EvalExp4(out result1);
                    switch (op)
                    {
                        case "*":
                            result = result * result1;
                            break;
                        case "/":
                            if (result1 == 0.0)
                            {
                                SyntaxErr(Errors.DIVBYZERO);
                            }
                            else
                                result = result / result1;
                            break;
                        case "%":
                            if (result1 == 0.0)
                            {
                                SyntaxErr(Errors.DIVBYZERO);
                            }
                            else
                                result = result % result1;
                            break;
                        case "|":
                            if (result1 == 0.0)
                            {
                                SyntaxErr(Errors.DIVBYZERO);
                            }
                            else result = (int)result / (int)result1;
                            break;
                    }
                }
            }

            void EvalExp4(out double result)
            {
                double result1, ex;
                EvalExp5(out result);
                if (token == "^")
                {
                    GetToken();
                    EvalExp4(out result1);
                    ex = result;
                    if (result1 == 0.0)
                    {
                        result = 1.0;
                        return;
                    }
                    for(int i = (int)result1 - 1; i > 0; i--)
                    {
                        result = result * (double)ex;
                    }
                }
            }

            void EvalExp5(out double result)
            {
                string op;
                op = "";
                if ( (type == Types.DELIMITER) && token == "+" || token == "-")
                {
                    op = token;
                    GetToken();
                }
                EvalExp6(out result);
                if (op == "-") result = -result;
            }
            void EvalExp6(out double result)
            {
                if (token == "(")
                {
                    GetToken();
                    EvalExp1(out result);
                    if (token != ")")
                    {
                        SyntaxErr(Errors.UNBALPARENS);
                    }
                    GetToken();
                }
                else Atom(out result);
            }
            void Atom(out double result)
            {
                switch(type)
                {
                    case Types.NUMBER:
                        try
                        {
                            result = Double.Parse(token);
                        }
                        catch(FormatException)
                        {
                            result = 0.0;
                            SyntaxErr(Errors.SYNTAX);
                        }
                        GetToken();
                        return;
                    case Types.VARIABLE:
                        result = FindVar(token);
                        GetToken();
                        return;
                    default:
                        result = 0.0;
                        SyntaxErr(Errors.SYNTAX);
                        break;
                }
            }
        //giving back the value of exp
        double FindVar(string name)
        {
            if (!Char.IsLetter(name[0]))
            {
                SyntaxErr(Errors.SYNTAX);
                return 0.0;

            }
            return vars[Char.ToUpper(name[0]) - 'A'];
        }

        //giving back the token
        void PutBack()
        {
            for (int i = 0; i < token.Length; i++)
                expInd--;
        }
        void SyntaxErr(Errors error)
        {
            string[] err = {
                "Syntax Error",
                "Disbalanced parances",
                "No expression",
                "Dividing by ZERO",
                "Reccursion",
                "Invalid Cell",
                "No cell"
                            };
            throw new ParserException(err[(int)error], Row_cur, Col_cur);
        }

        void GetToken()
        {
            type = Types.NONE;
            token = "";
            if (expInd == exp.Length)
            {
               // MessageBox.Show(expInd.ToString());
               // MessageBox.Show(exp.Length.ToString());
                return;
            }
            while (expInd < exp.Length && Char.IsWhiteSpace(exp[expInd])) ++expInd;
            if (expInd == exp.Length) return;
            if (isDelim(exp[expInd]))
            {
                token += exp[expInd];
                expInd++;
                //if (token == "%" || token == "|") expInd++;
                if (token == "<" && (exp[expInd] == '=' || exp[expInd] == '>') || (token == ">" && exp[expInd] == '='))
                {
                    token += exp[expInd];
                    expInd++;
                }
                type = Types.DELIMITER;
            }
            else if (Char.IsDigit(exp[expInd]))
            {
                while (!isDelim(exp[expInd]))
                {
                    token += exp[expInd];
                    expInd++;
                    if (expInd >= exp.Length) break;
                }
                type = Types.NUMBER;
            }
            else if (Char.IsLetter(exp[expInd]))
            {
                while (!isDelim(exp[expInd]))
                {
                    token += exp[expInd];
                    expInd++;
                    if (expInd >= exp.Length) break;
                }
                char Column = token[0];
                string Row = token.Substring(1);
                int Row_ind;
                if (!int.TryParse(Row, out Row_ind))
                {
                    SyntaxErr(Errors.SYNTAX);
                }
                int row_ind = Row_ind - 1;
                int col_ind = (int)Column - 64 - 1;
                if (row_ind >= Data.cells.Count || col_ind >= Data.cells[row_ind].Count)
                {
                    Cell notExistCell = new Cell()
                    {
                        RowNumber = row_ind + 1,
                        ColumnLetter = Column
                    };
                    curCell.References.Add(notExistCell);
                    SyntaxErr(Errors.NOTEXISTCELL);
                }
                Cell parsedCell = Data.cells[row_ind][col_ind];
                curCell.References.Add(parsedCell);
                if (!string.IsNullOrEmpty(parsedCell.Error))
                    SyntaxErr(Errors.REFWITHERROR);
                token = parsedCell.Value.ToString();
                type = Types.NUMBER;
            }
        }
        bool isDelim(char c)
        {
            if ("+-/*%^|()<>=".IndexOf(c) != -1) return true;
            return false;
        }
        
    }

}
