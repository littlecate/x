using CellModelToPdfLib.Model;
using CSScriptLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CellModelToPdfLib.CellOp
{
    public class CalFormulaClass
    {
        CellOpClass op;
        public CalFormulaClass(CellOpClass _op)
        {
            op = _op;
        }

        private string GetCalFormulaCode(string expression)
        {
            string code = File.ReadAllText(@"G:\项目\CellModelToPdf\CellModelToPdfLib\CalFormula.cs");
            code = code.Replace("\"具体内容占位符\"", expression);
            return code;
        }

        public void CalFormula(Formula formula, int sheet)
        {
            List<string> L = Comman.拆分字串为词组(formula.str);
            List<string> L1 = new List<string>();
            List<string> L2 = new List<string>();
            for (var i = 0; i < L.Count; i++)
            {
                var p = L[i];
                if (i + 1 < L.Count && L[i + 1] == "!")
                {
                    L1.Add(p);
                    L2.Add("");
                    L1.Add("!");
                    L2.Add("");
                    i++;
                }
                else if (Comman.IsColRowMark(p))
                {
                    int col = 0, row = 0;
                    Comman.GetColRowFromStrMark(p, ref col, ref row);
                    if (row < 65535)
                    {
                        var t = op.GetCellString(col, row, sheet);
                        L1.Add(p);
                        L2.Add("\"" + t + "\"");
                    }
                }
                else if (p.ToLower() == "if")
                {
                    L1.Add(p);
                    L2.Add("o.myif");
                }
                else if (p.ToLower() == "value")
                {
                    L1.Add(p);
                    L2.Add("o.myvalue");
                }
                else if (p.ToLower() == "or")
                {
                    L1.Add(p);
                    L2.Add("||");
                }
                else if (p.ToLower() == "and")
                {
                    L1.Add(p);
                    L2.Add("&&");
                }
                else if (p.ToLower() == "string")
                {
                    L1.Add(p);
                    L2.Add("o.mystring");
                }
                else if (Regex.IsMatch(p, @"^[a-zA-Z]+?$"))
                {
                    L1.Add(p);
                    L2.Add("o." + p.ToLower());
                }
            }
            for (var m = 0; m < L1.Count; m++)
            {
                for (var i = 0; i < L.Count; i++)
                {
                    if (L[i] == L1[m])
                    {
                        L[i] = L2[m];
                    }
                }
            }
            CSScript.EvaluatorConfig.Engine = EvaluatorEngine.CodeDom;
            string code = GetCalFormulaCode(string.Join(" ", L));
            try
            {
                dynamic o = CSScript.Evaluator.LoadCode(code);
                string r = o.GetResult();
                op.S(formula.col, formula.row, formula.sheet, r);
            }
            catch (Exception ex)
            {
                op.S(formula.col, formula.row, formula.sheet, "");
            }
        }        
    }
}
