/**********************************************************************************
 * 
 *      PARSER / EVALUATOR TEST HARNESS 
 *
 * 
 *      Copyright (c) Barry Briggs
 *      All Rights Reserved 
 *      Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the 
 *      License. You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0 
 *  
 *      THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS OR IMPLIED, 
 *      INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE, 
 *      MERCHANTABLITY OR NON-INFRINGEMENT. 
 *  
 *      See the Apache 2 License for the specific language governing permissions and limitations under the License. 
 *      
 *      This is a general purpose Excel-style formula parser and evaluator. This
 *      application is intended as a test harness for the CloudSheet cloud-based
 *      computation engine. Certain features of CloudSheet such as cell references
 *      are stubbed out. 
 *      
 *      The Parser is a recursive descent style engine with precedence rules. It handles
 *      arbitrary arithmetic expressions as well as a variety of Excel-like functions,
 *      as listed below (see instance variable _funcs).  
 *      
 *      The key data structure is the "formentry" (formula entry) which
 *      holds atomic entries in the formula. The numberstack is a list of formentry's 
 *      and there may be multiple numberstacks corresponding to the recursion level. The
 *      operator stack behaves similarly but holds operators (e.g., +, -). There are 
 *      several features in the parser which are present (such as STARTPARSE and ENDPARSE) 
 *      primarily to make debugging easier. 
 *      
 *      The evaluator is similarly recursive; to fetch an argument value for a given
 *      function for example may require traversing multiple levels of recursion. 
 *      
 *      Some of the functions in CloudSheet which depend on the environment will not
 *      function in this test harness. 
 *      
 *      Lastly: this is a work in progress. The code has NOT been optimized and suffice it
 *      to say there are lots of opportunities. 
 *      
 * ********************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Data;
using System.Net;
using System.Runtime.Serialization;

/************************** v5 *******************/

/** OPEN ISSUES **/
/*
 * 1. FIXD      Expressions in functions with no parens e.g. =pmt(0.3/12,12*30,500000), or =IF(3>4,1,0) 
 * 2. OPEN      Error handling
 * 3. OPEN      Unaries and precedence for same 
 * 4. RFEN      Should clean up all the (int)TokenType and (int)SimpleOps stuff where possible
 * 5. OPEN      Logical expressions (3>4, 3>(A1*A2/A3), etc.) 
 */

namespace Parser
{
    public enum TokenTypes
    {
        NUMBER=0,
        CELLREF=1,
        OPERATOR=2,
        NAME=3,
        RANGE=4,
        SUBSTRING=5,
        FUNCTION=6,
        UNARY=7,
        PRECEDENCE=8,
        DATE=9,
        ARGSEP=10, 
        LPAREN=11,
        RPAREN=12,
        FUNC=13,
        COMPARISON=14,
        STARTPARSE=98,
        ENDPARSE=99
    }
    /* currently overloaded with both operators and functions */ 
    public enum SimpleOps
    {
        NULLOP=-1,
        ADD = 0,
        SUB = 1,
        MUL = 2,
        DIV = 3,
        POW=4, 
        SQRT = 5,
        ABS = 6,
        ACOS = 7,
        ASIN = 8,
        ATAN = 9,
        CEILING = 10,
        FLOOR = 11,
        COS = 12,
        COSH = 13,
        EXP = 14,
        LOG = 15,
        ROUND = 16,
        SIGN = 17,
        SIN = 18,
        SINH = 19,
        TAN = 20,
        TANH = 21,
        TRUNCATE = 22,
        SUM = 23,
        AVG=24,
        PI=25,
        STOCK=26,
        TODAY=27,
        DATE=28,
        POWER=29,
        DATA=30,
        GETDATAVAL=31,
        PUTDATAVAL=32,
        TIMEDGETDATAVAL=33,
        TIMEDPUTDATAVAL=34,
        DATASUM = 35,
        DATAAVG = 36,
        DATAMIN = 37,
        DATAMAX = 38,
        PMT=39,
        FV=40,
        MAX=41,
        MIN=42,
        IF=43,
        PV=44,
        NPV=45,
        LASTFUNC=46
    }
    #region Ranges
    public enum RangeTypes
    {
        VERTRANGE=0,
        HORIZRANGE=1,
        RECTRANGE=2,
        THREEDRANGE=3
    }

    public class Range
    {
        public Pt topleft;
        public Pt botright;
        public uint orientation;
        public ulong[] cells;
        public double[] values;
        public uint cellcount;

        public Range()
        {
            topleft.X = -1;
            topleft.Y = -1;
            botright.X = -1;
            botright.Y = -1;
        }
        public Range(Pt tl, Pt br)
        {
            topleft = tl;
            botright = br;
            int i = 0; 
            /* normalize the range */
            if (botright.Y < topleft.Y)
            {
                int s = topleft.Y;
                topleft.Y = botright.Y;
                botright.Y = s;
            }
            if (botright.X < topleft.X)
            {
                int s = topleft.X;
                topleft.X = botright.X;
                botright.X = s;
            }
            if (tl.X == br.X)
            {
                orientation = (uint)RangeTypes.VERTRANGE;
                cellcount = (uint)(botright.Y - topleft.Y) +1;
                cells = new ulong[cellcount];
                values = new double[cellcount];
                for (i = 0; i < cellcount; i++)
                {
                    ulong y = ((ulong)(topleft.Y + i)) << 32;
                    ulong cell = y + (ulong)topleft.X;
                    cells[i] = cell;
                }
            }
            else if (tl.Y == br.Y)
            {   
                orientation = (uint)RangeTypes.HORIZRANGE;
                cellcount = (uint)(botright.X - topleft.X) +1;
                cells = new ulong[cellcount];
                values = new double[cellcount];
                ulong y = ((ulong)(topleft.Y + i)) << 32;
                for (i = 0; i < cellcount; i++)
                {
                    ulong cell = y + (ulong)(topleft.X + i);
                    cells[i] = cell;
                }
            }
            else
            {
                int r = 0;
                int c =0;
                int j = 0;
                uint rowct = (uint)(botright.Y - topleft.Y) + 1;
                uint colct = (uint)(botright.X - topleft.X) + 1;
                orientation = (uint)RangeTypes.RECTRANGE;
                cellcount = rowct * colct; 
                cells = new ulong[cellcount];
                values = new double[cellcount];
                for (r = topleft.Y; r < topleft.Y+rowct; r++)
                {
                    for (c = topleft.X; c < topleft.X+colct; c++)
                    {
                        ulong y = (ulong)r << 32;
                        cells[j++] = (ulong)(y) + (ulong)(c);
                    }
                }

            }
        }
    }
    #endregion
    #region Formentry
    public class formentry : IEquatable <formentry>
    {
        public int tokentype;
        public double value;
        public int func; 
        public long ptr;
        public string text;
        public List<formentry> pushednumstack;
        public List<formentry> pushedopstack;
        public List<formentry> newnumstack;
        public List<formentry> newopstack; 
        public Pt celladdr;
        public Range range;

        public formentry()
        {
            tokentype=-1; 
        }
        /* create a new formentry of a given type */ 
        public formentry(int type)
        {
            tokentype = type; 
        }
        /* clone */ 
        public formentry(formentry f)
        {
            tokentype = f.tokentype;
            func = f.func; 
            value = f.value;
            ptr = f.ptr;
            text = f.text;
            pushednumstack = f.pushednumstack;
            pushedopstack = f.pushedopstack;
            newnumstack = f.newnumstack;
            newopstack = f.newopstack; 
        }
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            formentry f=obj as formentry;
            if (f == null)
                return false;
            else
                return Equals(f); 
        }
        public bool Equals(formentry f)
        {
            if (f.tokentype == tokentype)
                return true;
            else
                return false; 
        }
    }
    #endregion
    /* we use the custom Pt class as CloudSheet doesn't like the .NET provided Point class */ 
    public class Pt
    {
        public int X;
        public int Y;
        public Pt(int x, int y)
        {
            X = x;
            Y = y;
        }
    }
    public class ParseException : Exception
    {
        public ParseException(string message)
        {
        }
    }
    class Program
    {

        static void Main(string[] args)
        {
            Parser p = new Parser();
            while (true)
            {
                Console.WriteLine("Enter new formula"); 
                string text = Console.ReadLine();
                if (text.Length > 0)
                {
                    p.Init(text);
                    p.Parse(); 
                }

            }
        }
    }
    class Parser
    {
        string                                  _text;                         /* text we are parsing                    */
        int                                     _index;                        /* current pointer inside the text buffer */ 

        private string[] _funcs = new string[] 
                                                {"+","-","*","/","^","SQRT(","ABS(","ACOS(","ASIN(","ATAN(",     /* 0-9 */ 
                                                 "CEIL(","FLOOR(","COS(","COSH(","EXP(","LOG(","ROUND(","SIGN(",    /* 10-17 */
                                                 "SIN(","SINH(","TAN(","TANH(","TRUNC(", "SUM(", "AVG(","PI(", "STOCK(","TODAY(", "DATE(", "POWER(",
                                                 "DATA(","GETDATAVAL(", "PUTDATAVAL(", "TIMEDGETDATAVAL(", "TIMEDPUTDATAVAL(",
                                                 "DATASUM(","DATAAVG(","DATAMIN(","DATAMAX(",
                                                 "PMT(","FV(","MAX(","MIN(","IF(","PV(","NPV("};                              /* 18-22 */


        private char[] _unaries = new char[] { '+',                                           /* positive */
                                               '-',                                           /* negation */
                                               '~',                                           /* 1's complement */
                                               '!' };                                         /* logical not */

        private string[]                    _comparisons=new string [] {"=",">","<",">=","<=","<>"};             /* for comparisons in =IF */ 

        private ulong[]                     _powers26 = { 26, 676, 17576, 456976, 11881376, 308915776, 8031810176, 208827064576 };
        private int[]                       months = new int[12] { 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334 };


        private List<formentry>             _numberstack;                                     /* number stack */
        private List<formentry>             _opstack;                                         /* operator stack */

        private bool                        _expectingvalue = true;                           /* as we parse are we expecting a value
                                                                                               * or an operator; if expecting a value and
                                                                                               * get an op, it's probably a unary */

        private bool                        _expectingcomparison = false;                     /* this will be set when we encounter an =IF
                                                                                               * function and we are expecting a logical
                                                                                               * expression */

        private string                      _stockurl = "http://finance.yahoo.com/d/quotes.csv?f=l1&s=";

        private string[]                    _headings = new string[64];
        private string[,]                   _data;
        private int                         _datacols = 0;
        private int                         _datarows = 0;
        private string                      _errorstring = "";
        private delegate double             eval_func(formentry f);

        private List<int>                   _opcolumns = new List<int>(); 

        /* debugging purposes */ 
        private int                        _stackdepth = 0; 

        public Parser()
        {
        }
        public Parser(string t)
        {
            double result=double.NaN;
            _stackdepth = 0;
            this._text=t;
            if (_text[0] == '-')
                _index = 0;
            else
                _index = 1;                                                                     /* skip the + or = */
            _numberstack=new List<formentry>();
            _opstack=new List<formentry>();
            try
            {
                parse(_numberstack, _opstack);
                result = eval();
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error: " + exc.Message);
            }
            Console.WriteLine("Result = " + result.ToString());

        }
        public void Init(string t)
        {

            this._text = t;

            _index = 1;                                                                     /* skip the + or = */

            if (_numberstack == null)
                _numberstack = new List<formentry>();
            else
                _numberstack.Clear();
            if (_opstack == null)
                _opstack = new List<formentry>();
            else
                _opstack.Clear(); 
        }

        public void Parse()
        {
            double result = double.NaN;
            _stackdepth = 0;
            try
            {
                parse(_numberstack, _opstack);
                result = eval();
            }
            catch (Exception exc)
            {
                Console.WriteLine("Error: " + exc.Message);
            }
            Console.WriteLine("Result = " + result.ToString());
        } 
        /* parse a string until parens or higher precedence found -- recursive
         * 
         * 
         */
        private void parse(List<formentry> numstack, List<formentry> opstack)
        {

            while (_index < _text.Length)
            {
                formentry tok = gettoken();
                if(tok!=null)
                {
                    /* parens indicate precedence level changes +1 */
                    if(tok.tokentype==(int)TokenTypes.LPAREN)
                    {
                        _expectingvalue = true; 
                        formentry f=push(numstack, opstack); 
                        parse(f.newnumstack, f.newopstack);
                        continue; 
                    }
   
                    /* note we aren't keeping track of the # of parens */
                    if (tok.tokentype == (int)TokenTypes.RPAREN)
                    {
                        pop(numstack);
                        if (numstack.Count > 0 && numstack[0].tokentype == (int)TokenTypes.STARTPARSE)
                        {
                            opstack = numstack[0].pushedopstack; // must do first
                            numstack = numstack[0].pushednumstack;
                        }
                        _expectingvalue = false; //TEST
                        continue;
                    }
                    /* argument separator within a function */ 
                    if (tok.tokentype == (int)TokenTypes.ARGSEP)
                    {
                        /* this hairball handles =pmt(12*30, 0.05/12, 23234) without internal parens; force creation of precedence after the fact. 
                         * 
                         * essentially the idea here is to move the expression into its own recursion level and replace the current level with 
                         * a precedence indicator; a sort of 'promotion' of the expression.
                         */ 

                        /* step 0. copy everything from the last argument separator */    
                        int sep = numstack.Count;
                            int a; 
                            for (a = sep - 1; a >= 0; a--)
                                if (numstack[a].tokentype == (int)TokenTypes.ARGSEP)
                                    break;
                            a++;

                            /* step 1. copy numstack from the last argument separator so we can move it to the new recursion level */
                            List<formentry> templist = new List<formentry>();
                            int b;
                            for (b = a; b < numstack.Count;b++ )
                            {
                                formentry f = numstack[b]; 
                                formentry f2 = new formentry(f);
                                templist.Add(f2);
                            }
                            if (templist.Count > 1)
                            {
                                numstack.RemoveRange(a, numstack.Count - a);

                                if (numstack.Count == 0 && templist[0].tokentype == (int)TokenTypes.STARTPARSE)
                                {
                                    formentry ff = new formentry((int)TokenTypes.STARTPARSE);
                                    ff.pushednumstack = templist[0].pushednumstack;
                                    ff.pushedopstack = templist[0].pushedopstack;
                                    numstack.Add(ff);
                                }


                                /* step 2. copy opstack so we can move it as well to the new recursion level */
                                List<formentry> tempoplist = new List<formentry>();
                                foreach (formentry fop in opstack)
                                {
                                    formentry fop2 = new formentry(fop);
                                    tempoplist.Add(fop2);
                                }

                                /* step 3. clear the opstack (this is problematic) */
                                opstack.Clear();

                                /* step 4. create the new numstack and create the linkages between the levels; this creates the 
                                 * PRECEDENCE formentry in current level and the STARTPARSE in the new level. p here is that 
                                 * formentry */

                                formentry p = push(numstack, opstack);

                                /* step 5. in the new numstack, put the terms of this expression */
                                /* there is a STARTPARSE already created */

                                for (int i = (a == 0) ? 1 : 0; i < templist.Count; i++)
                                {
                                    formentry f3 = new formentry(templist[i]);
                                    p.newnumstack.Add(f3);
                                }
                                /* Step 6. decrement stack depth and close off the new stack with an ENDPARSE */
                                pop(p.newnumstack);

                                /* Step 7. Promote the opstack to the new level and pretend nothing happened. */
                                foreach (formentry f5 in tempoplist)
                                {
                                    p.newopstack.Add(f5);
                                }
                            }

                        /* argument separator on the original numstack */ 
                        numstack.Add(new formentry((int)TokenTypes.ARGSEP));
                        _expectingvalue = true;
                        continue; 
                    }
                    if(tok.tokentype==(int)TokenTypes.COMPARISON)
                    {
                        if (_expectingcomparison)
                        {
                            opstack.Add(tok);
                            _expectingcomparison = false;
                        }
                        //need an 'else' here to handle the error condition 
                        continue; 
                    }

                    if (tok.tokentype == (int)TokenTypes.FUNCTION)
                    {
                        _expectingvalue = true;    
                        formentry f = push(numstack, opstack);
                        f.tokentype = (int)TokenTypes.FUNC;
                        f.func = (int)tok.value;
                        parse(f.newnumstack, f.newopstack); 
                        _expectingvalue = true;                    /* ????? */
                        continue; 
                    }

                    if (_expectingvalue == true && tok.tokentype == (int)TokenTypes.UNARY)
                    {
                        formentry funary = push(numstack, opstack, null, tok);
                        parse(funary.newnumstack, funary.newopstack);
                        _expectingvalue = true;
                        continue; 
                    }

                    else if (tok.tokentype == (int)TokenTypes.OPERATOR || tok.tokentype == (int)TokenTypes.UNARY)
                    {
                        /* see if it is really a unary or actually an operator */
                        if (tok.tokentype == (int)TokenTypes.UNARY && _expectingvalue == false)
                            tok.tokentype = (int)TokenTypes.OPERATOR;
                        /* here handle the case of 1+2*3 where we infer precedence */
//                        if (tok.tokentype == (int)TokenTypes.OPERATOR && tok.value >= 2 && opstack[opstack.Count - 1].value < 2)
                        if (tok.tokentype==(int)TokenTypes.OPERATOR && opstack.Count != 0 && tok.value >= 2 && opstack[opstack.Count - 1].value < 2)
                        {
                            /* create precedence and recurse */
                            formentry numsave = numstack[numstack.Count - 1];
                            formentry numnew = new formentry(numsave);
                            numstack.Remove(numsave);
                            formentry opsave = opstack[opstack.Count - 1];
                            formentry opnew = new formentry(opsave);
                            opstack.Remove(opsave);
                            /* MUST make sure this gets popped */ 
                            formentry pf = push(numstack, opstack, numnew, tok);
                            parse(pf.newnumstack, pf.newopstack);
                            /* previous numstack/opstack restored on return */
                            opstack.Add(opnew);
                            _expectingvalue = true;
                        }

                        else
                        {
                            opstack.Add(tok);
                            _expectingvalue = true;
                        }
                    } 

                    else
                    {
                        numstack.Add(tok);
                        _expectingvalue = false;
                    } /* is number */
                }
                else
                    return;

                if (_index >= _text.Length)
                    return; 
            }
            return; 
         }
        /* get the next token and return a token struct
         * we use exceptions here to make the code flow easier */
        private formentry gettoken()
        {
            int unary; 
            int comp; 
            int op;
            string s = ""; 
            formentry f = new formentry();

            while (_index < _text.Length && _text[_index] == 0x20)
                _index++;

            if (ScanForArgSep(f) == true)
            {
                 _index++;
                 return f;
            }
            if (ScanForParen(f) == true)
            {
                _index++;
                return f;
            } 
   
            /* unary operators have the highest precedence */ 
            if ((unary=ScanForUnary()) != -1)
            {
                f.tokentype = (int)TokenTypes.UNARY;
                f.value = unary;
                _index++;
                return f; 
            }
            if((comp=ScanForComparison())!=-1)
            {
                f.tokentype = (int)TokenTypes.COMPARISON;
                f.value = comp;
                _index++;
                return f; 
            }
            /* now see if properly formatted date; must be nn/nn/nnnn */
            try
            {
                string date = ScanForDate(f);
                Console.WriteLine("Found date " + date);
                _index++;
                return f;
            }
            catch (Exception exc)
            {
            }

            try
            {
                double num = ScanForNumber(f);
                return f;
            }
            catch (Exception exc)
            {
            }

            try 
            {
                Pt pt=ScanForCellAddress(f);
                if (_index < _text.Length && _text[_index] == ':')
                {
                    _index++;
                    formentry f1 = new formentry();
                    Pt pt2 = ScanForCellAddress(f1);
                    f.tokentype = (int)TokenTypes.RANGE;
                    Range r = new Range(pt, pt2);
                    f.range = r;
                }
                return f;
            }
            catch (Exception exc)
            {
            }

            op = ScanForFunction(f);
            if (op >= 0)
                return f;

            /* we'll consider it a string for now; find end of string */
            s = ScanForString(f);
            if (s != null || s != "")
                 return f;
            else
                return null;
        }
        #region Stack Push and Pop
        private formentry push(List<formentry> oldnumstack, List<formentry> oldopstack)
        {
            return (push(oldnumstack, oldopstack, null, null));
        }

        private formentry push(List<formentry> oldnumstack, List<formentry> oldopstack, formentry num, formentry op)
        {
            _stackdepth++; 
            List<formentry> newnumstack = new List<formentry>();
            List<formentry> newopstack = new List<formentry>();
            formentry p = new formentry();
            /* tells evaluator to go to another stack */
            p.tokentype = (int)TokenTypes.PRECEDENCE;
            /* when we evaluate we will switch over to this stack */
            p.newnumstack = newnumstack;
            p.newopstack = newopstack;
            p.pushednumstack = oldnumstack;
            p.pushedopstack = oldopstack;
            oldnumstack.Add(p);
            if (op != null)
            { /* happens we already found a function */
                newopstack.Add(op);
                op.newnumstack = newnumstack; /* operator points to its arguments */
            }

            formentry fnew = new formentry();
            fnew.tokentype = (int)TokenTypes.STARTPARSE;
            fnew.pushednumstack = oldnumstack;
            fnew.pushedopstack = oldopstack;
            newnumstack.Add(fnew); 

            if (num != null)
                newnumstack.Add(num);
            return p;
        }
        private void pop(List<formentry> stack)
        {
            _stackdepth--; 
            /* useful for debugging */
            formentry p = new formentry();
            p.tokentype = (int)TokenTypes.ENDPARSE;
            stack.Add(p);
            return; 
        }
        /* fills in pushed number stack and operator stack entries in case */
        private void addformentry(formentry f, List<formentry> numstack)
        {
            if (numstack.Count > 0)
            {
                f.pushednumstack = numstack[0].pushednumstack;
                f.pushedopstack = numstack[0].pushedopstack;
                numstack.Add(f);
            }
            return; 
        }
        #endregion
        #region Token Evaluators
        /* date format n/n/nnnn; n/nn/nnnn; or nn/nn/nnnn */ 
        private string ScanForDate(formentry f)
        {

            int i=_index; 
            int date=0; 
            int day=0;
            int mon=0;
            int yr=0; 

            try
            {
                if(IsNum(_text[i]))
                    mon=_text[i++]-0x30;
                else
                    throw (new ParseException("not a date"));
                if(IsNum(_text[i]))
                    mon=mon*10+_text[i++]-0x30; 
                if(_text[i]=='/')
                    i++;
                else
                    throw (new ParseException("not a date"));
                if(IsNum(_text[i]))
                    day=_text[i++]-0x30;
                else
                    throw (new ParseException("not a date"));
                if(IsNum(_text[i]))
                    day=day*10+_text[i++]-0x30; 
                if(_text[i]=='/')
                    i++;
                else
                    throw (new ParseException("not a date"));

                if (IsNum(_text[i]) && IsNum(_text[i + 1]) && IsNum(_text[i + 2]) && IsNum(_text[i + 3]))
                {
                    yr = (_text[i] - 0x30) * 1000 + (_text[i + 1] - 0x30) * 100 + (_text[i + 2] - 0x30) * 10 + (_text[i + 3] - 0x30);
                    if (yr < 1900)
                        yr = 1900;
                    if (mon < 1 || mon > 12)
                        throw (new ParseException("not a date"));
                    if (day > 31)
                        throw (new ParseException("not a date"));

                    date = ((yr - 1900)*365) + ((yr-1900)/4 + 1) + months[mon - 1] + day;
                }
                else
                    throw (new ParseException("not a date"));
 
                f.text=_text.Substring(_index,(i+4)-_index);
                f.tokentype = (int)TokenTypes.DATE;
                f.value = (double)date; 
                _index =i+4;
                return f.text;
            }
            /* get here if we run out of buffer */ 
            catch
            {
                throw (new ParseException("not a date"));
            }
        }
        private bool ScanForArgSep(formentry f)
        {
            if (_text[_index] == ',')
            {
                f.tokentype = (int)TokenTypes.ARGSEP;
                return true;
            }
            else
                return false; 
        }
        private bool ScanForParen(formentry f)
        {
            if (_text[_index] == '(')
            {
                f.tokentype = (int)TokenTypes.LPAREN;
                return true;
            }
            else if (_text[_index] == ')')
            {
                f.tokentype = (int)TokenTypes.RPAREN;
                return true;
            }
            else
                return false;
        }

        private double ScanForNumber(formentry f)
        {
            double n = 0;
            string s = "";
            int isav = _index; /* in case of error */
            if (_text[_index] >= '0' && _text[_index] <= '9')
            {

                while (_text[_index] >= '0' && _text[_index] <= '9')
                {
                    s += _text[_index];
                    n = n * 10;
                    n += (byte)(_text[_index]) - 0x30;
                    _index++;
                    if (_index >= _text.Length)
                        break;
                }
                if (_index < _text.Length)
                {
                    if (_text[_index] == '.')
                    {
                        s += _text[_index];
                        double mul = 0.1;
                        _index++;
                        if (_text[_index] >= '0' && _text[_index] <= '9')
                        {
                            while (_text[_index] >= '0' && _text[_index] <= '9')
                            {
                                int r = (byte)(_text[_index]) - 0x30;
                                s += _text[_index];
                                n += (double)r * mul;
                                mul = mul / 10;
                                _index++;
                                if (_index >= _text.Length)
                                    break;
                            }
                        }
                    }
                    else
                    {
                        /* end of token delimiters -- either another argument or a close parentheses */ 
                        if (_text[_index] == ',' || _text[_index] == ')' || _text[_index]==0x20 || IsOperator(_text[_index]) || IsComparator(_text[_index]))
                        {
                            f.tokentype = (int)TokenTypes.NUMBER;
                            f.value = n;
                            f.text = s;
                            return n;
                        }
                        /* otherwise it's a string, probably */ 
                        else
                        {
                            _index = isav;
                            throw (new ParseException("not a number"));
                        }

                    }
                }

                f.tokentype = (int)TokenTypes.NUMBER;
                f.value = n;
                f.text = s;
                return n;
            }
            else
            {
                _index = isav;
                throw (new ParseException("not a number"));
            }
      
        }
        /* clone of above, sort of -- used by various functions */ 


        private int ScanForFunction(formentry f)
        {
            int i = 0;
            int last = (int)(SimpleOps.LASTFUNC);

            for (i = 0; i < last; i++)
            {
                string func = _funcs[i];
                if (func.Length >= _text.Length - _index)
                    continue;

                string g = _text.Substring(_index, func.Length);

                if (g.IndexOf(func, StringComparison.OrdinalIgnoreCase)==0 && g.Length==func.Length)
                //if (func.Equals(g))
                {
                    _index += func.Length;
                    f.text = g;
                    f.value = i;
                    if (i < 5)
                        f.tokentype = (int)TokenTypes.OPERATOR;
                    else
                    {
                        f.tokentype = (int)TokenTypes.FUNCTION;
                        if (i == (int)SimpleOps.IF)
                            _expectingcomparison = true;
                    }
                    return i;
                }
            }
            return -1;
        }
        /* strings can be any character data; scan until next operator, next quote, or next right paren
         * a string is declared when all other possibilities have been ruled out*/
        private string ScanForString(formentry f)
        {
            string s="";

            /* if first character is a quote, then it must be closed */
            if (_text[_index] == '"')
            {
                int i = _index + 1;
                while (i < _text.Length)
                {
                    if (_text[i] == '"')
                    {
                        f.tokentype = (int)TokenTypes.SUBSTRING;
                        f.text = _text.Substring(_index + 1, (i-1) - _index);
                        _index = i + 1;
                        return f.text;
                    }
                    i++;
                }
            }

            while (_index < _text.Length)
            {
                if (_text[_index] == ')' ||
                    _text[_index]==','  ||
                    _text[_index] == '"' )
 
#if false
                    _text[_index] == _funcs[0][0] ||
                    _text[_index] == _funcs[1][0] ||
                    _text[_index] == _funcs[2][0] ||
                    _text[_index] == _funcs[3][0] ||
                    _text[_index] == _funcs[4][0])
#endif
                {
                    f.tokentype = (int)TokenTypes.SUBSTRING;
                    f.text = s;
                    return s;
                }
                s += _text[_index];
                _index++;
            }
            f.tokentype = (int)TokenTypes.SUBSTRING;
            f.text = s;
            return s;
        }

        private Pt ScanForCellAddress(formentry f)
        {
            Pt pt = new Pt(-1, -1);
            int r = 0;
            int c = 0;
            int i = 0;
            bool syntaxerror = true;
            int index = _index;         /* if there is a parse error we cannot change the value of _index */ 

            /* supports column A through ZZ; to change just increase the maximum value of i */
            while (i < 2)
            {
                if ((_text[index] >= 'A' && _text[index] <= 'Z') || (_text[index] >= 'a' && _text[index] <= 'z'))
                {
                    syntaxerror = false;
                    c = i * 26 + (_text[index] <= 'Z' ? _text[index] - 0x41 : _text[index] - 0x61);
                    index++;
                    if (index >= _text.Length) break;
                }
                i++;
            }
            /* syntax error */
            if (syntaxerror)
                throw (new ParseException("not a cell"));
            else
                pt.X = c;

            syntaxerror = true;
            i = 0;
            /* supports up to row 99999999 */
            while (i < 8)
            {
                /* potential order of evaluation issue */
                while (index < _text.Length && _text[index] >= '0' && _text[index] <= '9')
                {
                    char a = _text[index];
                    syntaxerror = false;
                    r = r * 10 + (_text[index++] - 0x30);
                    i++;
                }
                break;
            }
            if (syntaxerror == true)
                throw (new ParseException("not a cell"));
            else
        
                pt.Y = r - 1;

            f.tokentype = (int)TokenTypes.CELLREF;
            f.celladdr = pt;
            _index = index;                         /* point past the cell address */ 

            return pt;
        }
        /* doesn't advance _index?????? */
        private int ScanForUnary()
        {
            int i;
            for (i = 0; i < 4; i++)
            {
                if (_text[_index] == _unaries[i])
                {
                    return i;
                }
            }
            return -1; 
        }
        private int ScanForComparison()
        {
            int i;
            int lim;

            /* check to make sure we don't compare past the end of the buffer */
            if (_text.Length - _index > 1)
                lim = _comparisons.Length;
            else
                lim = 3;

            for(i=0;i<lim;i++)
            {
                if (_text.Substring(_index, _comparisons[i].Length) == _comparisons[i])
                {
                    _index += _comparisons[i].Length-1;
                    return i;
                }
            }
            return -1; 
        }
        private string NumToCol(ulong num)
        {
            string s = null;
            ulong z = num;
            ulong r = num;
            int i=0;

            while (num >= 26)
            {
                while (num >= _powers26[i])
                    i++;
                r = num / _powers26[i - 1];
                s += (char)(r + 0x40);
                num = num - (r * _powers26[i-1]);
                i = 0;
            }
            s += (char)(num + 0x41);
            return s;

        }
        /* used by various functions; needs work   */ 
        private double StringToNum(string s)
        {
            double n = 0;
            int i = 0;
            if (s[i] >= '0' && s[i] <= '9')
            {

                while (s[i] >= '0' && s[i] <= '9')
                {
                    n = n * 10;
                    n += (byte)(s[i]) - 0x30;
                    i++;
                    if (i >= s.Length)
                        break;
                }
                if (i < s.Length)
                {
                    if (s[i] == '.')
                    {
                        double mul = 0.1;
                        i++;
                        if (s[i] >= '0' && s[i] <= '9')
                        {
                            while (s[i] >= '0' && s[i] <= '9')
                            {
                                int r = (byte)(s[i]) - 0x30;
                                n += (double)r * mul;
                                mul = mul / 10;
                                i++;
                                if (i >= s.Length)
                                    break;
                            }
                        }
                    }
                }
            }
            else
            {
                n = Double.NaN;
            }

            return n; 
        }
        /* takes a string and attempts to convert it to a Point coordinate. If it does not parse, one or the
         * other coordinates will be -1; this function is duplicated (slightly modified) from the one in CellGrain and is used
         * for XML file parsing where we have string cell addresses  */
        private Pt AddressToCell(string input)
        {
            Pt pt = new Pt(-1, -1);
            int r = 0;
            int c = 0;
            int i = 0;
            int bufptr = 0;
            bool syntaxerror = true;
            /* supports column A through ZZ; to change just increase the maximum value of i */
            while (i < 2)
            {
                if ((input[bufptr] >= 'A' && input[bufptr] <= 'Z') || (input[bufptr] >= 'a' && input[bufptr] <= 'z'))
                {
                    syntaxerror = false;
                    int n=(input[bufptr] <= 'Z' ? input[bufptr] - 0x40 : input[bufptr] - 0x60);
                    c = ((c+1) * (i * 26)) + n - 1; 
                    bufptr++;
                    if (bufptr >= input.Length) break;
                }
                else
                {
                    break;
                }
               
                i++;
            }
            /* syntax error */
            if (syntaxerror)
                return pt;
            else
                pt.X = c;

            syntaxerror = true;
            i = 0;
            /* supports up to row 99999999 */
            while (i < 8)
            {
                /* potential order of evaluation issue */
                while (bufptr < input.Length && input[bufptr] >= '0' && input[bufptr] <= '9')
                {
                    char a = input[bufptr];
                    syntaxerror = false;
                    r = r * 10 + (input[bufptr++] - 0x30);
                    i++;
                }
                break;
            }
            if (syntaxerror == true)
                return pt;
            else
                pt.Y = r - 1;

            return pt;
        }

        private bool IsOperator(char c)
        {
            return (c == '+' || c == '-' || c == '/' || c == '*' || c == '~' || c == '!');
        }
        private bool IsComparator(char c)
        {
            /* covers '<=', '>=', and '<>' as well */ 
            return (c == '=' || c == '<' || c == '>'); 
        }
        private bool IsNum(char c)
        {
            return (c >= '0' && c <= '9');
        } 
        #endregion
        #region Evaluation
        /* evaluator */

        private double eval()
        {
            Console.WriteLine("Evaluating ... Stack depth= "+_stackdepth.ToString()); 
            double res=0; 
            res = eval_worker(_numberstack, _opstack);
            return res; 
        }

        private double eval_worker(List<formentry> numstack, List<formentry> opstack)
        {
            int p = 0;  /* index into numstack */
            int q = 0;  /* index into opstack */ 
            
            double result=0.0; /* running result */

            formentry f = numstack[p++];

            RemoveArgSeps(f); 

            if (f.tokentype == (int)TokenTypes.STARTPARSE)
                f = numstack[p++];

            /* we always get the first value in case there are no operators, e.g., a simple number
             * or a standalone function like =sum(3,4) or =pi().
             */
            result = getvalue(f); 

            /* loop through the operators */
            while (q < opstack.Count)
            {
                int theoperator=-1; 

                formentry op = opstack[q++];
                theoperator = (int)op.value; 
                switch (theoperator)
                {
                    case (int)SimpleOps.ADD:
                        {
                            if (op.tokentype == (int)TokenTypes.UNARY)
                            {
                                result = Math.Abs(f.value);
                            }
                            else
                            {
                                formentry f1 = numstack[p++];
                                result += (double)getvalue(f1);
                            }
                            break;
                        }
                    case (int)SimpleOps.SUB:
                        {
                            if (op.tokentype == (int)TokenTypes.UNARY)
                            {
                                result = f.value * (-1);
                            }
                            else
                            {
                                formentry f1 = numstack[p++];
                                result -= (double)getvalue(f1);
                            }
                            break;
                        }
                    case (int)SimpleOps.MUL:
                        {
                            formentry f1 = numstack[p++];
                            result *= (double)getvalue(f1);
                            break;
                        }
                    case (int)SimpleOps.DIV:
                        {
                            formentry f1 = numstack[p++];
                            result /=getvalue(f1);
                            break;
                        }

                }
                
            }
            return result; 
        }
        /* a function is effectively a term */ 
        private double evaluate_function(formentry f)
        {
            double result = 0.0;

            RemoveArgSeps(f);


            switch(f.func) {
                    case (int)SimpleOps.POW:
                    case (int)SimpleOps.POWER:
                        {
                            formentry f0 = getarg(f, 0); 
                            formentry f1 = getarg(f,1);                       
                            result = Math.Pow((double)getvalue(f0), (double)getvalue(f1)); 
                            break; 
                        }
                    case (int)SimpleOps.SQRT:
                        {
                            result = (double)Math.Sqrt(getvalue(getarg(f,0))); 
                            break;
                        }
                    case (int)SimpleOps.ABS:
                        result = (double)Math.Abs(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.ACOS:
                        result = (double)Math.Acos(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.ASIN:
                        result = (double)Math.Asin(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.ATAN:
                        result = (double)Math.Atan(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.CEILING:
                        result = (double)Math.Ceiling(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.FLOOR:
                        result = (double)Math.Floor(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.COS:
                        result = (double)Math.Cos(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.COSH:
                        result = (double)Math.Cosh(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.EXP:
                        result = (double)Math.Exp(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.LOG:
                        result = (double)Math.Log(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.ROUND:
                        result = (double)Math.Round(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.SIGN:
                        result = (double)Math.Sign(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.SIN:
                        result = (double)Math.Sin(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.SINH:
                        result = (double)Math.Sinh(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.TAN:
                        result = (double)Math.Tan(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.TANH:
                        result = (double)Math.Tanh(getvalue(getarg(f, 0)));
                        break;
                    case (int)SimpleOps.TRUNCATE:
                        result = (double)Math.Truncate(getvalue(f));
                        break;
                    case (int)SimpleOps.SUM: 
                        {
                            int c = 0;
                            formentry fe;
                            eval_func e = sum_range;
                            while ((fe = getarg(f, c)) != null)
                            {
                                result += getvalue(fe, e);
                                c++; /*  nice name for a language */ 
                            }
                            //f.value = result; 
                        break;
                        }
                    case (int)SimpleOps.AVG:
                        {
                            int c = 0;
                            formentry fe;
                            eval_func e = avg_range;
                            while ((fe = getarg(f, c)) != null)
                            {
                                result += getvalue(fe, e);
                                c++; /*  nice name for a language */
                            }
                            result = result / c;
                            break;
                        }
                    case (int)SimpleOps.PI:
                        result = 3.141592654;
                        break;
                    case (int)SimpleOps.STOCK:
                        {
                            string s = GetSimpleString(getarg(f, 0));      /* get the stock symbol */
                            if (s == null || s == "")
                            {
                                result = double.NaN;
                                break;
                            }
                            else
                            {

                                WebClient wc = new WebClient();
                                string url = _stockurl + s;
                                wc.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                                Stream data = wc.OpenRead(url);
                                StreamReader reader = new StreamReader(data);
                                string res = reader.ReadToEnd();
                                Console.WriteLine(res);
                                data.Close();
                                reader.Close();
                                result = StringToNum(res);
                                break;
                            }
                        }
                    case (int)SimpleOps.TODAY:
                        {
                            DateTime date = DateTime.Today;
                            int y = date.Year;
                            int m = date.Month;
                            int d = date.Day;
                            int subtotal = (y - 1900) * 365;
                            subtotal += (y - 1900) / 4 + 1; /* leap days */
                            subtotal += months[m - 1]; // 0-based
                            subtotal += d;
                            result = (double)subtotal;
                            break;
                        }
                    case (int)SimpleOps.DATE:
                        {
                            formentry f0 = getarg(f, 0);
                            formentry f1 = getarg(f, 1);
                            formentry f2 = getarg(f, 2);
                            int y = (int)getvalue(f0);
                            int m = (int)getvalue(f1);
                            int d = (int)getvalue(f2);
                            if(y<0 || m<0||d<0)
                                throw new ParseException("Syntax Error"); 

                            int subtotal = (y - 1900) * 365;
                            subtotal += (y - 1900) / 4 + 1; /* leap days */
                            subtotal += months[m - 1]; // 0-based
                            subtotal += d;
                            result = (double)subtotal;
                            break;
                        }
                    /* load a block of data from a blob into a 2d table. assumed to be in CSV format */
                    case (int)SimpleOps.DATA:
                        {

                            int cells=0;
                            formentry f0 = getarg(f, 0);

                            string s = GetSimpleString(f0);      /* get the blob name */
                            if (s == null || s == "")
                            {
                                result = double.NaN;
                            }
                            else
                            {
                                try
                                    {
                                        cells = LoadDataBlob(s);
                                        result = (double)(cells);
                                    }
                                    catch (Exception exc1)
                                    {
                                        _errorstring += "Could not load data file";
                                        Console.WriteLine("Error: " + exc1.Message); 
                                    }

                                } 
                            break;
                        }


                    /* syntax: =GETDATAVAL(cell, keycol, key, col) where 'cell' is a =DATA cell anything else is an error
                     * ex: =GETDATAVAL(A10, 3, 2342, 4)       where we are looking up a value based upon a date */
                    case (int)SimpleOps.GETDATAVAL:
                        {
                            string ans=""; 
                            formentry f0 = getarg(f,0); /* cell address */
                            formentry f1 = getarg(f,1); /* key column */
                            formentry f2 = getarg(f,2); /* key */
                            formentry f3 = getarg(f,3); /* column of value to return */

                            Pt celladdr = f0.celladdr;

                            int i = 0;
                            int keycol = (int)f1.value;
                            for (i = 0; i < _datarows; i++)

                                if (_data[i, keycol] == f2.text)
                                {
                                    ans = _data[i, (int)f3.value]; /* some thing like this */
                                    return 1.0;
                                }

                            result= 0.0; 
                            break;
                        }
                    case (int)SimpleOps.PUTDATAVAL:
                        {
                            break;
                        }
                    case (int)SimpleOps.TIMEDGETDATAVAL:
                        {
                            break;
                        }
                    case (int)SimpleOps.TIMEDPUTDATAVAL:
                        {
                            break;
                        }
                    case (int)SimpleOps.PMT:
                        {
                            // p=(rate*principle)/1-(1+r)^(-nperiods)
                            List<formentry> nstack = f.newnumstack;
                            formentry f0 = getarg(f, 0);
                            formentry f1 = getarg(f, 1);
                            formentry f2 = getarg(f, 2);
                            double rate = getvalue(f0);
                            double nperiods = getvalue(f1);
                            double principle = getvalue(f2);
                            double pmt = (rate * principle) / (1 - Math.Pow((1 + rate), -nperiods));
                            return pmt;
                        }
                        
                    case (int)SimpleOps.FV:
                        {
                            List<formentry> nstack = f.newnumstack;
                            formentry f0 = getarg(f, 0);
                            formentry f1 = getarg(f, 1);
                            formentry f2 = getarg(f, 2);
                            double rate = f0.value;
                            double nperiods = f1.value;
                            double payment = f2.value;
                            double futurevalue = payment * ((Math.Pow((1 + rate), nperiods) - 1 )/ rate);
                            return futurevalue; 

                        }
                    case (int)SimpleOps.PV:
                        {
                            List<formentry> nstack = f.newnumstack;
                            formentry f0 = getarg(f, 0);
                            formentry f1 = getarg(f, 1);
                            formentry f2 = getarg(f, 2);
                            double rate = f0.value;
                            double nperiods = f1.value;
                            double payment = f2.value;
                            double presentvalue = payment / Math.Pow((1 + rate), nperiods);
                            return presentvalue;
                        }
                    case (int)SimpleOps.NPV:
                        {
                            return 0.0;
                        }
                    case (int)SimpleOps.MAX:
                        {
                            int c = 0;
                            formentry fe;
                            double inter = double.NegativeInfinity;
                            double inter2 = double.NegativeInfinity;
                            while ((fe = getarg(f, c)) != null)
                            {
                                inter = getvalue(fe, max_range);
                                if (inter > inter2)
                                    inter2 = inter;
                                c++;
                            }
                            result = inter2;
                            break;
                        }
                    case (int)SimpleOps.MIN:
                        {
                            int c = 0;
                            formentry fe;
                            double inter = double.PositiveInfinity;
                            double inter2 = double.PositiveInfinity;
                            while ((fe = getarg(f, c)) != null)
                            {
                                inter = getvalue(fe, min_range);
                                if (inter < inter2)
                                    inter2 = inter;
                                c++;
                            }
                            result = inter2;
                            break;
                        }
                    case (int)SimpleOps.IF:
                        {
                            formentry f0 = getarg(f, 0);
                            formentry f1 = getarg(f, 1);
                            formentry f2 = getarg(f, 2);
                            bool res = evaluate_logical(f0);
                            if (res)
                                result = getvalue(f1);
                            else
                                result = getvalue(f2); 
                            break; 
                        }
                    default:
                        throw new ParseException("Evaluation error, unrecognized operator"); 
 
                }
            return result;
        }

        private bool evaluate_logical(formentry f)
        {
            double res1;
            double res2;
            int comparator; 
            /* we get here with a precedence operator pointing a numstack with two args and an op with a comparator
             * normally there will be either three or four items on the numberstack: STARTPARSE, a single item that evaluates
             * either to zero or not, OR two items to be compared, followed by ENDPARSE*/
            if (f.tokentype == (int)TokenTypes.PRECEDENCE)
            {
                List<formentry> nstack = f.newnumstack;
                List<formentry> ostack = f.newopstack;

                /* syntax error checks */ 
                if (nstack == null || ostack == null || nstack.Count < 3)
                    return false;
                if (nstack.Count == 4 && ostack.Count != 1)
                    return false; 
                if (nstack[0].tokentype != (int)TokenTypes.STARTPARSE)
                    return false;

                res1 = getvalue(nstack[1]);
                if (nstack.Count == 3)
                    return (!(res1 == 0.0));
                res2 = getvalue(nstack[2]);
                comparator = (int)ostack[0].value;
                switch(comparator)
                {
                    case 0:                     /* equal */
                        return (res1 == res2);
                    case 1:                     /* greater than */
                        return (res1 > res2);
                    case 2:                     /* less than */
                        return (res1 < res2);
                    case 3:                     /* greater than or equal */
                        return (res1 >= res2);
                    case 4:                     /* less than or equal */
                        return (res1 <= res2);
                    case 5:                     /* not equal */
                        return (res1!=res2); 
                    default:
                        return false; 
                }
            }

            return true; 
        }

        private formentry getarg(formentry f, int arg)
        {
            if (f.tokentype == (int)TokenTypes.PRECEDENCE || f.tokentype==(int)TokenTypes.FUNC)
            {
                List<formentry> nstack = f.newnumstack;
                if (nstack == null)
                    return null; 
                if (arg >= nstack.Count-1)
                    return null;
                if (nstack[0].tokentype==(int)TokenTypes.STARTPARSE? nstack[arg+1].tokentype==(int)TokenTypes.ENDPARSE : nstack[arg].tokentype == (int)TokenTypes.ENDPARSE)
                    return null; 
                else
                    return (nstack[0].tokentype==(int)TokenTypes.STARTPARSE? nstack[arg+1] : nstack[arg]); /* STARTPARSE will be the first one */ 
            }
            else
                return null; 

        }

        private void RemoveArgSeps(formentry f)
        {
            bool barg = false;
            if (f.newnumstack != null)
            {
                while (barg = (f.newnumstack.Remove(new formentry((int)TokenTypes.ARGSEP) { tokentype = (int)TokenTypes.ARGSEP }) == true))
                    continue;
            }
        }


        private double max_range(formentry f)
        {
            int i;
            f.value = double.NegativeInfinity;
            if (f.range.cellcount == 0)
                return double.NaN;
            for (i = 0; i < f.range.cellcount; i++)
                if (f.range.values[i] > f.value)
                    f.value = f.range.values[i];
            return f.value;
        }

        private double min_range(formentry f)
        {
            int i;
            f.value = double.PositiveInfinity;
            if (f.range.cellcount == 0)
                return double.NaN;
            for (i = 0; i < f.range.cellcount; i++)
                if (f.range.values[i] < f.value)
                    f.value = f.range.values[i];
            return f.value;
        }
        private double sum_range(formentry f)
        {

            int i;
            for (i = 0; i < f.range.cellcount; i++)
                f.value += f.range.values[i];
            return f.value;
        }

        private double avg_range(formentry f)
        {
            int i;
            double sum = 0.0;
            for (i = 0; i < f.range.cellcount; i++)
                sum += f.range.values[i];
            f.value = sum / f.range.cellcount;
            return f.value;
        }

        private double getvalue(formentry f, eval_func e=null)
        {
            if (f == null)
                return double.NaN;

            /* if it's a range we'll handle it elsewhere */
            if (f.tokentype == (int)TokenTypes.RANGE)
                if (e != null)
                {
                    e(f);
                    return f.value;
                }

            if (f.tokentype == (int)TokenTypes.FUNC)
            {
                f.value = evaluate_function(f);
                return f.value; 
            }

            if (f.tokentype == (int)TokenTypes.PRECEDENCE)
            {
                f.value=(eval_worker(f.newnumstack, f.newopstack));
                return f.value; 
            }

            if (f.tokentype == (int)TokenTypes.NUMBER || f.tokentype==(int)TokenTypes.DATE)
                return (double)f.value;

            else
            {
                if (f.tokentype != (int)TokenTypes.CELLREF)
                    return double.NaN;
                else
                    return (getcellvalue(f.celladdr));
            }

        }
        /* use this for functions that take one string argument only
         * like =STOCK(IBM); must be one string at next level of precedence */ 
        private string GetSimpleString(formentry f)
        {
            if (f.tokentype == (int)TokenTypes.PRECEDENCE)
            {
                List<formentry> nstack = f.newnumstack;
                return nstack[1].text;
            }
            else
                return f.text;
        }

        /* stub for test */ 
        private float getcellvalue(Pt cell)
        {
            return (float)15.0;
        }

        private int LoadDataBlob(string name)
        {
            int lines = 0;
            int cols = 0;
            int filetype = 0; 

            System.IO.MemoryStream memStream = new MemoryStream(File.ReadAllBytes(name));

            if (name.Substring(name.Length - 3, 3).ToUpper() == "CSV")
                filetype = 0;
            if (name.Substring(name.Length - 2, 2).ToUpper() == "OP")
                filetype = 1; 

            StreamReader sr = new StreamReader(memStream); 

            string line;
            string hline=""; 
            while ((line = sr.ReadLine()) != null)
            {
                if (lines == 0)
                {
                    if (filetype == 0)
                        cols=CountColumnsCSV(line);
                    else if (filetype == 1)
                        hline = line; 
                }
                if (lines == 1)
                    if(filetype==1)
                        cols=CountColumnsOP(hline, line); 

                lines++;
            }


            _datarows = lines;
            _datacols = cols;
            _data = new string[_datarows, _datacols];
            memStream.Position = 0;
            try
            {
                int i = 0;
                using (StreamReader sr2 = new StreamReader(memStream))
                {

                    while ((line = sr2.ReadLine()) != null)
                    {
                        if (filetype == 0)
                            ParseCSVLine(line, i++);
                        if (filetype == 1)
                            ParseOPLine(line, i++);
                    }
                }

            }
            catch
            {
                _errorstring = "file parse error";
            }
            sr.Dispose();
            return lines * cols;
        } 

        /* data functions; I am aware this reads the file twice */
        private int CountLines(string fn)
        {
            int lines=-1;
            try
            {
                // Create an instance of StreamReader to read from a file. 
                // The using statement also closes the StreamReader. 
                using (StreamReader sr = new StreamReader("c:\\temp\\testw.csv"))
                {
                    string line;
                    // Read and display lines from the file until the end of  
                    // the file is reached. 
                    while ((line = sr.ReadLine()) != null)
                    {
                        lines++;
                    }
                }
            }
            catch (Exception exc)
            {
                return -1;
            }
                  
            return lines; 
        }
        private int CountColumnsCSV(string line)
        {
            int index = 0;
            int newind = 0;
            string s;
            char[] ch = { '"' };
            int i = 0; 
            int cols = 0; 
            while (true)
            {
                newind = line.IndexOf(',', index);
                if (newind != -1)
                {
                    /* second arg is length not end index ... sigh */
                    s = line.Substring(index, newind - index);
                    if (s == "")
                    {
                        _headings[i++] = s;
                        cols++;
                    }
                    else
                    {
                        if (s[0] == '"')
                            s = s.Substring(1);
                        if (s[s.Length - 1] == '"')
                            s = s.Trim(ch);

                        _headings[i++] = s;
                        cols++;
                    }
                    index = newind + 1;
                }
                else
                {
                    if (index < line.Length)
                    {
                        s = line.Substring(index, line.Length - index);
                        _headings[i] = s;
                        cols++;
                        return cols; 
                    }
                } /* else */ 
            } /* while */ 
        }

        private int CountColumnsOP(string hline, string line2)
        {
            int col = 0; 
            _opcolumns.Clear();

            /* use first data line to count columns since heading line can contain blanks */ 
            while (col < line2.Length)
            {
                if (line2[col] != ' ')
                    _opcolumns.Add(col); 
                while (col < line2.Length && line2[col] != ' ')
                    col++;
                while (col < line2.Length && line2[col] == ' ')
                    col++; 
            }
            return _opcolumns.Count;
            
        } 
        private int ParseCSVLine(string line, int lineno)
        {
            int index = 0;
            int i = 0;
            string s;
            int newind = 0;
            int cols = 0;
            char[] ch = { '"' };
            /* since data items can contain commas we look for quotes ... this is a limitation */
            while (true)
            {
                newind = line.IndexOf(',', index);
                if (newind != -1)
                {
                    /* second arg is length not end index ... sigh */
                    s = line.Substring(index, newind - index);
                    if (s == "")
                    {
                        _data[lineno,i++]=s;
                        cols++; 
                    }
                    else
                    {
                        if (s[0] == '"')
                            s = s.Substring(1);
                        if (s[s.Length - 1] == '"')
                            s = s.Trim(ch);
                        _data[lineno, i++] = s;
                        cols++;
                    }
                    index = newind + 1;
                }
                else
                {
                    if (index < line.Length)
                    {
                        s = line.Substring(index, line.Length - index);
                        _data[lineno, i++] = s;
                        cols++;
                        return cols;
                    }
                } /* else */
            } /* while */ 
        } 
        private int ParseOPLine(string line, int lineno)
        {
            string s = "";
            int i = 0;

            if (line.Length == 0)
                return 0;
            try
            {
                foreach (int col in _opcolumns)
                {
                    if (line[col] == ' ')
                        s = " ";
                    else
                    {
                        int j = col;
                        while (j<line.Length && line[j] != ' ')
                            s += line[j++];
                    }
                    _data[lineno, i++] = s;
                    s = "";
                } 
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return 1; 
        }

        /* for debugging */
        private void DumpFormentry(formentry f)
        {
            Console.WriteLine("Token type: "+ f.tokentype.ToString());
            Console.WriteLine("Value: " + f.value.ToString());
            Console.WriteLine("Pointer: " + f.ptr.ToString());
            if (f.text != null)
                Console.WriteLine("Text: " + f.text);
            else
                Console.WriteLine("Text null");
            if (f.celladdr != null)
                Console.WriteLine("Celladdr: X: " + f.celladdr.X.ToString() + "Y: " + f.celladdr.Y.ToString());
            else
                Console.WriteLine("Celladdr null"); 
        }
        #endregion

    }
}