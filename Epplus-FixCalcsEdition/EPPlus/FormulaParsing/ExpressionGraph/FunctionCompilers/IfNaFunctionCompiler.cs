﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CompuMaster.Epplus4.FormulaParsing.Excel.Functions;
using CompuMaster.Epplus4.FormulaParsing.Exceptions;

namespace CompuMaster.Epplus4.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class IfNaFunctionCompiler : FunctionCompiler
    {
        public IfNaFunctionCompiler(ExcelFunction function, ParsingContext context)
            :base(function, context)
        {
            
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            if (children.Count() != 2) return new CompileResult(eErrorType.Value);
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var firstChild = children.First();
            var lastChild = children.ElementAt(1);
            try
            {
                var result = firstChild.Compile();
                if (result.DataType == DataType.ExcelError && (Equals(result.Result,
                    ExcelErrorValue.Create(eErrorType.NA))))
                {
                    args.Add(new FunctionArgument(lastChild.Compile().Result));
                }
                else
                {
                    args.Add(new FunctionArgument(result.Result));
                }

            }
            catch (ExcelErrorValueException)
            {
                args.Add(new FunctionArgument(lastChild.Compile().Result));
            }
            return Function.Execute(args, Context);
        }
    }
}
