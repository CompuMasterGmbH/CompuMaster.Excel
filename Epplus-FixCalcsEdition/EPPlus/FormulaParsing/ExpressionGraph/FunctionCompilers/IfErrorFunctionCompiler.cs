using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EpplusFreeOfficeOpenXml.FormulaParsing.Excel.Functions;
using EpplusFreeOfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using EpplusFreeOfficeOpenXml.FormulaParsing.Exceptions;
using EpplusFreeOfficeOpenXml.FormulaParsing.Utilities;

namespace EpplusFreeOfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class IfErrorFunctionCompiler : FunctionCompiler
    {
        public IfErrorFunctionCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {
            Require.That(function).Named("function").IsNotNull();
          
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            if (children.Count() != 2) throw new ExcelErrorValueException(eErrorType.Value);
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var firstChild = children.First();
            var lastChild = children.ElementAt(1);
            try
            {
                var result = firstChild.Compile();
                if (result.DataType == DataType.ExcelError)
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
