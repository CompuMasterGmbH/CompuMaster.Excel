﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EpplusFreeOfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EpplusFreeOfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    public class IfError : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.First();
            return GetResultByObject(firstArg.Value);
        }
    }
}
