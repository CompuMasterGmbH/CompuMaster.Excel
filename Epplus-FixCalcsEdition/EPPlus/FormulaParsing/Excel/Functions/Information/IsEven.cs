﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CompuMaster.Epplus4.FormulaParsing.ExpressionGraph;
using CompuMaster.Epplus4.Utils;

namespace CompuMaster.Epplus4.FormulaParsing.Excel.Functions.Information
{
    public class IsEven : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg1 = GetFirstValue(arguments);//arguments.ElementAt(0);
            if (!ConvertUtil.IsNumeric(arg1))
            {
                ThrowExcelErrorValueException(eErrorType.Value);
            }
            var number = (int)System.Math.Floor(ConvertUtil.GetValueDouble(arg1));
            return CreateResult(number % 2 == 0, DataType.Boolean);
        }
    }
}
