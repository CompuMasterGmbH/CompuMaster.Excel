﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EpplusFreeOfficeOpenXml.FormulaParsing.Exceptions;
using EpplusFreeOfficeOpenXml.FormulaParsing.ExpressionGraph;
using EpplusFreeOfficeOpenXml.Utils;

namespace EpplusFreeOfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class N : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = GetFirstValue(arguments);
            
            if (arg is bool)
            {
                var val = (bool) arg ? 1d : 0d;
                return CreateResult(val, DataType.Decimal);
            }
            else if (IsNumeric(arg))
            {
                var val = ConvertUtil.GetValueDouble(arg);
                return CreateResult(val, DataType.Decimal);
            }
            else if (arg is string)
            {
                return CreateResult(0d, DataType.Decimal);
            }
            else if (arg is ExcelErrorValue)
            {
                return CreateResult(arg, DataType.ExcelError);
            }
            throw new ExcelErrorValueException(eErrorType.Value);
        }
    }
}
