﻿/* Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EpplusFreeOfficeOpenXml.FormulaParsing.Utilities;

namespace EpplusFreeOfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class BoolArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            if (obj is ExcelDataProvider.IRangeInfo)
            {
                var r = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault();
                obj = (r == null ? null : r.Value);
            }
            if (obj == null) return false;
            if (obj is bool) return (bool)obj;
            if (obj.IsNumeric()) return Convert.ToBoolean(obj);
            bool result;
            if (bool.TryParse(obj.ToString(), out result))
            {
                return result;
            }
            return result;
        }
    }
}
