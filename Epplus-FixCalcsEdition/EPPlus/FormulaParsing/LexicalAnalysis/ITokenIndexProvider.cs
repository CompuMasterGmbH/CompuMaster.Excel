using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EpplusFreeOfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public interface ITokenIndexProvider
    {
        int Index { get;  }

        void MoveIndexPointerForward();
    }
}
