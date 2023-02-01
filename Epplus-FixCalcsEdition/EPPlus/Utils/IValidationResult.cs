using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CompuMaster.Epplus4.Utils
{
    public interface IValidationResult
    {
        void IsTrue();
        void IsFalse();
    }
}
