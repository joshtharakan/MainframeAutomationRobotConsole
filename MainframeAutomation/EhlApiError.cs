using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MainframeAutomation
{
    class EhlApiError : Exception
    {
        public EhlApiError()
        {

        }

        public EhlApiError(string message)
            : base(message)
        {
            System.Console.WriteLine(message);
        }

        public EhlApiError(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
