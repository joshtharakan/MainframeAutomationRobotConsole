using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MainframeAutomation
{
	class EhlApiExceptions : Exception
	{
        public EhlApiExceptions()
        {
        
        }

        public EhlApiExceptions(string message)
            : base(message)
        {
            System.Console.WriteLine(message);
        }

        public EhlApiExceptions(string message, Exception inner)
            : base(message, inner)
        {
        }
	}
}

