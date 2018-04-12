using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Log
{
    public class HttpLoggerHandler : BaseLoggerHandler
    {
        protected override void Log(LoggerMessage message)
        {
            throw new NotImplementedException();
        }

        protected override void OnShutdown()
        {
        }
    }
}
