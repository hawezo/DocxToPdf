using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxToPdf.Core
{
    public static class Logger
    {

        public static bool Silent = false;

        public static void Normal(string message)
        {
            if (!Silent)
                Logger.Entry("Normal", message);
        }

        public static void Normal(string message, params string[] args)
        {
            if (!Silent)
                Logger.Entry("Normal", String.Format(message, args));
        }

        public static void Error(string message)
        {
            Logger.Entry("Error", message);
        }

        public static void Error(string message, params string[] args)
        {
            Logger.Entry("Error", String.Format(message, args));
        }

        private static void Entry(string prefix, string message)
        {
            Console.WriteLine(String.Format("[{0}] ({1}) {2}", DateTime.Now.ToString("HH:mm:ss"), prefix, message));
        }
    }
}
