using System;

namespace HomeConnect
{
    public class LogHelper
    {
        /*
         *  Created by Lucas-VD on 1st of April 2017
         *  This project is currently under a CC BY-NC-SA 4.0 license (Because I don't really know how licensing works)
         */

        public static void Log(String logLevel, Object obj)
        {
            Console.WriteLine(DateTime.Now.ToString("[HH:mm:ss.ff]") + " " + logLevel + " " + obj.ToString());
        }
        public static void Debug(Object obj)
        {
            Log("[DEBUG]", obj);
        }

        public static void Error(Object obj)
        {
            Log("[ERROR]", obj);
        }

        public static void Info(Object obj)
        {
            Log("[INFO]", obj);
        }

        public static void Speak(Object obj)
        {
            Log("[SPEAK]", obj);
        }

        public static void Warn(Object obj)
        {
            Log("[WARN]", obj);
        }
    }
}
