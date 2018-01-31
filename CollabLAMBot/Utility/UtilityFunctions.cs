using System;

namespace CollabLAMBot.Utility
{
    public class UtilityFunctions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string setGreetingContext()
        {
            string greetingMessage = string.Empty;

            DateTime todaynow = DateTime.Now;
            if (todaynow.Hour > 6 && todaynow.Hour < 12)
                greetingMessage = "Good morning";
            else if (todaynow.Hour >= 12 && todaynow.Hour <= 16)
                greetingMessage = "Good afternoon";
            else if (todaynow.Hour > 16 && todaynow.Hour <= 20)
                greetingMessage = "Good evening";
            else
                greetingMessage = "Hi";
            return greetingMessage;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string getWindowsUser()
        {
            return System.Security.Principal.WindowsIdentity.GetCurrent().Name;

           
        }

    }
}