using System.Diagnostics;

namespace TaskScheduler
{
    class _Format
    {
        public string Date_(string t) 
        {
            Debug.WriteLine("Значение даты =  {0} ", t);
            string d = "";
            try
            {
                if (t.Length < 2)
                {
                    d = "0" + t;
                    Debug.WriteLine("Значение даты после обработки =  {0} ", d);
                    return d;
                }
            }
            catch { }
            return t;
        }
        public string DateNow_()
        {
            string t = System.DateTime.Now.Year.ToString() + "_" + Date_(System.DateTime.Now.Month.ToString()) + "_" + Date_(System.DateTime.Now.Day.ToString());
            return t;
        }
        public string DateTimeNow_()
        {
            string t = System.DateTime.Now.Year.ToString() + Date_(System.DateTime.Now.Month.ToString()) + Date_(System.DateTime.Now.Day.ToString()) + "_" + Date_(System.DateTime.Now.Hour.ToString()) + ":" + Date_(System.DateTime.Now.Minute.ToString()) + ":" + Date_(System.DateTime.Now.Second.ToString());
            return t;
        }
    }
}
