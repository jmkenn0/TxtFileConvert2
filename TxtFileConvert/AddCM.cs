using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace TxtFileConvert
{
    public  class AddCM
    {
        public string AddCMUtility(string existingCM, string addCM)
        {
            double existingCMdub = 0;
            var cultureInfo = CultureInfo.GetCultureInfo("de-DE");
            try
            {
                //MessageBox.Show(existingCM.ToString());
                existingCMdub = double.Parse(existingCM, cultureInfo);
            }
            catch (Exception e)
            {
                
                throw;
            }
            
            double addCMdub = existingCMdub+ double.Parse(addCM, cultureInfo);
           // MessageBox.Show(String.Format(cultureInfo, "{0}", addCMdub));
            return String.Format(cultureInfo, "{0}", addCMdub);
        }
        
              
    }
}
