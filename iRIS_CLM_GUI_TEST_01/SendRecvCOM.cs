using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace iRIS_CLM_GUI_TEST_02
{

    public class SendRecvCOM
    {
        //====================================================================
        public static string[] ChopString(string stringToChop)
        {
            string[] rtnstr = new string[3];
            int strLgh = stringToChop.Length;
            rtnstr[0] = stringToChop.Substring(1, 2);
            rtnstr[1] = stringToChop.Substring(3, 2);
            rtnstr[2] = stringToChop.Substring(5, (strLgh - 7));
            strLgh = 0;
            return rtnstr;
        }
        //====================================================================

        //====================================================================

        //====================================================================


    }
}
