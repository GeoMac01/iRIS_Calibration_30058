using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.IO.Ports;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Thorlabs.PM100D_32.Interop;
using MccDaq;

namespace iRIS_LASER_TEST_01
{

    public class COMsSendRecv
    {
        public void ChopString(string stringToChop)
        {
            int strLgh = stringToChop.Length;
            //rtnHeader = stringToChop.Substring(1, 2);
            //rtnCmd = stringToChop.Substring(3, 2);
            //rtnValue = stringToChop.Substring(5, (strLgh - 7));
            strLgh = 0;
        }
    }

}
