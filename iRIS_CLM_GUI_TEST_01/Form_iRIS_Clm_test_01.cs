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

//iRIS Production 30058_01
//14/12/2017

namespace iRIS_CLM_GUI_TEST_01
{
    public partial class Form_iRIS_Clm_test_01 : Form
    {
        #region Commands Definition
        const string rtnNull            = "00";
        const string CmdLaserEnable     = "02";
        const string CmdSetLsPw         = "03";
        const string CmdRdSerialNo      = "04";
        const string CmdRdFirmware      = "06";
        const string CmdRdBplateTemp    = "07";
        const string CmdRdWavelen       = "08";
        const string CmdSetUnitNo       = "12";
        const string CmdLaserStatus     = "14";
        const string CmdRdTecTemprt     = "15";
        const string CmdHELP            = "16";
        const string CmdsetTTL          = "17";
        const string CmdEnablLogicvIn   = "18";
        const string CmdAnalgInpt       = "19";
        const string CmdRdLsrStatus     = "20";
        const string CmdSetPwMonOut     = "21";
        const string CmdRdCalDate       = "22";
        const string CmdSetPwCtrlOut    = "23";
        const string CmdSetOffstVolt    = "24";
        const string CmdSetVgaGain      = "25";
        const string CmdOperatingHr     = "26";
        const string CmdRdSummary       = "27";
        const string CmdSetInOutPwCtrl  = "28";
        const string CmdSet0mA          = "29";
        const string CmdSetStramind     = "30";
        const string CmdRdCmdStautus2   = "34";
        const string CmdManufDate       = "40";
        const string CmdRdPwSetPcon     = "41";
        const string CmdRdInitCurrent   = "42";
        const string CmdRdModelName     = "43";
        const string CmdRdLaserPow      = "44";
        const string CmdRdPnNb          = "45";
        const string CmdRdCustomerPm    = "46";
        const string CmdRatedPower      = "47";
        const string CmdCurrentRead     = "56";
        const string CmdSetPwtoVout     = "59";
        const string CmdSetCalAPw       = "60";
        const string CmdSetCalBPw       = "61";
        const string CmdSetCalAPwtoVint = "62";
        const string CmdSetCalBPwtoVint = "63";
        const string CmdRstTime         = "66";
        //const string CmdRstPtr          = "67";
        //const string CmdRstTon          = "68";
        //const string CmdRstCntr1000     = "69";
        //const string CmdSetFirmware     = "70";
        const string CmdSetSerNumber    = "71";
        const string CmdSetWavelenght   = "72";
        const string CmdSetLsMominalPw  = "73";
        const string CmdSetCustomerPm   = "74";
        const string CmdSetMaxIop       = "76";
        const string CmdSetCalDate      = "77";
        const string CmdSeManuDate      = "78";
        const string CmdSetPartNumber   = "79";
        const string CmdSetModel        = "80";
        const string CmdSetCalAVtoPw    = "81";
        const string CmdSetCalBVtoPw    = "82";
        const string CmdTestMode        = "83";
        const string CmdSetPSU          = "84";
        //const string CmdRdPSUvolt       = "85";//check cmd 86 / 85
        const string CmdReadPSU         = "86";
        const string CmdSetBaseTempCal  = "87";
        const string CmdSetTECTemp      = "90";
        const string CmdSetTECkp        = "91";
        const string CmdSetTECki        = "92";
        const string CmdSetTECsmpTime   = "93";
        const string CmdRdTECsetTemp    = "94";
        const string CmdRdTECsetkp      = "95";
        const string CmdRdTECsetki      = "96";
        const string CmdRdTECsmpTime    = "97";
        const string CmdSetTECena_dis   = "98";
        const string CmdRdUnitNo        = "99";
        const string Footer             = "\r\n";
        const string Header             = "#";
        const string StrEnable          = "0001";
        const string StrDisable         = "0000";
        #endregion

        //=================================================
        #region Test Sequence Definition
        //=================================================

        string[,] bulkSetLaserIO = new string[8, 2] {   //the rest of the string is build with case...
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetTECena_dis,     StrEnable  },
            { CmdSetInOutPwCtrl,    StrDisable },       //external PCON
            { CmdAnalgInpt,         StrDisable },       //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },       //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable },        //Inv. TTL line in nothing connected
            { CmdRdWavelen,         StrDisable }, };    
 
        string[,] bulkSetTEC = new string[6, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetTECTemp,        StrDisable },
            { CmdSetTECkp,          StrDisable },
            { CmdSetTECki,          StrDisable },
            { CmdSetTECsmpTime,     StrDisable },
            { CmdSetTECena_dis,     StrEnable} };

        string[,] bulkSetVarialble = new string[7, 2] {
            {CmdTestMode,       StrEnable },
            //{CmdSetWavelenght,    StrDisable},
            //{CmdSetLsMominalPw,   StrDisable},
            //{CmdSetMaxIop,        StrDisable},
            //{CmdSetSerNumber,     StrDisable},
            //{CmdSetModel,         StrDisable},
            //{CmdSeManuDate,       StrDisable},
            //{CmdSetCalDate,       StrDisable},
            //{CmdSetPartNumber,    StrDisable},
            {CmdSetCalAPw,          StrDisable},
            {CmdSetCalBPw,          StrDisable},
            {CmdSetCalAPwtoVint,    StrDisable},
            {CmdSetCalBPwtoVint,    StrDisable},
            {CmdSetCalAVtoPw,       StrDisable},
            {CmdSetCalBVtoPw,       StrDisable}, };
            //{CmdRdFirmware,       StrDisable} };

        string[,] bulkSetdefaultCtrl = new string[6, 2] {
            {CmdTestMode,       StrEnable  },
            {CmdRatedPower,     StrDisable },
            {CmdSetPwMonOut,    StrDisable },
            {CmdSetVgaGain,     StrDisable },
            {CmdSetOffstVolt,   StrDisable },      //Offset 2.500V
            {CmdSetPwCtrlOut,   StrDisable } };    //Internal PCON 2.500V
    
        //=================================================

        string[,] bulkSetVga = new string[6, 2] {
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetInOutPwCtrl,    StrDisable },      //external PCON
            { CmdAnalgInpt,         StrDisable },      //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },      //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable } };    //Inv. TTL line in

        //=================================================

        string[,] analogRead = new string[7, 2] {//read all analog inputs
            { CmdTestMode,          StrEnable },
            { CmdRdCmdStautus2,     StrDisable },
            { CmdRdPwSetPcon,       StrDisable },
            { CmdRdLaserPow,        StrDisable },
            { CmdCurrentRead,       StrDisable },
            { CmdRdTecTemprt,       StrDisable },
            { CmdRdBplateTemp,      StrDisable } };

        #endregion
        //=================================================
        string[]  testStringArr   =     new string[2];//used to load commands in bulk send
        
        string indata_USB =     string.Empty;
        string outdata_RS232 =  string.Empty;
        string rString =        string.Empty;
        string cmdTrack =       string.Empty;

        string  rtnHeader =     string.Empty;
        string  rtnCmd =        string.Empty;
        string  rtnValue =      string.Empty;
        string dataBaseName =   string.Empty;

        byte[] byteArrayToTest1 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest2 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest3 = new byte[8];//reads back "bits"

        bool USB_Port_Open =    false;
        bool RS232_Port_Open =  false;
        bool intExtCmd =        false;
        bool engFlag =          false;
        bool testMode =         false;
            
        int arrayLgth =         0;
        
        //======================================================================
        SendRecvCOM sendRcv = new SendRecvCOM();
        //======================================================================
        StringBuilder LogString_01 = new StringBuilder();
        //======================================================================
        Thorlabs.PM100D_32.Interop.PM100D pm = null;
        bool pm100ok = false;
        //======================================================================
        SerialPort USB_CDC =        new SerialPort();
        SerialPort RS232 =          new SerialPort();
        //======================================================================
        SqlConnection con;
        SqlCommand cmd;
        SqlDataAdapter adapt;
        int iD = 0; //loaded with the new row index
        //======================================================================
    #region// Setting ADCDAC IO USB Interface
        public MccDaq.DaqDeviceDescriptor[] inventory;
        public MccDaq.MccBoard DaqBoard;
        public MccDaq.ErrorInfo ULStat;
        public MccDaq.Range Range;
        public MccDaq.Range AORange;
        public MccDaq.Range RangeSelected;
        
        Int32 numchannels = 0;
        int   nudAInChannel = 0;
        //======================================================================
        //System.String LastPass = "Upper";
        public const String AllowedCharactersInt = "0123456789";
        public const String AllowedCharactersFloat = "0123456789.";
        //======================================================================
        #endregion
        //======================================================================
         public Form_iRIS_Clm_test_01()
        {
            InitializeComponent();
            Getportnames();

            USB_CDC.DataReceived += new SerialDataReceivedEventHandler(CDCDataReceivedHandler);
            RS232.DataReceived   += new SerialDataReceivedEventHandler(RS232DataReceivedHandler);
            //OpenSqlConnection();
            //DisplayData();
        }
        //================================================================================

        //================================================================================
        #region SQL stuff
        private void OpenSqlConnection()
        {
            dataBaseName = " " + Tb_DatabaseString.Text + " ";
            string connStr = GetConnectionString();

            try {
                con = new SqlConnection(connStr);
                con.Open();
                MessageBox.Show("  ServerVersion: " + con.ServerVersion + "  State: " + con.State);
                con.Close();
            }
            catch (Exception e) {MessageBox.Show("Dtb Error " + e.ToString()); }
         }
        //================================================================================
        private string GetConnectionString() {
            string DbUsername = Tb_User1.Text;
            string DbPassword = Tb_Pw1.Text;
            string DbServerName = Tb_ServerName.Text;
            string DbName = Tb_InitialCatalog.Text;

            string connectionString =   "Server = " + DbServerName      + ";" + //Data Source
                                        "Database = " + DbName          + ";" + //Initial catalog  
                                        "Trusted_Connection = false"    + ";" +//false: user name and PW
                                        "User Id = " + DbUsername       + ";" +
                                        "Password = " + DbPassword; 
            return connectionString;
        }
        //================================================================================
        private void DisplayData() {

            string adapterString = "SELECT * FROM" + dataBaseName + "WHERE LaserDriverBdTestId = (SELECT max(LaserDriverBdTestId) FROM" + dataBaseName +")";
            con.Open();
	        DataTable dt = new DataTable();
            adapt = new SqlDataAdapter(adapterString, con);

            try {
                adapt.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception e) { MessageBox.Show("DB adaptor error" + e.ToString()); }

            con.Close();
            }
        //================================================================================
        #endregion SQL stuff
        //================================================================================
        private void CDCDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            string indata_USB = string.Empty;
            indata_USB = USB_CDC.ReadExisting();  
            this.BeginInvoke(new Action(() => Process_String(indata_USB)));
            //this.BeginInvoke(new EventHandler(delegate { Process_String(indata_USB); }));
            //Application.DoEvents();
        }
        //================================================================================
        private void RS232DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            string indata_RS232 = string.Empty;
            indata_RS232 = RS232.ReadExisting();  //read data to string 
            this.BeginInvoke(new Action(() => Process_String(indata_RS232)));
        }
        //======================================================================
        private async Task<bool> SendToSerial(string strCmd, string strData, int sendDelay)
        {
            string mdlNumb = Tb_SetAdd.Text;
            string stuffToSend = string.Empty;
            int dly = 300;

            if (sendDelay > 300 ) {dly = sendDelay; }
            else dly = Convert.ToInt16(Tb_RsDelay.Text);
            
            cmdTrack = strCmd;//used for the read back test

            stuffToSend = Header + mdlNumb + strCmd + strData + Footer;
            Rt_ReceiveDataUSB.AppendText(">>  " + stuffToSend); //displays anything....

            try
            {
                if (USB_Port_Open == true)
                {
                    USB_CDC.DiscardInBuffer();
                    USB_CDC.DiscardOutBuffer();
                    USB_CDC.Write(stuffToSend);
                }
                else if(RS232_Port_Open == true)
                {
                    RS232.DiscardOutBuffer();
                    RS232.DiscardInBuffer();
                    RS232.Write(stuffToSend);
                }

              await Task.Delay(dly);
            }
            catch (Exception) { MessageBox.Show("COM Write Error"); }
            return true;                        
        }
        //======================================================================
        #region Process_String long case where the received string is analysed
        private void Process_String(string strRcv)
        {
            string[] returnChop = new string[3];  // 0/Header, 1/cmd, 2/value
            string rbCmd = string.Empty;
            int rtnValueInt = 0;

            Rt_ReceiveDataUSB.AppendText("<<  " + strRcv); //displays anything....
            returnChop = SendRecvCOM.ChopString(strRcv);
            rtnHeader = returnChop[0];
            rtnCmd = returnChop[1];

            rtnValue = returnChop[2];
                                    
            if (rtnCmd=="00") MessageBox.Show("rtn null");
            else if (rtnCmd == cmdTrack)
            {   cmdTrack = string.Empty;

                switch (rtnCmd)
                {

                    case CmdRdUnitNo://99
                        if (RS232_Port_Open == true)
                        {
                            Tb_RSConnect.BackColor = Color.LawnGreen;
                            Tb_USBConnect.BackColor = Color.Red;
                            /*
                            con.Open();
                            cmd = new SqlCommand("update " + dataBaseName + " set RS232Connet = @RS232OK where LaserId = @LaserId", con);
                            cmd.Parameters.AddWithValue("@LaserId", iD);
                            cmd.Parameters.AddWithValue("@RS232OK", 1);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            DisplayData();
                            */
                        }
                        else if (USB_Port_Open == true)
                        {
                            Tb_RSConnect.BackColor = Color.Red;
                            Tb_USBConnect.BackColor = Color.LawnGreen;
                            /*
                            con.Open();
                            cmd = new SqlCommand("update " + dataBaseName + " set USBConnect = @USBOK where LaserId = @LaserId", con);
                            cmd.Parameters.AddWithValue("@LaserId", iD);
                            cmd.Parameters.AddWithValue("@USBOK", 1);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            DisplayData();
                            */
                        }
                        break;

                    case CmdSetUnitNo://12
                        lbl_RdAdd.Text = rtnHeader;

                        if (lbl_RdAdd.Text == Tb_SetAdd.Text)
                        {
                            Tb_EepromGood.BackColor = Color.LawnGreen;
                            lbl_RdAdd.ForeColor = Color.Green;
                            lbl_RdAdd.Text = rtnHeader;
                            /*
                            con.Open();
                            cmd = new SqlCommand("update " + dataBaseName + " set PgmAddr = @PgmAddr where LaserId = @LaserId", con);
                            cmd.Parameters.AddWithValue("@LaserId", iD);
                            cmd.Parameters.AddWithValue("@PgmAddr", Tb_SetAdd.Text);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            DisplayData();
                            */
                        }
                        else {Tb_EepromGood.BackColor = Color.Red; } 
                        break;

                    case CmdRdWavelen:
                        Lbl_WaveLg.Text = rtnValue;
                        break;

                    case CmdLaserEnable://uses serial send to set
                        if (rtnValue == StrDisable) { Bt_LaserEn.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_LaserEn.BackColor = Color.LawnGreen; }
                        break;

                    case CmdSetLsPw:
                        break;

                    case CmdRdSerialNo:
                        lbl_SerNbReadBack.ForeColor = Color.Green;
                        lbl_SerNbReadBack.Text = rtnValue.PadLeft(8,'0');
                        break;

                    case CmdRdFirmware:
                        lbl_SWLevel.ForeColor = Color.Green;
                        lbl_SWLevel.Text = rtnValue.PadLeft(8,'0');
                        /*
                        con.Open();
                        cmd = new SqlCommand("update " + dataBaseName + " set SwLevel = @SwLevel where LaserId = @LaserId", con);
                        cmd.Parameters.AddWithValue("@LaserId", iD);
                        cmd.Parameters.AddWithValue("@SwLevel", lbl_SWLevel.Text);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        DisplayData();
                        */
                        break;

                    case CmdRdBplateTemp:
                        Lbl_TempBplt.Text = rtnValue.PadLeft(4,'0');
                        break;

                    case CmdLaserStatus:

                        byteArrayToTest1 = ConvertToByteArr(rtnValue);

                        rtnValueInt = byteArrayToTest1[7];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit0.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit0.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest1[6];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit1.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit1.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest1[5];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit2.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit2.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest1[4];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit3.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit3.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest1[3];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit4.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit4.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest1[2];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd14Bit5.BackColor = Color.Red;
                        }
                        else tb_Cmd14Bit5.BackColor = Color.LawnGreen;

                        break;

                    case CmdRdTecTemprt:
                        //lbl_TecTemp.Text = rtnValue.PadLeft(5,'0');
                        break;

                    case CmdHELP:
                        break;

                    case CmdsetTTL:
                        if (rtnValue == StrDisable) { Bt_InvDigtMod.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_InvDigtMod.BackColor = Color.LawnGreen; }
                        break;

                    case CmdEnablLogicvIn:
                        if (rtnValue == StrDisable) { Bt_InvEnable.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_InvEnable.BackColor = Color.LawnGreen; }
                        break;

                    case CmdAnalgInpt:
                        if (rtnValue == StrDisable) { Bt_InvAnlg.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_InvAnlg.BackColor = Color.LawnGreen; }
                        break;

                    case CmdRdLsrStatus:

                        byteArrayToTest2 = ConvertToByteArr(rtnValue);

                        rtnValueInt = byteArrayToTest2[7];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit0.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit0.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[6];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit1.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit1.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[5];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit2.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit2.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[4];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit3.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit3.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[3];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit4.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit4.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[2];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit5.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit5.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[1];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit6.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit6.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest2[0];
                        if (rtnValueInt == 0)
                        {
                            tb_Cmd20Bit7.BackColor = Color.Red;
                        }
                        else tb_Cmd20Bit7.BackColor = Color.LawnGreen;

                        break;

                    case CmdSetPwMonOut:
                        break;

                    case CmdRdCalDate:
                        break;

                    case CmdSetPwCtrlOut:
                        break;

                    case CmdSetOffstVolt:
                        break;

                    case CmdSetVgaGain:
                        break;

                    case CmdOperatingHr:
                        break;

                    case CmdRdSummary:
                        break;

                    case CmdSetInOutPwCtrl:
                        if (rtnValue == StrDisable) {
                            intExtCmd = false;//used to select text box
                            Bt_IntExtPw.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) {
                            intExtCmd = true;
                            Bt_IntExtPw.BackColor = Color.LawnGreen; }
                        break;

                    case CmdSet0mA:
                        //firmware "error" : returns 00 as cmd
                        break;

                    case CmdSetStramind:
                        break;

                   case CmdRdCmdStautus2://cmd 34 note the array starts from 7 to 0....

                        byteArrayToTest3 = ConvertToByteArr(rtnValue);

                        rtnValueInt = byteArrayToTest3[7];
                        if (rtnValueInt == 0) { tb_Cmd34Bit0.BackColor = Color.Red; }
                        else tb_Cmd34Bit0.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest3[6];
                        if (rtnValueInt == 0) { tb_Cmd34Bit1.BackColor = Color.Red; }
                        else tb_Cmd34Bit1.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest3[5];
                        if (rtnValueInt == 0) { tb_Cmd34Bit2.BackColor = Color.Red; }
                        else tb_Cmd34Bit2.BackColor = Color.LawnGreen;

                        rtnValueInt = byteArrayToTest3[4];// Bit indicating status of the CPU control line LASER_EN_OUT_CPU
                        if (rtnValueInt == 0) tb_Cmd34Bit3.BackColor = Color.Red;
                        else tb_Cmd34Bit3.BackColor = Color.LawnGreen;

                        break;

                    case CmdManufDate:
                        break;

                    case CmdRdPwSetPcon:
                        if (intExtCmd == true) lbl_ADCpconRd.Text = rtnValue.PadLeft(5,'0');
                        else if (intExtCmd == false) lbl_ADCpconRd.Text = rtnValue.PadLeft(5,'0');
                        break;

                    case CmdRdInitCurrent:
                        break;

                    case CmdRdModelName:
                        break;

                    case CmdRdLaserPow:
                        lbl_LaserPD.Text = rtnValue.PadLeft(5,'0');
                        break;

                    case CmdRdPnNb:
                        break;

                    case CmdRdCustomerPm:
                        break;

                    case CmdRatedPower:
                        break;

                    case CmdCurrentRead:
                        if (testMode == true)
                        {
                            Lbl_uClsCurrent.Text = rtnValue.PadLeft(5, '0');
                            Lbl_MaOrBits.Text = "uC Ls. bits";
                        }
                        else if (testMode==false)
                        {
                            Lbl_uClsCurrent.Text = rtnValue.PadLeft(4, '0');
                            Lbl_MaOrBits.Text = "uC Ls. ImA";
                        }

                        break;

                    //case CmdSetPwtoVout://no return
                    //    break;

                    case CmdSetCalAPw:
                        Tb_CalA_Pw.Text = rtnValue;
                        break;

                    case CmdSetCalBPw:
                        Tb_CalB_Pw.Text = rtnValue;
                        break;

                    case CmdSetCalAVtoPw:
                        Tb_CalAcmdToPw.Text = rtnValue;
                        break;

                    case CmdSetCalBVtoPw:
                        Tb_CalBcmdToPw.Text = rtnValue;
                        break;

                    case CmdRstTime:
                        break;

                    case CmdSetSerNumber:
                        Tb_SerNb.Text = rtnValue.PadLeft(8, '0');                      
                        break;

                    case CmdSetWavelenght:
                        Lbl_WaveLg.Text = rtnValue;
                        break;

                    case CmdSetLsMominalPw:
                        break;

                    case CmdSetCustomerPm:
                        break;

                    case CmdSetMaxIop:
                        break;

                    case CmdSetCalDate:
                        break;

                    case CmdSeManuDate:
                        break;

                    case CmdSetPartNumber:
                        break;

                    case CmdSetModel:
                        break;

                    case CmdTestMode:
                        if (rtnValue == "0000") {
                            testMode = false;
                            Bt_EnableTest.BackColor = Color.SandyBrown; }
                        else if (rtnValue == "0001") {
                            testMode = true;
                            Bt_EnableTest.BackColor = Color.LawnGreen; }
                        break;

                    case CmdSetPSU:
                        break;

                    case CmdSetBaseTempCal:
                        break;

                    case CmdSetCalAPwtoVint:
                        Tb_CalA_PwToADC.Text = rtnValue;
                        break;

                    case CmdSetCalBPwtoVint:
                        Tb_CalB_PwToADC.Text = rtnValue;
                        break;

                    case CmdSetTECTemp:
                        break;

                    case CmdSetTECkp:
                        break;

                    case CmdSetTECki:
                        break;

                    case CmdSetTECsmpTime:
                        break;

                    case CmdRdTECsetTemp:
                        break;

                    case CmdRdTECsetkp:
                        break;

                    case CmdRdTECsetki:
                        break;

                    case CmdRdTECsmpTime:
                        break;

                    case CmdSetTECena_dis:
                        break;

                    default:
                        MessageBox.Show("Case Default Reached");
                        break;

                }//end of "SWITCH"
            }//end of "IF"

            else
            {
                Tb_RSConnect.BackColor = Color.Red;
                Tb_USBConnect.BackColor = Color.Red;
                MessageBox.Show("Read Back Missmatch");
            }
        }//end of "ProcessString"
        #endregion
        //======================================================================
        #region BuildSendString long case where the commands and argument are put together
        private async Task<bool> BuildSendString(string[] strCmd)
        {
            string dataToAppd = string.Empty;
            string cmdToTest = string.Empty;
            int sndDl = 300;
            int comThresh = 9;

            cmdToTest = strCmd[0];
            dataToAppd = strCmd[1]; //if anything different will be changed in the case below

            switch (cmdToTest)
            {
                case CmdRdUnitNo:
                    break;

                case CmdSetUnitNo:
                    dataToAppd = "00" + Tb_SetAdd.Text;//0002 i.e.
                    sndDl = 300;
                    break;

                case CmdSetPartNumber:
                    dataToAppd = Tb_LaserPN.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdWavelen:
                    break;

                case CmdLaserEnable://uses serial send to set
                    break;

                case CmdSetLsPw:
                    break;

                case CmdRdSerialNo:
                    break;

                case CmdRdFirmware:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdBplateTemp:
                    break;

                case CmdLaserStatus:
                    break;

                case CmdRdTecTemprt:
                    break;

                case CmdHELP:
                    break;

                case CmdsetTTL:
                    if (dataToAppd == StrDisable) { Bt_InvDigtMod.BackColor = Color.SandyBrown; }
                    else if (dataToAppd == StrEnable) { Bt_InvDigtMod.BackColor = Color.LawnGreen; }
                    break;

                case CmdEnablLogicvIn:
                    if (dataToAppd == StrDisable) { Bt_InvEnable.BackColor = Color.SandyBrown; }
                    else if (dataToAppd == StrEnable) { Bt_InvEnable.BackColor = Color.LawnGreen; }
                    break;

                case CmdAnalgInpt:
                    if (dataToAppd == StrDisable) { Bt_InvAnlg.BackColor = Color.SandyBrown; }
                    else if (dataToAppd == StrEnable) { Bt_InvAnlg.BackColor = Color.LawnGreen; }
                    break;

                case CmdRdLsrStatus:
                    break;

                case CmdSetPwMonOut:
                    break;

                case CmdRdCalDate:
                    break;

                case CmdSetPwCtrlOut:
                    dataToAppd = tb_SetIntPw.Text;
                    break;

                case CmdSetOffstVolt:
                    dataToAppd = Tb_SetOffset.Text;
                    break;

                case CmdSetVgaGain:
                    dataToAppd = Tb_VGASet.Text;
                    break;

                case CmdOperatingHr:
                    break;

                case CmdRdSummary:
                    break;

                case CmdSetInOutPwCtrl:
                    break;

                case CmdSet0mA:
                    break;

                case CmdSetStramind:
                    break;

                case CmdRdCmdStautus2:
                    break;

                case CmdManufDate:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdPwSetPcon:
                    break;

                case CmdRdInitCurrent:
                    break;

                case CmdRdModelName:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdLaserPow:
                    break;

                case CmdRdPnNb:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdCustomerPm:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRatedPower:
                    dataToAppd = Tb_NomPw.Text;
                    break;

                case CmdCurrentRead:
                    if(testMode==true) comThresh = 10;
                    break;

                case CmdSetPwtoVout:
                    dataToAppd = Tb_PwToVcal.Text;
                    sndDl = 600;
                    break;

                case CmdSetCalAPw:
                    dataToAppd = Tb_CalA_Pw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalBPw:
                    dataToAppd = Tb_CalB_Pw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalAVtoPw:
                    dataToAppd = Tb_CalAcmdToPw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalBVtoPw:
                    dataToAppd = Tb_CalBcmdToPw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRstTime:
                    break;

                case CmdSetSerNumber:
                    dataToAppd = Tb_SerNb.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetWavelenght:
                    dataToAppd = Tb_Wavelength.Text;
                    sndDl = 300;
                    break;

                case CmdSetLsMominalPw:
                    dataToAppd = Tb_NomPw.Text;
                    sndDl = 300;
                    break;

                case CmdSetCustomerPm:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetMaxIop:
                    dataToAppd = Tb_MaxLsCurrent.Text;
                    sndDl = 300;
                    break;

                case CmdSetCalDate:
                    dataToAppd = dateTimePicker1.Value.Date.ToString("ddMMyyyy");
                    sndDl = 600;
                    comThresh = 15;
                    break;

                case CmdSeManuDate:
                    dataToAppd = dateTimePicker1.Value.Date.ToString("ddMMyyyy");
                    sndDl = 600;
                    comThresh = 15;
                    break;

                case CmdSetModel:
                    dataToAppd = Lbl_MdlName.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdTestMode:
                    break;

                case CmdSetPSU:
                    break;

                case CmdSetBaseTempCal:
                    //compensation value for temperature
                    break;

                case CmdSetCalAPwtoVint:
                    dataToAppd = Tb_CalA_PwToADC.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalBPwtoVint:
                    dataToAppd = Tb_CalB_PwToADC.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetTECTemp:
                    dataToAppd = Tb_TECpoint.Text;
                    sndDl = 300;
                    break;

                case CmdSetTECkp:
                    dataToAppd = Tb_Kp.Text;
                    sndDl = 300;
                    break;

                case CmdSetTECki:
                    dataToAppd = Tb_Ki.Text;
                    sndDl = 300;
                    break;

                case CmdSetTECsmpTime:
                    dataToAppd = Tb_LoopT.Text;
                    sndDl = 300;
                    break;

                case CmdRdTECsetTemp:
                    break;

                case CmdRdTECsetkp:
                    break;

                case CmdRdTECsetki:
                    break;

                case CmdRdTECsmpTime:
                    break;

                case CmdSetTECena_dis:
                    break;

                default:
                    MessageBox.Show("No command Found to send");
                    break;
            }

            if (comThresh > 9) USB_CDC.ReceivedBytesThreshold = comThresh;
            else USB_CDC.ReceivedBytesThreshold = 9;

            bool result = await SendToSerial(cmdToTest, dataToAppd, sndDl);

            return true;
        }
        #endregion
        //======================================================================
        private async Task<bool> LoadGlobalTestArray(string[,] testListArr)// from test sequence
        {
            arrayLgth = (testListArr.Length) / 2; //how many steps in the bulk sequence

            for (int i = 0; i < arrayLgth;)
            {
                testStringArr[0] = testListArr[i, 0];
                testStringArr[1] = testListArr[i, 1];
                i++;
                bool result = await BuildSendString(testStringArr);
            }
            return true;
        }
        //======================================================================
        private static byte[] ConvertToByteArr(string str)
        {
            char[] charArr = str.ToCharArray();
            byte[] bytes = new byte[charArr.Length];

            for (int i = 0; i < charArr.Length; i++) {
                //byte current = Convert.ToByte(charArr[i]);
                int val1 = Convert.ToInt16(charArr[i]);
                bytes[i] = Convert.ToByte(val1 - 48);
            }
            return bytes;
        }
        //======================================================================
        #region COM buttons and address settings
        //======================================================================
        private void Bt_USB_Click(object sender, EventArgs e) {

            if (USB_CDC.IsOpen) {
                USB_CDC.Close();
                USB_CDC.Dispose();
                USB_Port_Open = false;
                Cb_USB.Enabled = true;
                Bt_RefrCOMs.Enabled = true;
                Bt_USBcom.BackColor = Color.SandyBrown;
                Bt_USBcom.Text = "USB Connect";
                Tb_USBConnect.BackColor = Color.Red;
                Bt_SetAddr.Enabled = false;
                Tb_EepromGood.BackColor = Color.Red;
                lbl_RdAdd.Text = "--";
            }
            else {
                if (RS232.IsOpen) {
                    RS232.Close();
                    RS232.Dispose();
                    RS232_Port_Open = false;
                    Cb_USB.Enabled = true;
                    Bt_RefrCOMs.Enabled = true;
                    Bt_Rs232com.BackColor = Color.SandyBrown;
                    Bt_Rs232com.Text = "RS Connect";
                    Tb_RSConnect.BackColor = Color.Red;
                 }

                try {
                    this.Cursor = Cursors.WaitCursor;
                    USB_CDC.PortName = Cb_USB.Text;
                    USB_CDC.Parity = Parity.None;
                    USB_CDC.StopBits = StopBits.One;
                    USB_CDC.DataBits = 8;
                    USB_CDC.BaudRate = 115200;
                    USB_CDC.ReadTimeout = 500;
                    USB_CDC.WriteTimeout = 500;
                    USB_CDC.ReceivedBytesThreshold = 9;
                    USB_CDC.Open();
                    USB_CDC.DiscardOutBuffer();
                    USB_CDC.DiscardInBuffer();
                    USB_Port_Open = true;
                    Cb_USB.Enabled = false;
                    Bt_USBcom.BackColor = Color.LawnGreen;
                    Bt_USBcom.Text = "USB Connected";
                    Bt_RefrCOMs.Enabled = false;
                    Bt_SetAddr.Enabled = true;
                    Task<bool> usbadd = SetAddress();
                    MessageBox.Show(" Connect PM100 \n" + " Connect USB IO interface \n");
                    //Reset_Form();
                }
                catch (Exception)
                {
                    USB_CDC.Close();
                    USB_CDC.Dispose();
                    USB_Port_Open = false;
                    Bt_USBcom.Text = "COM error";
                    Cb_USB.Enabled = true;
                    Bt_RefrCOMs.Enabled = true;
                    Bt_USBcom.BackColor = Color.SandyBrown;
                    Bt_USBcom.Text = "USB Connect";
                    Tb_USBConnect.BackColor = Color.Red;
                    Tb_EepromGood.BackColor = Color.Red;
                    lbl_RdAdd.Text = "--";
                    MessageBox.Show("USB_Port_Open COM Error");
                }
            }

            this.Cursor = Cursors.Default;
        }
        //======================================================================
        private void Bt_RS232_Click(object sender, EventArgs e)
        {
            if (RS232.IsOpen)
            {
                RS232.Close();
                RS232.Dispose();
                RS232_Port_Open = false;
                Cb_USB.Enabled = true;
                Bt_RefrCOMs.Enabled = true;
                Bt_Rs232com.BackColor = Color.SandyBrown;
                Bt_Rs232com.Text = "RS Connect";
                Tb_RSConnect.BackColor = Color.Red;
                Bt_USBcom.Enabled = true;
                Bt_SetAddr.Enabled = false;
                Tb_EepromGood.BackColor = Color.Red;
                lbl_RdAdd.Text = "--";
            }
            else
            {
                if (USB_CDC.IsOpen)
                {
                    USB_CDC.Close();
                    USB_CDC.Dispose();
                    USB_Port_Open = false;
                    Cb_USB.Enabled = true;
                    Bt_RefrCOMs.Enabled = true;
                    Bt_USBcom.BackColor = Color.SandyBrown;
                    Bt_USBcom.Text = "USB Connect";
                    Tb_USBConnect.BackColor = Color.Red;
                }
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    RS232.PortName = Cb_USB.Text;
                    RS232.Parity = Parity.None;
                    RS232.StopBits = StopBits.One;
                    RS232.DataBits = 8;
                    RS232.BaudRate = 115200;
                    RS232.ReadTimeout = 500;
                    RS232.WriteTimeout = 500;
                    RS232.ReceivedBytesThreshold = 9;
                    RS232.Open();
                    RS232.DiscardOutBuffer();
                    RS232.DiscardInBuffer();
                    RS232_Port_Open = true;
                    Bt_Rs232com.BackColor = Color.LawnGreen;
                    Bt_Rs232com.Text = "RS Connected";
                    Cb_USB.Enabled = false;
                    Bt_RefrCOMs.Enabled = false;
                    Bt_SetAddr.Enabled = true;
                }
                catch (Exception)
                {
                    RS232.Close();
                    RS232.Dispose();
                    RS232_Port_Open = false;
                    Bt_Rs232com.Text = "COM error";
                    Cb_USB.Enabled = true;
                    Bt_RefrCOMs.Enabled = true;
                    Bt_Rs232com.BackColor = Color.SandyBrown;
                    Bt_Rs232com.Text = "RS Connect";
                    Tb_RSConnect.BackColor = Color.Red;
                    Tb_EepromGood.BackColor = Color.Red;
                    lbl_RdAdd.Text = "--";
                    MessageBox.Show("RS232 COM Error");
                }
            }
            Task<bool> usbadd = SetAddress();
            this.Cursor = Cursors.Default;
        }
        //======================================================================
        private void Bt_Scan_Click(object sender, EventArgs e) { Getportnames(); }
        //======================================================================
        private void Getportnames()
        {
            this.Cursor = Cursors.WaitCursor;

            string[] portnames = SerialPort.GetPortNames();
            Cb_USB.Items.Clear(); //combo box ComConnect
            
            foreach (string s in portnames) Cb_USB.Items.Add(s);

            if (Cb_USB.Items.Count > 0) Cb_USB.SelectedIndex = 0;
            else Cb_USB.Text = "No COM port";

            this.Cursor = Cursors.Default;
        }
        //======================================================================
        private void Bt_SetAddr_Click(object sender, EventArgs e) { Task<bool> setadd = SetAddress(); }
        //======================================================================
        private async Task<bool> SetAddress()
        {
            string[] sentobuild = new string[2];
            sentobuild[1] = StrDisable;
            sentobuild[0] = CmdSetUnitNo;
            bool setad = await BuildSendString(sentobuild);//use polymorphism...
            sentobuild[0] = CmdRdUnitNo;
            setad = await BuildSendString(sentobuild);
            return true;
        }
        //======================================================================
        #endregion COM settings
        //======================================================================
        private void Rtb_Display_USB_DoubleClick(object sender, EventArgs e) { Rt_ReceiveDataUSB.Clear(); }
        //======================================================================
        private void Frm_iRIS_Prod_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (USB_CDC.IsOpen)
            {
                USB_CDC.Close();
                USB_CDC.Dispose();
            }
            if (RS232.IsOpen)
            {
                RS232.Close();
                RS232.Dispose();
            }
            //con.Close();
            Thread.Sleep(100);
            Application.Exit();
        }
        //======================================================================
        private void Form_iRIS_Clm_01_Load(object sender, EventArgs e)
        {
            //MessageBox.Show(" Enter Boards serial numbers\n Validate entry press ENTER\n", "Initial Set");
        }
        //======================================================================
       /*
                        int TbRow = 0;
                        con.Open();
                        cmd = new SqlCommand("select count(*) from " + dataBaseName , con);
                        TbRow = ((int)cmd.ExecuteScalar()) + 1;
                        iD = TbRow;
                        cmd = new SqlCommand("insert into " + dataBaseName + "(SerialNb, LaserDriverBdTestId, LaserId, TestDate, AnalogTestPass, DigitalTestPass, RS232Connet, USBConnect, LedOn, LaserDriverBdPgm, TecName)"
                            + " values(@BdSn, @LaserDriverBdTestId, @LaserId, @TestDate, @AnalogTestPass, @DigitalTestPass, @RS232Connet, @USBConnect, @LedOn, @LaserDriverBdPgm, @TecName)", con);
                        
                        cmd.Parameters.AddWithValue("@LaserDriverBdTestId", TbRow);
                        cmd.Parameters.AddWithValue("@LaserId", TbRow);
                        cmd.Parameters.AddWithValue("@TestDate", dateTimePicker1.Text);
                        cmd.Parameters.AddWithValue("@AnalogTestPass", 0);
                        cmd.Parameters.AddWithValue("@DigitalTestPass", 0);
                        cmd.Parameters.AddWithValue("@RS232Connet", 0);
                        cmd.Parameters.AddWithValue("@USBConnect", 0);
                        cmd.Parameters.AddWithValue("@LedOn", 0);
                        cmd.Parameters.AddWithValue("@LaserDriverBdPgm", 1);
                        cmd.Parameters.AddWithValue("@TecName", Tb_User.Text);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        DisplayData();
                        */
        //======================================================================
        private void Tb_TecSerNumb_MouseClick(object sender, MouseEventArgs e) { tb_TecSerNumb.Clear(); }
        //======================================================================
        private void Tb_TecSerNumb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                bool t = Information.IsNumeric(tb_TecSerNumb.Text);
                int strLgh = tb_TecSerNumb.Text.Length;

                if (t == true)
                {
                    if (strLgh < 9)
                    {
                        /*
                        con.Open();
                        cmd = new SqlCommand("update " + dataBaseName + " set TEC_Board_Sn = @tecSn where LaserId = @LaserId", con);
                        cmd.Parameters.AddWithValue("@LaserId", iD);
                        cmd.Parameters.AddWithValue("@tecSn", tb_TecSerNumb.Text);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        DisplayData();
                        */
                        tb_TecSerNumb.BackColor = Color.White;
                        tb_TecSerNumb.Enabled = false;
                    }
                    else { MessageBox.Show("8 Digits Maximum"); }
                }
                else { MessageBox.Show("Provide Details!"); }
            }
        }
        //======================================================================
        private void Tb_User_MouseClick(object sender, MouseEventArgs e) { Tb_User.Clear(); }
        //======================================================================
        private void Tb_User_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                
            }
        }
        //======================================================================
        private void Bt_EnableDBstring_Click(object sender, EventArgs e)
        {
            if (Bt_EnableDBstring.BackColor == Color.White)
            {
                string answBox = Microsoft.VisualBasic.Interaction.InputBox(" Enter Password", "SQL ACCESS");

                if (answBox == "Qioptiq") {
                    con.Close();
                    Bt_EnableDBstring.BackColor = Color.LawnGreen;
                    label36.Visible = true;
                    label37.Visible = true;
                    label38.Visible = true;
                    label39.Visible = true;
                    label40.Visible = true;
                    Tb_DatabaseString.Visible = true;
                    Tb_DatabaseString.Visible = true;
                    Tb_ServerName.Visible = true;
                    Tb_InitialCatalog.Visible = true;
                    Tb_User1.Visible = true;
                    Tb_Pw1.Visible = true;
                }
                else MessageBox.Show("Wrong Pass Word");
            }
            else if (Bt_EnableDBstring.BackColor == Color.LawnGreen) {
                Bt_EnableDBstring.BackColor = Color.White;
                label36.Visible = false;
                label37.Visible = false;
                label38.Visible = false;
                label39.Visible = false;
                label40.Visible = false;
                Tb_DatabaseString.Visible = false;
                Tb_ServerName.Visible = false;
                Tb_InitialCatalog.Visible = false;
                Tb_User1.Visible = false;
                Tb_Pw1.Visible = false;
                OpenSqlConnection();
            }
        }
        //======================================================================
        #region // PM100 Code...
        //======================================================================
        private void Bt_PM100_Click(object sender, EventArgs e) { Task<bool> pm100bt = PM100Button(); }
        //======================================================================
        private async Task<bool> PM100Button()
        {
            if (Bt_PM100.BackColor == Color.Coral)
            {
                pm100ok = await InitialiseTestHardwarePM();

                if (pm100ok == true)
                {
                    Bt_PM100.Text = "PM100 Connected";
                    Bt_PM100.BackColor = Color.LawnGreen;
                }
                else if (pm100ok == false) { MessageBox.Show("PM100 Not Connected"); }
            }
            else if (Bt_PM100.BackColor == Color.LawnGreen)
            {
                Bt_PM100.Text = "PM100 Connect";
                Bt_PM100.BackColor = Color.Coral;
                pm.reset();
                pm.Dispose();
                pm100ok = false;
            }
            return true;
        }
        //======================================================================
        private async Task<bool> InitialiseTestHardwarePM()
        {
            this.Cursor = Cursors.WaitCursor;
            uint rsrc = 0;
            bool pm100cnt = false;
            string pm100InitString = CmBx_PM100str.Text;

            try
            {
                pm = new PM100D(pm100InitString, false, true);
                pm.findRsrc(out rsrc);

                if (rsrc != 0)
                {
                    MessageBox.Show("Thorlab resources " + Convert.ToString(rsrc));
                    pm.setWavelength(Convert.ToInt16(Tb_Wavelength.Text));
                    //pm.setPowerAutoRange(true);
                    pm100cnt = true;
                }
                else
                {
                    pm100cnt = false;
                    MessageBox.Show("No PM100D");
                }
            }

            catch (Exception e) { MessageBox.Show("PM100 Error" + e.ToString()); }
                    
            await Task.Delay(1);
            this.Cursor = Cursors.Default;
            return pm100cnt;
        }
        //=====================================================================
        private void Bt_RdPM100_Click(object sender, EventArgs e) { ReadPM100();  }
        //=====================================================================
        private double ReadPM100()//in mW
        {
            string pwrStr = string.Empty;
            double powerRd = 0;

            if (pm100ok == true)
            {
                pm.measPower(out double power);
                powerRd = Math.Round((power * 1000), 4);
                pwrStr = powerRd.ToString("00.000");
                Lbl_PM100rd.Text = pwrStr;
            }
            else if (pm100ok == false) { MessageBox.Show("PM100 Error"); }

            return powerRd;
        }
        //=====================================================================
        #endregion
        //======================================================================
        #region // USB Interface Code...
        private void Bt_USBinterf_Click(object sender, EventArgs e) { Task usbInter = SetUsbInterface(); }
        //======================================================================
        private async Task<bool> SetUsbInterface()
        {
            if (Bt_USBinterf.BackColor == Color.Coral)
            {
                bool boardFound = false;
                Int16 BoardNum = 99;

                MccDaq.DaqDeviceManager.IgnoreInstaCal();
                inventory = MccDaq.DaqDeviceManager.GetDaqDeviceInventory(MccDaq.DaqDeviceInterface.Any);
                Int32 numDevDiscovered = inventory.Length;

                if (numDevDiscovered > 0)
                {
                    for (BoardNum = 0; BoardNum < numDevDiscovered; BoardNum++)
                    {
                        try
                        {
                            DaqBoard = MccDaq.DaqDeviceManager.CreateDaqDevice(BoardNum, inventory[BoardNum]);
                            if (DaqBoard.BoardName.Contains(CmBx_USBinterface.Text))
                            {
                                boardFound = true;
                                DaqBoard.FlashLED();
                                break;
                            }
                            else {MccDaq.DaqDeviceManager.ReleaseDaqDevice(DaqBoard); }
                        }
                        catch (MccDaq.ULException ule) { System.Windows.Forms.MessageBox.Show(ule.Message, "No USB-3103 found in system. Run InstaCal"); }
                    }
                }
                else {
                    System.Windows.Forms.MessageBox.Show("No Board detected");
                    this.Close(); }

                if (boardFound == true)
                {
                    ULStat = DaqBoard.DConfigPort(DigitalPortType.AuxPort, DigitalPortDirection.DigitalIn);
                    ULStat = DaqBoard.DConfigBit(DigitalPortType.AuxPort, 0, DigitalPortDirection.DigitalOut);
                    ULStat = DaqBoard.DConfigBit(DigitalPortType.AuxPort, 1, DigitalPortDirection.DigitalOut);
                    ULStat = DaqBoard.DConfigBit(DigitalPortType.AuxPort, 2, DigitalPortDirection.DigitalIn);
                    ULStat = DaqBoard.DConfigBit(DigitalPortType.AuxPort, 3, DigitalPortDirection.DigitalIn);

                    DaqBoard.BoardConfig.GetNumAdChans(out numchannels);
                    nudAInChannel = numchannels - 1;
                    Range = MccDaq.Range.Uni5Volts;
                    RangeSelected = (MccDaq.Range)(0);
                    DaqBoard.AInputMode(MccDaq.AInputMode.SingleEnded);

                    String mystring = DaqBoard.BoardName.Substring(DaqBoard.BoardName.Trim().Length) +
                    " board number: " + BoardNum.ToString() + nudAInChannel.ToString();
                    Text = mystring;

                    MessageBox.Show(Text + " " + "");
                }
                Bt_USBinterf.BackColor = Color.LawnGreen;
                Bt_USBinterf.Text = "USB Interface Connected";
            }

            else if (Bt_USBinterf.BackColor == Color.LawnGreen) {
                MccDaq.DaqDeviceManager.ReleaseDaqDevice(DaqBoard);
                Bt_USBinterf.BackColor = Color.Coral;
                Bt_USBinterf.Text = "USB Interface Connect"; }

            await Task.Delay(1);
            return true;
        }
        //======================================================================
        private void Set_USB_Digit_Out(sbyte portNb, sbyte state)
        {
         if (state == 1) {
                if (portNb == 0) { Bt_LsEnable.BackColor = Color.Plum; }
                else if (portNb == 1) { Bt_DigMod.BackColor = Color.Plum; }
                ULStat = DaqBoard.DBitOut(DigitalPortType.AuxPort, portNb, DigitalLogicState.High);
            }
                
         else if (state == 0) {
                if (portNb==0) { Bt_LsEnable.BackColor = Color.PeachPuff; }
                else if (portNb == 1) { Bt_DigMod.BackColor = Color.PeachPuff; }
                ULStat = DaqBoard.DBitOut(DigitalPortType.AuxPort, portNb, DigitalLogicState.Low);
            }
                
        }
        //======================================================================
        private sbyte Read_USB_Digit_in(sbyte portNb)//0 to 7
        {
        ULStat = DaqBoard.DBitIn(DigitalPortType.AuxPort, portNb, out DigitalLogicState digiIn);
        return ((sbyte)digiIn);
        }
        //======================================================================
        private void Bt_LsEnable_Click(object sender, EventArgs e)
        {
            if (Bt_LsEnable.BackColor == Color.PeachPuff) Set_USB_Digit_Out(0, 1);
            else Set_USB_Digit_Out(0, 0);
        }
        //======================================================================
        private void Bt_DigMod_Click(object sender, EventArgs e)
        {
            if (Bt_DigMod.BackColor==Color.PeachPuff) Set_USB_Digit_Out(1, 1);
            else Set_USB_Digit_Out(1, 0);
        }
        //======================================================================
        private void Bt_RsLaserOk_Click(object sender, EventArgs e) {
            sbyte lsOK = 0;
            lsOK = Read_USB_Digit_in(2);
            if (lsOK == 1) Tb_LaserOK.BackColor = Color.Green;
            else Tb_LaserOK.BackColor = Color.Red; }
        //======================================================================
        private void WriteDAC(double dacValue, int dacChannel)//value in volts
        {
            double dacValue1 = 0;
            dacValue1 = dacValue * 0.5;//unipolar convertion
            ULStat = DaqBoard.VOut(dacChannel, RangeSelected, (float)dacValue1, MccDaq.VOutOptions.Default);
            if(dacChannel == 0) { Tb_VPcon.Text = dacValue.ToString("00.000"); }
        }
        //======================================================================
        private void Bt_SetPcon_Click(object sender, EventArgs e)//values in volts
        {
            double pconV = Convert.ToDouble(Tb_VPcon.Text);
            WriteDAC(pconV, 0);//Pcon Channel
        }
        //======================================================================
        private double ReadADC(int adcChannel)//returns Volts
        {
            Range = MccDaq.Range.Bip10Volts;//connect ch low to AGND
            ULStat = DaqBoard.VIn32(Convert.ToInt16(adcChannel), Range, out double VInVolts, MccDaq.VInOptions.Default);
            return VInVolts;
        }
        //======================================================================
        private void Bt_RdAnlg_Click_1(object sender, EventArgs e) { Task<bool> rdadc = ReadAllanlg(false); }
        //======================================================================
        #endregion  External Hardware
        //======================================================================
        private async Task<bool> RampDAC1(double startRp, double stopRp, double stepRp)//external PCON
        {
            double maxPw = Convert.ToDouble(Tb_maxMaxPw.Text);
            bool rampDAC1task = false;//cannot initialise in "for loop" ????

            int arrIndex = Convert.ToInt16((stopRp - stopRp) / stepRp);
            double[,] dataADC = new double[arrIndex, 3];//could "dynamically" size array....
            dataADC.Initialize();

            for (double startRpLp = startRp; startRpLp <= stopRp; startRpLp = startRpLp + stepRp)
            {
                    WriteDAC(startRpLp, 0);
                    rampDAC1task = await ReadAllanlg(false);//displays current in bits

                double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW
                if (pm100Res > maxPw) {
                    WriteDAC(0, 0);
                    MessageBox.Show("Power Error");
                    return false; }//ramp error

                //populate array with results
                //dataADC[arrIndex, 0] = pm100Res;
                //dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                //dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                //arrIndex++;
            }

            //Rt_ReceiveDataUSB.Clear();
            //foreach (double dbl in dataADC) {
            //    Rt_ReceiveDataUSB.AppendText(Convert.ToString(dbl));
            //    Rt_ReceiveDataUSB.AppendText("\n"); }

            return true;
        }
        //======================================================================
        private async Task<bool> ReadAllanlg(bool fullRd) {//reads all data

            double pwrRead = 0;             //pm100
            double pconRead = ReadADC(0);   //PCON feedback
            double lsrPwRead = ReadADC(1);  //PD Vout
            double lsrCurrRead = ReadADC(2);//Current Vout

            //display ADC results**
            Lbl_Vpcon.Text = pconRead.ToString("00.000");
            Lbl_PwreadV.Text = lsrPwRead.ToString("00.000");//*294.12
            if (testMode == true)
            {//test mode
                Lbl_Viout.Text = lsrCurrRead.ToString("00.000");
                Lbl_Ma.Text = "Laser I in V /5";
            }
            else if (testMode == false)
            {//run mode
                string currentMa = Convert.ToString(lsrCurrRead * 200);
                int endIndex = currentMa.LastIndexOf(".");
                Lbl_Viout.Text = currentMa.Substring(0, endIndex);
                Lbl_Ma.Text = "Laser I in mA";
            }
 
            if (pm100ok == true) { pwrRead = ReadPM100();
                Lbl_PM100rd.Text = pwrRead.ToString("00.000"); }//in mW

            if (fullRd == true) { bool readAdc = await LoadGlobalTestArray(analogRead); }//internal uCadc

            await Task.Delay(50);//this is there as the compiler will not see the await in the if statement.

            return true;
        }
        //======================================================================
        private void Bt_NewTest_Click(object sender, EventArgs e)
        {
            if (USB_Port_Open == true) { Task<bool> test2 = FirtInit(); }
            else MessageBox.Show("USB not connected");
        }
        //======================================================================
        private async Task<bool> FirtInit()
        {
            MessageBox.Show(" Power ON/12V-2A \n");

            bt_NewTest.BackColor = Color.LawnGreen;
            this.Cursor = Cursors.WaitCursor;

            bool test2 = await LoadGlobalTestArray(bulkSetLaserIO);
            //test2 = await LoadGlobalTestArray(bulkSetVarialble);
            //test2 = await LoadGlobalTestArray(bulkSetdefaultCtrl);
            //test2 = await LoadGlobalTestArray(bulkSetTEC);

            Tb_VGASet.Text = "0020";
            Tb_SetOffset.Text = "02.500";
            Tb_VPcon.Text = "00.000";

            this.Cursor = Cursors.Default;
            MessageBox.Show("\n" + "Wait for TEC lock LED");
            //bt_NewTest.BackColor = Color.Coral;
            return true;
        }
        //======================================================================
        private void Bt_CalVGA_Click(object sender, EventArgs e) {
            //if green stop ramp....
            Task<bool> calvga = CalVGA();
            //if (USB_Port_Open == true) { Task<bool> calvga = CalVGA(); }
            //else MessageBox.Show("USB not connected");
        }
        //======================================================================
        private void Bt_SetVGA_Click(object sender, EventArgs e) { Task<bool> setVGA = SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300); }
        //======================================================================
        private async Task<bool> CalVGA()
        {
            Bt_CalVGA.BackColor = Color.LawnGreen;
            this.Cursor = Cursors.WaitCursor;

            double calPower = 0;
            double setOffSet = 0;
            double setPower = Convert.ToDouble(Tb_NomPw.Text);

            string offset = string.Empty;
            bool initvga = false;

            initvga = await LoadGlobalTestArray(bulkSetVga);
            Set_USB_Digit_Out(0, 0);//enable
            Set_USB_Digit_Out(1, 0);
            Tb_VPcon.Text = "00.000";
            WriteDAC(0, 0);

            if (Convert.ToBoolean(Read_USB_Digit_in(0)) == false) {//Laser OK //test

                initvga = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300);//default VGA gain 20
                initvga = await SendToSerial(CmdLaserEnable, StrEnable, 300);//Laser Enable
                initvga = await SendToSerial(CmdSetOffstVolt, Tb_SetOffset.Text, 300);
                Set_USB_Digit_Out(0, 1);//Laser Enable
                
                initvga = await ReadAllanlg(true);//test if OK
                MessageBox.Show("Start VGA Cal");

                for (int i = 0; i <= 2; i++)//3 VGA set iteration //test
                {
                    MessageBox.Show("Ramp" + i.ToString());
                    bool boolCalVGA1 = await RampDAC1(0, 4.950, 0.05);//set VGA MAX power

                        for (int vgaVal = 20; vgaVal <= 80;)
                        {
                            Tb_VGASet.Text = vgaVal.ToString("0000");
                            bool vgaset = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300);

                            vgaset = await ReadAllanlg(false);

                            double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW
                            double maxPw = Convert.ToDouble(Tb_minMaxPw.Text);

                            if (pm100Res >= maxPw)
                            {
                                MessageBox.Show("VGA set pass");
                                break;
                            }

                            if (vgaVal >= 80)
                            {
                                MessageBox.Show("VGA error");
                                break;
                            }

                            vgaVal++;
                        }

                    WriteDAC(0.500, 0);//set to 0.5V PCON

                    for (int j = 0; j < 59; j++)
                    {
                        bool vgaset02 = await ReadAllanlg(false);
                        calPower = Convert.ToDouble(Lbl_PM100rd.Text);//mW @ 0.1%
                        setOffSet = Convert.ToDouble(Tb_SetOffset.Text);//offset value re-initialise for new test
 
                        if (calPower > (setPower * 0.00125) )//upper limit over 0.1% + 25% max power
                        {
                            setOffSet = setOffSet + 0.002;//add offset...reduces power
                        }

                        else if (calPower < (setPower * 0.00075) )//between 0.1%Pw and 0.1%Pw-0.02 lower limit
                        {
                            setOffSet = setOffSet - 0.002;//increase power
                        }
                        offset = setOffSet.ToString("0.000");
                        Tb_SetOffset.Text = offset;
                        bool boolCalVGA5 = await SendToSerial(CmdSetOffstVolt, offset, 400);//update offset
                    }
                }
            }
            else MessageBox.Show("Laser NOT OK");

            initvga = await SendToSerial(CmdLaserEnable, StrDisable, 300);//end VGA stop test
            WriteDAC(0, 0);

            Bt_CalVGA.BackColor = Color.Coral;
            this.Cursor = Cursors.Default;

            return true;
        }
 
        //======================================================================
        private async Task<bool> RampVGA()
        {
            int vgaVal = Convert.ToInt16(Tb_VGASet.Text);//need to be initialised (20)

            for (; vgaVal <= 79; ) {

                Tb_VGASet.Text = vgaVal.ToString("0000");
                bool vgaset = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300);

                vgaset = await ReadAllanlg(false);

                double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW
                double maxPw = Convert.ToDouble(Tb_minMaxPw.Text);

                if (pm100Res >= maxPw){
                    MessageBox.Show("VGA set pass");
                    break; }

                if (vgaVal >= 80){
                    MessageBox.Show("VGA error");
                    break; }

                vgaVal++;
            }

            return true;
        }
        //======================================================================
        #region Current Zero 
        //======================================================================
        private void Bt_ReaduCcurrent_Click(object sender, EventArgs e) { Task<bool> readuCLsCurrent = SendToSerial(CmdCurrentRead, StrDisable, 300); }
        //======================================================================
        private void Bt_ZeroI_Click(object sender, EventArgs e) { Task<bool> zeroI = ZerroCurrent(); }
        //======================================================================
        private async Task<bool> ZerroCurrent() {

            Bt_ZeroI.BackColor = Color.LawnGreen;
            Set_USB_Digit_Out(0, 1);//Laser OFF
            WriteDAC(0, 0);//Pcon Channel = 0V
           
            bool rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300);//read current value from cpu displayed on label
            double lsrCurrRead = ReadADC(2);//Current monitor voltage

            rdIcal = await SendToSerial(CmdSet0mA, StrDisable ,600);//zero value cal

            rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300);//recheck new cpu value...same voltage offset at Imon OUT

            return true;
        }
        #endregion
        //======================================================================
        private void TbPg_InsTest_Click(object sender, EventArgs e) { Task<bool> calPwMonOut = CalPwMonOut(); }
        //======================================================================
        private async Task<bool> CalPwMonOut()
        {

            await Task.Delay(1);

            return true;
        }
        //======================================================================
        private void Bt_pdCalibration_Click(object sender, EventArgs e) { Task<bool> pdcal = PD_Calibration(); }
        //======================================================================
        private async Task<bool> PD_Calibration()
        {
            bool pdCalTask = false;
            Bt_pdCalibration.BackColor = Color.LawnGreen;
            Cursor.Current = Cursors.WaitCursor;

            pdCalTask = await LoadGlobalTestArray(bulkSetVarialble);
            
            Set_USB_Digit_Out(0, 1);                  //Enable laser  
            pdCalTask = await SendToSerial(CmdLaserEnable, StrEnable, 300);// test disable

            pdCalTask = await RampDAC1(0, 5.000, 0.1);

            Set_USB_Digit_Out(0, 0);                    
            pdCalTask = await SendToSerial(CmdLaserEnable, StrDisable, 300);

            Cursor.Current = Cursors.Default;
            return true;
        }
        //======================================================================
        private void Bt_FinalLsSetup_Click(object sender, EventArgs e)
        {
            Task<bool> endSetup = LsFinalSet();
  
        }
        //======================================================================
        private async Task<bool> LsFinalSet()
        {

            await Task.Delay(1);

            return true;
        }
        //======================================================================
        private void Bt_LaserEn_Click_1(object sender, EventArgs e) {
            if (Bt_LaserEn.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdLaserEnable, StrEnable, 300);
                Bt_LaserEn.BackColor = Color.LawnGreen;
            }
            else if (Bt_LaserEn.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdLaserEnable, StrDisable, 300);
                Bt_LaserEn.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_InvDigtMod_Click_1(object sender, EventArgs e) {
            if (Bt_InvDigtMod.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdsetTTL, StrEnable, 300);
                Bt_InvDigtMod.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvDigtMod.BackColor == Color.LawnGreen){
                Task<bool> sdEnd = SendToSerial(CmdsetTTL, StrDisable, 300);
                Bt_InvDigtMod.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_InvAnlg_Click_1(object sender, EventArgs e) {
            if (Bt_InvAnlg.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdAnalgInpt, StrEnable, 300);
                Bt_InvAnlg.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvAnlg.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdAnalgInpt, StrDisable, 300);
                Bt_InvAnlg.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_IntExtPw_Click_1(object sender, EventArgs e) {
            if (Bt_IntExtPw.BackColor == Color.SandyBrown)  {
                Task<bool> sdEne = SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300);
                Bt_IntExtPw.BackColor = Color.LawnGreen;
            }
            else if (Bt_IntExtPw.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdSetInOutPwCtrl, StrDisable, 300);
                Bt_IntExtPw.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_SetIntPwDAC_Click(object sender, EventArgs e) { Task<bool> sdEne = SendToSerial(CmdSetPwCtrlOut, tb_SetIntPw.Text, 600); }
        //======================================================================
        private void Bt_EnableTest_Click(object sender, EventArgs e) {
            if (Bt_EnableTest.BackColor == Color.SandyBrown) {
                Bt_EnableTest.BackColor = Color.LawnGreen;
                Task<bool> sdEne = SendToSerial(CmdTestMode, StrEnable, 300);
            }
            else if (Bt_EnableTest.BackColor == Color.LawnGreen) {
                Bt_EnableTest.BackColor = Color.SandyBrown;
                Task<bool> sdEnd = SendToSerial(CmdTestMode,StrDisable,300);
            }
        }
        //======================================================================
        private void Bt_setOffDac_Click(object sender, EventArgs e) { Task<bool> sdEne = SendToSerial(CmdSetOffstVolt, Tb_SetOffset.Text, 600); }
        //======================================================================
        private void Bt_SetPwDac_Click(object sender, EventArgs e) { Task<bool> sdPddac = SendToSerial(CmdSetPwMonOut, Tb_PwToVout.Text, 600); }
        //======================================================================
        private void Bt_InvEnable_Click(object sender, EventArgs e)
        {
            if (Bt_InvEnable.BackColor == Color.SandyBrown)
            {
                Task<bool> sdEne = SendToSerial(CmdEnablLogicvIn, StrEnable, 300);
                Bt_InvEnable.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvEnable.BackColor == Color.LawnGreen)
            {
                Task<bool> sdEnd = SendToSerial(CmdEnablLogicvIn, StrDisable, 300);
                Bt_InvEnable.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        #region Calibrate Power Monitor Output
        private void Bt_PwOutMonCal_Click(object sender, EventArgs e) { Task<bool> runPwCal = PwMonOutCal(); }
        //======================================================================
        private async Task<bool> PwMonOutCal()
        {
            Bt_PwOutMonCal.BackColor = Color.LawnGreen;

            string pmonVmax = Tb_PwToVcal.Text;
            WriteDAC(00.000, 0);
            bool sendCalPw = await SendToSerial(CmdTestMode, StrEnable, 300);
            sendCalPw = await SendToSerial(CmdLaserEnable, StrEnable, 300);
            Set_USB_Digit_Out(0, 1);

            bool rampdac1 = await RampDAC1(0, 5.000, 0.100);//adjust PCON to MAX power

            sendCalPw = await SendToSerial(CmdSetPwtoVout, pmonVmax, 600);
            sendCalPw = await ReadAllanlg(true);

            MessageBox.Show("Pw Mon. Vmax");

            Set_USB_Digit_Out(0, 0);
            sendCalPw = await SendToSerial(CmdLaserEnable, StrDisable, 300);
            WriteDAC(00.000, 0);
 
            sendCalPw = await ReadAllanlg(true);

            MessageBox.Show("Pw Mon. Vmin");
 
            return true;
        }
        #endregion
        //======================================================================
        #region Temp Comp Base plate
        //======================================================================
        private void Bt_BasepltTemp_Click(object sender, EventArgs e) { Task<bool> readtempBplt = SendToSerial(CmdRdBplateTemp, StrDisable, 300); }
        //======================================================================
        private void Bt_BasePltTempComp_Click(object sender, EventArgs e) { Task<bool> setTcomp = CompBpltTemp(); }
        //======================================================================
        private async Task<bool> CompBpltTemp() {

            Bt_BasepltTemp.BackColor = Color.LawnGreen;
            bool setCompT =     await SendToSerial(CmdSetBaseTempCal, "0000", 300);                         //set init comp to 0000 remember to reset for next init.
            setCompT =          await SendToSerial(CmdRdBplateTemp, StrDisable, 300);                       //read initial value
            
            int measTemp =  ReadExtTemp();                                                                  //get user temp //wait
            int tempComp1 = Convert.ToInt16(Lbl_TempBplt.Text) - measTemp;
            setCompT =         await SendToSerial(CmdSetBaseTempCal, tempComp1.ToString("0000"), 300);      //set init comp to 0000 remember to reset for next init.
            
            setCompT =      await SendToSerial(CmdRdBplateTemp, StrDisable, 300);                            //read comp data
            return true;
        }
         //======================================================================
        private int ReadExtTemp()//user thermometer reading
        {
            int strpopupInt = 0;
            Getitright:
            string strpopup = Microsoft.VisualBasic.Interaction.InputBox(" Enter Temperature in 1/10C \n", "Base Plate Temperature Compensation", "000");

            bool t = Information.IsNumeric(strpopup);
            int strLgh = strpopup.Length;

            if (t == true) {
                if (strLgh < 4) { strpopupInt = Convert.ToInt16(strpopup); }
                else {
                    MessageBox.Show("3 Digits Maximum");
                    goto Getitright; }
            }
            else {
                MessageBox.Show("Numerical only");
                goto Getitright;
            }

            return strpopupInt;
        }
        //======================================================================
        #endregion
        //======================================================================
        private double[] FindLinearLeastSquaresFit(double[,] dataXy, int strtX, int endX)
        {
            double[] rtnAb = new double[2];
            double S1 = 0;
            double Sx = 0;
            double Sy = 0;
            double Sxx = 0;
            double Sxy = 0;

            //if (ChkBx_B.Checked == true)
            //{
                for (int lp2 = strtX; lp2 <= endX; lp2++)
                {
                    Sx += dataXy[lp2, 0];//x
                    Sy += dataXy[lp2, 2];//y

                    Sxx += dataXy[lp2, 0] * dataXy[lp2, 0];//xx
                    Sxy += dataXy[lp2, 0] * dataXy[lp2, 2];//xy

                    Rt_ReceiveDataUSB.AppendText(   Convert.ToString(dataXy[lp2, 0]) + " " +
                                                    Convert.ToString(dataXy[lp2, 2]) + " " +
                                                    Footer);
                    S1++;
                }

                double m = (Sxy * S1 - Sx * Sy) / (Sxx * S1 - Sx * Sx);
                double b = (Sxy * Sx - Sy * Sxx) / (Sx * Sx - S1 * Sxx);

                rtnAb[0] = m;
                rtnAb[1] = b;
            //}
            /*
            else if (ChkBx_B.Checked == false)//default 
            {
                for (int lp2 = strtX; lp2 <= endX; lp2++)
                {
                    Sxx += dataXy[lp2, 0] * dataXy[lp2, 0];//xx sum square x
                    Sxy += dataXy[lp2, 0] * dataXy[lp2, 2];//xy sum product xy

                    Rtb_Snoop.AppendText(Convert.ToString(dataXy[lp2, 0]) + " " +
                                         Convert.ToString(dataXy[lp2, 2]) + " " +
                                            iRIScoms.Footer);
                    S1++;
                }

                double m = (Sxy) / (Sxx);

                rtnAb[0] = m;
                rtnAb[1] = 0;
            }
            */
            return rtnAb;
        }
        //======================================================================
    }
    //======================================================================
    //======================================================================
}
//======================================================================
//======================================================================

/*
namespace ConsoleApplication1
{

using System.Text;
using System.Data.Odbc;
using System.Data;
using System.Web;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Linq;

using System;
using System.Collections.Generic;
using System.Management; // need to add System.Management to your project references.

class Program
{

    static void Main(string[] args)
    {
        var usbDevices = GetUSBDevices();

        foreach (var usbDevice in usbDevices)
        {
            string m_pendid;

            Console.WriteLine("Device ID: {0}, PNP Device ID: {1}, Description: {2}, USBVersion: {3}, SystemName: {4}",
            usbDevice.DeviceID, usbDevice.PnpDeviceID, usbDevice.Description, usbDevice.usbversion, usbDevice.SystemName);

            // m_pendid=catch["usbDevice.DeviceID"];
            m_pendid = usbDevice.DeviceID;

            Console.WriteLine("Test" + m_pendid);

        }

        // Console.Write("DeviceID :DeviceID");
        Console.Read();

    }

    static List<usbdeviceinfo> GetUSBDevices()
    {
        List<usbdeviceinfo> devices = new List<usbdeviceinfo>();

        ManagementObjectCollection collection;
        using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))
            collection = searcher.Get();

        foreach (var device in collection)
        {
            devices.Add(new USBDeviceInfo(
            (string)device.GetPropertyValue("DeviceID"),
            (string)device.GetPropertyValue("PNPDeviceID"),
            (string)device.GetPropertyValue("Description"),
            (string)device.GetPropertyValue("USBVersion"),
            (string)device.GetPropertyValue("SystemName")

            ));

        }

        collection.Dispose();
        return devices;
    }
}

class USBDeviceInfo
{
    public USBDeviceInfo(string deviceID, string pnpDeviceID, string description, string usbversion1, string SystemName2)
    {
        this.DeviceID = deviceID;
        this.PnpDeviceID = pnpDeviceID;
        this.Description = description;
        this.usbversion = usbversion1;
        this.SystemName = SystemName2;
    }
    public string DeviceID { get; private set; }
    public string PnpDeviceID { get; private set; }
    public string Description { get; private set; }
    public string usbversion { get; private set; }
    public string SystemName { get; private set; }
}
}

*/
