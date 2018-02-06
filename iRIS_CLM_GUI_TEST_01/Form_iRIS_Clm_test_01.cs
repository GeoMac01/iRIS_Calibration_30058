using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.IO.Ports;
using System.IO;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Thorlabs.PM100D_32.Interop;
using MccDaq;

//iRIS Production 30058_01
//15/01/2018

namespace iRIS_CLM_GUI_TEST_01
{
    public partial class Form_iRIS_Clm_test_01 : Form
    {
        #region Commands Definition
        const string rtnNull = "00";
        const string CmdLaserEnable = "02";
        const string CmdSetLsPw = "03";
        const string CmdRdSerialNo = "04";
        const string CmdRdFirmware = "06";
        const string CmdRdBplateTemp = "07";
        const string CmdRdWavelen = "08";
        const string CmdSetUnitNo = "12";
        const string CmdLaserStatus = "14";
        const string CmdRdTecTemprt = "15";
        const string CmdHELP = "16";
        const string CmdsetTTL = "17";
        const string CmdEnablLogicvIn = "18";
        const string CmdAnalgInpt = "19";
        const string CmdRdLsrStatus = "20";
        const string CmdSetPwMonOut = "21";
        const string CmdRdCalDate = "22";
        const string CmdSetPwCtrlOut = "23";
        const string CmdSetOffstVolt = "24";
        const string CmdSetVgaGain = "25";
        const string CmdOperatingHr = "26";
        const string CmdRdSummary = "27";
        const string CmdSetInOutPwCtrl = "28";
        const string CmdSet0mA = "29";
        const string CmdSetStramind = "30";
        const string CmdRdCmdStautus2 = "34";
        const string CmdManufDate = "40";
        const string CmdRdPwSetPcon = "41";
        const string CmdRdInitCurrent = "42";
        const string CmdRdModelName = "43";
        const string CmdRdLaserPow = "44";
        const string CmdRdPnNb = "45";
        const string CmdRdCustomerPm = "46";
        const string CmdRatedPower = "47";
        const string CmdCurrentRead = "56";
        const string CmdSetPwtoVout = "59";
        const string CmdSetCalAPw = "60";
        const string CmdSetCalBPw = "61";
        const string CmdSetCalAPwtoVint = "62";
        const string CmdSetCalBPwtoVint = "63";
        const string CmdRstTime = "66";
        const string CmdSetSerNumber = "71";
        const string CmdSetWavelenght = "72";
        const string CmdSetLsMominalPw = "73";
        const string CmdSetCustomerPm = "74";
        const string CmdSetMaxIop = "76";
        const string CmdSetCalDate = "77";
        const string CmdSeManuDate = "78";
        const string CmdSetPartNumber = "79";
        const string CmdSetModel = "80";
        const string CmdSetCalAVtoPw = "81";
        const string CmdSetCalBVtoPw = "82";
        const string CmdTestMode = "83";
        const string CmdSetPSU = "84";
        const string CmdReadPSU = "86";
        const string CmdSetBaseTempCal = "87";
        const string CmdSetTECTemp = "90";
        const string CmdSetTECkp = "91";
        const string CmdSetTECki = "92";
        const string CmdSetTECsmpTime = "93";
        const string CmdRdTECsetTemp = "94";
        const string CmdRdTECsetkp = "95";
        const string CmdRdTECsetki = "96";
        const string CmdRdTECsmpTime = "97";
        const string CmdSetTECena_dis = "98";
        const string CmdRdUnitNo = "99";
        const string Footer = "\r\n";
        const string Header = "#";
        const string StrEnable = "0001";
        const string StrDisable = "0000";
        #endregion

        //=================================================
        #region Test Sequence Definition
        //=================================================

        string[,] bulkSetLaserIO = new string[7, 2] {   //the rest of the string is build with case...
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetTECena_dis,     StrDisable},
            { CmdSetInOutPwCtrl,    StrDisable },       //external PCON
            { CmdAnalgInpt,         StrDisable },       //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },       //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable } };      //Inv. TTL line in nothing connected

        string[,] bulkSetVarialble = new string[14, 2] {
            {CmdTestMode,           StrEnable },
            {CmdSetWavelenght,      StrDisable},
            {CmdSetLsMominalPw,     StrDisable},
            {CmdSetMaxIop,          StrDisable},
            {CmdSetModel,           StrDisable},
            {CmdSeManuDate,         StrDisable},
            {CmdSetCalDate,         StrDisable},
            {CmdSetPartNumber,      StrDisable},
            {CmdSetCalAPw,          StrEnable},
            {CmdSetCalBPw,          StrDisable},
            {CmdSetCalAPwtoVint,    StrEnable},
            {CmdSetCalBPwtoVint,    StrDisable},
            {CmdSetCalAVtoPw,       StrEnable},
            {CmdSetCalBVtoPw,       StrDisable} };

        string[,] bulkSetdefaultCtrl = new string[6, 2] {
            {CmdTestMode,       StrEnable  },
            {CmdRatedPower,     StrDisable },
            {CmdSetPwMonOut,    StrDisable },
            {CmdSetVgaGain,     StrDisable },
            {CmdSetOffstVolt,   StrDisable },      //Offset 2.500V
            {CmdSetPwCtrlOut,   StrDisable } };    //Internal PCON 2.500V

        string[,] bulkSetFinalSetup = new string[8, 2] {
            {CmdTestMode,           StrEnable},
            {CmdSetCalAPw,          StrEnable},
            {CmdSetCalBPw,          StrDisable},
            {CmdSetCalAPwtoVint,    StrEnable },
            {CmdSetCalBPwtoVint,    StrDisable},
            {CmdSetCalAVtoPw,       StrEnable},
            {CmdSetCalBVtoPw,       StrDisable},
            {CmdRstTime,            StrEnable } };

        string[,] bulkSetTEC = new string[6, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetTECTemp,        StrDisable },
            { CmdSetTECkp,          StrDisable },
            { CmdSetTECki,          StrDisable },
            { CmdSetTECsmpTime,     StrDisable },
            { CmdSetTECena_dis,     StrEnable} };

        string[,] bulkSetTEC532 = new string[2, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetTECTemp,        StrDisable } };
        //=================================================

        string[,] bulkSetVga = new string[6, 2] {
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetInOutPwCtrl,    StrDisable },     //external PCON
            { CmdAnalgInpt,         StrDisable },     //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },     //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable } };    //Inv. TTL line in
        //=================================================

        string[,] analogRead = new string[3, 2] {//read analog inputs
            { CmdRdPwSetPcon,       StrDisable },
            { CmdRdLaserPow,        StrDisable },
            { CmdCurrentRead,       StrDisable } };

        string[,] analogRead2 = new string[7, 2] {//read all analog inputs
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
        string indata_RS232 =   string.Empty;
        string outdata_RS232 =  string.Empty;
        string rString =        string.Empty;
        string cmdTrack =       string.Empty;

        string  rtnHeader =     string.Empty;
        string  rtnCmd =        string.Empty;
        string  rtnValue =      string.Empty;
        string dataBaseName =   string.Empty;

        string filePathLI = string.Empty;
        string filePathRep = string.Empty;

        byte[] byteArrayToTest1 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest2 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest3 = new byte[8];//reads back "bits"

        double[,] dataADC = new double[120, 5];
        double maxPw = 0;
        double maxCurr = 0;

        bool USB_Port_Open =    false;
        bool RS232_Port_Open =  false;
        bool testMode =         false;
            
        int arrayLgth   = 0;
        int arrIndex    = 0;

        //======================================================================
        //======================================================================
        StringBuilder LogString_01 = new StringBuilder();
        //======================================================================
        Thorlabs.PM100D_32.Interop.PM100D pm = null;
        bool pm100ok = false;
        //======================================================================
        SerialPort USB_CDC =        new SerialPort();
        SerialPort RS232 =          new SerialPort();
        //======================================================================
        SqlConnection con = null;
        SqlCommand cmd = null;
        SqlDataReader rdr = null;
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
            //OpenSqlConnection();//done when loading form
            //DisplayData();//not used here
        }
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
            catch (Exception e) { MessageBox.Show("Dtb Open Error " + e.ToString() + "\nInter Laser Parameters Manually"); }
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
        #endregion SQL stuff
        //================================================================================
        private void CDCDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            indata_USB = string.Empty;
            indata_USB = USB_CDC.ReadExisting();  
            this.BeginInvoke(new Action(() => Process_String(indata_USB)));
            //this.BeginInvoke(new EventHandler(delegate { Process_String(indata_USB); }));
            //this.BeginInvoke(new SetTextCallback(SetText), new object[] { InputData });
            //Application.DoEvents();
        }
        //================================================================================
        private void RS232DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            indata_RS232 = string.Empty;
            indata_RS232 = RS232.ReadExisting();  //read data to string 
            this.BeginInvoke(new Action(() => Process_String(indata_RS232)));
        }
        //======================================================================
        private async Task<bool> SendToSerial(string strCmd, string strData, int sendDelay, int intThreshold)
        {
            string mdlNumb      = Tb_SetAdd.Text;
            string stuffToSend  = string.Empty;
            int dly             = 300; //wait delay to fire next instruction send/process/receive time... only set here
            int threshldInt     = 9; //buffer threshold interrupt needs to be review as how to impement... only set here 

            if (sendDelay > 300 ) {dly = sendDelay; }
            else dly = Convert.ToInt16(Tb_RsDelay.Text);//300

            if(intThreshold > 9) { threshldInt = intThreshold; } //long string read back
            USB_CDC.ReceivedBytesThreshold = threshldInt;

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
        private string[]  ChopString(string stringToChop)
        {
            string[] rtnstr = new string[3];
            int strLgh = stringToChop.Length;
            rtnstr[0] = stringToChop.Substring(1, 2);
            rtnstr[1] = stringToChop.Substring(3, 2);
            rtnstr[2] = stringToChop.Substring(5, (strLgh - 7));
            strLgh = 0;
            return rtnstr;
        }
        //======================================================================
        #region Process_String long case where the received string is analysed
        private void  Process_String(string strRcv)
        {
            int strLgh = strRcv.Length;
            string[] returnChop = new string[3];  // 0/Header, 1/cmd, 2/value
            Rt_ReceiveDataUSB.AppendText("<<  " + strRcv); //displays anything....
 
            returnChop[0] = strRcv.Substring(1, 2);
            returnChop[1] = strRcv.Substring(3, 2);
            returnChop[2] = strRcv.Substring(5, (strLgh - 7));
            rtnHeader = returnChop[0];
            rtnCmd =    returnChop[1];
            rtnValue =  returnChop[2];

            if (rtnCmd == cmdTrack)
            {
                cmdTrack = string.Empty;
 
                switch (rtnCmd)
                {
                    case CmdRdUnitNo://99
                        if (RS232_Port_Open == true)
                        {
                            Tb_RSConnect.BackColor = Color.LawnGreen;
                            Tb_USBConnect.BackColor = Color.Red;
                        }
                        else if (USB_Port_Open == true)
                        {
                            Tb_RSConnect.BackColor = Color.Red;
                            Tb_USBConnect.BackColor = Color.LawnGreen;
                        }
                        break;

                    case CmdSetUnitNo://12
                        lbl_RdAdd.Text = rtnHeader;

                        if (lbl_RdAdd.Text == Tb_SetAdd.Text)
                        {
                            Tb_EepromGood.BackColor = Color.LawnGreen;
                            lbl_RdAdd.ForeColor = Color.Green;
                            lbl_RdAdd.Text = rtnHeader;
                        }
                        else { Tb_EepromGood.BackColor = Color.Red; }
                        break;

                    case CmdRdWavelen:
                        Lbl_WaveLg.ForeColor = Color.Green;
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
                        lbl_SerNbReadBack.Text = rtnValue.PadLeft(8, '0');
                        break;

                    case CmdRdFirmware:
                        lbl_SWLevel.ForeColor = Color.Green;
                        lbl_SWLevel.Text = rtnValue.PadLeft(8, '0');
                        break;

                    case CmdRdBplateTemp:
                        Lbl_TempBplt.Text = rtnValue.PadLeft(4, '0');
                        break;

                    case CmdLaserStatus:

                        byteArrayToTest1 = ConvertToByteArr(rtnValue);
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
                        break;

                    case CmdSetPwMonOut:
                        break;

                    case CmdRdCalDate:
                        break;

                    case CmdSetPwCtrlOut:
                        Lbl_RtnPwDACvalue.Text = rtnValue;
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
                        if (rtnValue == StrDisable) { Bt_IntExtPw.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_IntExtPw.BackColor = Color.LawnGreen; }
                        break;

                    case CmdSet0mA:
                        //firmware "error" : returns 00 as cmd
                        break;

                    case CmdSetStramind:
                        break;

                    case CmdRdCmdStautus2://cmd 34 note the array starts from 7 to 0....
                        byteArrayToTest3 = ConvertToByteArr(rtnValue);
                        break;

                    case CmdManufDate:
                        break;

                    case CmdRdPwSetPcon:
                        lbl_ADCpconRd.Text = rtnValue.PadLeft(5, '0');
                        break;

                    case CmdRdInitCurrent:
                        break;

                    case CmdRdModelName:
                        break;

                    case CmdRdLaserPow:
                        lbl_LaserPD.Text = rtnValue.PadLeft(5, '0');
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
                        else if (testMode == false)
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
                        Tb_CalA_PwToADC.Text = rtnValue;              
                        break;

                    case CmdSetCalBVtoPw:
                        Tb_CalB_PwToADC.Text = rtnValue;                      
                        break;

                    case CmdSetCalAPwtoVint:
                        Tb_CalAcmdToPw.Text = rtnValue;
                        break;

                    case CmdSetCalBPwtoVint:
                        Tb_CalBcmdToPw.Text = rtnValue;
                        break;

                    case CmdRstTime:
                        break;

                    case CmdSetSerNumber:
                        Tb_SerNb.Text = rtnValue.PadLeft(8, '0');
                        break;

                    case CmdSetWavelenght:
                        Lbl_WaveLg.ForeColor = Color.Green;
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
                        if (rtnValue == "0000")
                        {
                            testMode = false;
                            Bt_EnableTest.BackColor = Color.SandyBrown;
                        }
                        else if (rtnValue == "0001")
                        {
                            testMode = true;
                            Bt_EnableTest.BackColor = Color.LawnGreen;
                        }
                        break;

                    case CmdSetPSU:
                        break;

                    case CmdSetBaseTempCal:
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
                if (rtnCmd == "00") { /* do nothing */ }
                else {
                    MessageBox.Show("Read Back Missmatch");
                    Tb_RSConnect.BackColor = Color.Red;
                    Tb_USBConnect.BackColor = Color.Red; }
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
                    double RtPw = (Convert.ToDouble(Tb_NomPw.Text)) * 10;
                    dataToAppd = RtPw.ToString("0000");
                    break;

                case CmdCurrentRead:
                    //if(testMode==true) comThresh = 10;
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
                    dataToAppd = Tb_CalA_PwToADC.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalBVtoPw:
                    dataToAppd = Tb_CalB_PwToADC.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalAPwtoVint:
                    dataToAppd = Tb_CalAcmdToPw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetCalBPwtoVint:
                    dataToAppd = Tb_CalBcmdToPw.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRstTime:
                    sndDl = 8000;
                    comThresh = 14;
                    break;

                case CmdSetSerNumber:
                    dataToAppd = Tb_SerNb.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetWavelenght:
                    dataToAppd = Lbl_Wlgth1.Text;
                    break;

                case CmdSetLsMominalPw:
                    double nomPw = (Convert.ToDouble(Tb_SoftNomPw.Text))*10;
                    dataToAppd = nomPw.ToString("0000");
                    break;

                case CmdSetCustomerPm:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSetMaxIop:
                    dataToAppd = Tb_MaxLsCurrent.Text;
                    break;

                case CmdSetCalDate:
                    dataToAppd = dateTimePicker1.Value.Date.ToString("yyyyMMdd");
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdSeManuDate:
                    dataToAppd = dateTimePicker1.Value.Date.ToString("yyyyMMdd");
                    sndDl = 600;
                    comThresh = 14;
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

                case CmdSetTECTemp:
                    dataToAppd = Tb_TECpoint.Text;
                    break;

                case CmdSetTECkp:
                    dataToAppd = Tb_Kp.Text;
                     break;

                case CmdSetTECki:
                    dataToAppd = Tb_Ki.Text;
                     break;

                case CmdSetTECsmpTime:
                    dataToAppd = Tb_LoopT.Text;
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

            bool result = await SendToSerial(cmdToTest, dataToAppd, sndDl, comThresh);

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
                lbl_RdAdd.Text = "00";
                lbl_SerNbReadBack.Text = "00000000";
                lbl_SWLevel.Text = "00000000";
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
                    lbl_RdAdd.Text = "00";
                    lbl_SerNbReadBack.Text = "00000000";
                    lbl_SWLevel.Text = "00000000";
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
            setad = await SendToSerial(CmdRdSerialNo, StrDisable, 300, 9);
            setad = await SendToSerial(CmdRdFirmware, StrDisable, 300, 9);
            return true;
        }
        //======================================================================
        #endregion COM settings
        //======================================================================
        private void Rtb_Display_USB_DoubleClick(object sender, EventArgs e) { Rt_ReceiveDataUSB.Clear(); }
        //======================================================================
        private void Frm_iRIS_Prod_FormClosing(object sender, FormClosingEventArgs e)
        {
            Task<bool> exitAll = ExitPgm();
            Application.Exit();
        }
        //======================================================================
        private async Task<bool> ExitPgm()
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
            if (Bt_PM100.BackColor == Color.LawnGreen) { bool closePM = await PM100Button(); }
            if (Bt_USBinterf.BackColor == Color.LawnGreen) { bool closeUSBint = await SetUsbInterface(); }

            await Task.Delay(100);
            return true;
        }
        //======================================================================
        private void Form_iRIS_Clm_01_Load(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            OpenSqlConnection();
            this.Cursor = Cursors.Default;
        }
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
                        tb_TecSerNumb.BackColor = Color.White;
                        tb_TecSerNumb.Enabled = false;
                    }
                    else { MessageBox.Show("8 Digits Maximum"); }
                }
                else { MessageBox.Show("Provide Details!"); }
            }
        }
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

                    Int16 rdWavelgth = Convert.ToInt16(Tb_Wavelength.Text);

                    if (rdWavelgth <= 300 || rdWavelgth >= 900) { return pm100cnt = false; }//no wavelengh to set PM100
                    else
                    {
                        pm.setWavelength(rdWavelgth);
                        //Set Zero / Dark adjustment
                        //pm.setPowerAutoRange(true);
                        pm100cnt = true;
                    }
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
                     MessageBox.Show(mystring + " " + "");
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
        private async Task<bool> RampDACLI(double startRp, double stopRp, double stepRp, bool rdIntADC, bool invertedRamp)//external PCON
        {
            arrIndex = 0;

            if (invertedRamp == false) {//non inverted ramp 0V-5V

                for (double startRpLp = startRp; startRpLp <= stopRp; startRpLp = startRpLp + stepRp)
                {
                    bool rampDAC1task = await RampExtPcon(startRpLp, true);
                    if (rampDAC1task == false) { break; }
                    else if (rampDAC1task == true) { continue; }
                }
            }
            else if (invertedRamp == true) {//inverted ramp 5V-0V
                for (double startRpLp = startRp; startRpLp >= stopRp; startRpLp = startRpLp - stepRp)
                {
                    bool rampDAC1task = await RampExtPcon(startRpLp, true);
                    if (rampDAC1task == false) { break; }
                    else if (rampDAC1task == true) { continue; }
                }
            }

            return true;
        }
        //======================================================================
        private async Task<bool> RampExtPcon(double extDacValue, bool readAll)
        {
            WriteDAC(extDacValue, 0);//update Pcon DAC

            bool rampDAC1task = await ReadAllanlg(readAll);//displays current in bits

            if (readAll == true)
            {
                dataADC[arrIndex, 0] = Convert.ToDouble(Lbl_PM100rd.Text)*10;
                dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);

                arrIndex++;
            }

            return true;
        }
        //======================================================================
        private async Task<bool> RampDAC1(double startRp, double stopRp, double stepRp, bool rdIntADC)//external PCON  can be replaced....
        {
            bool rampDAC1task = false;
            arrIndex = 0;

            for (double startRpLp = startRp; startRpLp <= stopRp; startRpLp = startRpLp + stepRp) {

                    WriteDAC(startRpLp, 0);
                    rampDAC1task = await ReadAllanlg(rdIntADC);//displays current in bits

                if (rdIntADC == true) {
                    dataADC[arrIndex, 0] = Convert.ToDouble(Lbl_PM100rd.Text)*10; ;
                    dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                    dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                    dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                    dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);

                    arrIndex++; }
            }
            return true;
        }
        //======================================================================
        private async Task<bool> RampDAC1toPower(double toPower, double startRp, double stopRp, double stepRp, bool rdIntADC)//external PCON can be simplified
        {
            bool rampDAC1task = false;
            arrIndex = 0;

            for (double startRpLp = startRp; startRpLp <= stopRp; startRpLp = startRpLp + stepRp)
            {
                WriteDAC(startRpLp, 0);
                rampDAC1task = await ReadAllanlg(rdIntADC);//displays current in bits

                double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW
                if (pm100Res >= toPower) { return true; }//power good

                if (rdIntADC == true)
                {
                    dataADC[arrIndex, 0] = pm100Res * 10;
                    dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                    dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                    dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                    dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);
                    arrIndex++;
                }
            }
            return false;
        }
        //======================================================================
        private async Task<bool> RampDACint(double startRp, double stopRp, double stepRp, bool rdIntADC)//can be simplified
        {
            string intPwVolt = string.Empty;
            bool rmpInt = false;
            arrIndex = 0;

            rmpInt = await SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300, 9);//set internal power

            for (double startRpLp = startRp; startRpLp <= stopRp; startRpLp = startRpLp + stepRp) {//in volts

                intPwVolt = startRpLp.ToString("00.000");
                tb_SetIntPw.Text = intPwVolt;
                rmpInt = await SendToSerial(CmdSetPwCtrlOut, intPwVolt, 300, 9);
                rmpInt = await ReadAllanlg(rdIntADC);

                if (rdIntADC == true) {
                    dataADC[arrIndex, 0] = Convert.ToDouble(Lbl_PM100rd.Text)*10;
                    dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                    dataADC[arrIndex, 2] = Convert.ToDouble(Lbl_RtnPwDACvalue.Text);
                    dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                    dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);
                    arrIndex++; }
            }

            rmpInt = await SendToSerial(CmdSetInOutPwCtrl, StrDisable, 300, 9);//externa power control
            return true;
        }
        //======================================================================
        private async Task<bool> ReadAllanlg(bool fullRd) {//reads all data

            double pwrRead = 0;             //pm100
            double pconRead = ReadADC(0);   //PCON feedback
            double lsrPwRead = ReadADC(1);  //PD Vout
            double lsrCurrRead = ((ReadADC(2)/5.01)*1000); ; //Current Vout converted to mA compatible with laser setup data

            if (pm100ok == true) {//PM100 Connected
                pwrRead = ReadPM100();
                Lbl_PM100rd.Text = pwrRead.ToString("00.000");  } //update label in mW
            else if (pm100ok == false) {
                MessageBox.Show("PM100 not connected");
                return false; }

            Lbl_Vpcon.Text = pconRead.ToString("00.000");
            Lbl_PwreadV.Text = lsrPwRead.ToString("00.000");//*294.12
            Lbl_Viout.Text = lsrCurrRead.ToString("000.0");//already converted to mA
            label4.Text = " Imon mA";
            Lbl_V_I_out.Text = Lbl_Viout.Text;//this can only be refreshed here, tab 2 current lable

            if (lsrCurrRead > maxCurr)
            {
                Set_USB_Digit_Out(0, 0);//Laser disable
                WriteDAC(0, 0);
                MessageBox.Show("Current Error");
                return false;
            }

            if (pwrRead > maxPw) {
                Set_USB_Digit_Out(0, 0);//Laser disable
                WriteDAC(0, 0);
                MessageBox.Show("Power Error");
                return false; }

            if (fullRd == true) { bool readAdc = await LoadGlobalTestArray(analogRead); }//internal uCadc

            await Task.Delay(10);

            return true;
        }
        //======================================================================
        private void Bt_NewTest_Click(object sender, EventArgs e) {
            if (USB_Port_Open == true) { Task<bool> test2 = FirtInit(); }
            else MessageBox.Show("USB not connected");
            //Task<bool> test2 = FirtInit();
        }
        //======================================================================
        private async Task<bool> FirtInit()
        {
            if (bt_NewTest.BackColor == Color.Coral)
            {
                this.Cursor = Cursors.WaitCursor;
                Prg_Bar01.Maximum = 60;

                maxPw = Convert.ToDouble(Tb_maxMaxPw.Text);
                maxCurr = Convert.ToDouble(Tb_MaxLsCurrent.Text);

                Prg_Bar01.Increment(10);
                Lbl_WaveLg.Text = "0000";
                Lbl_WaveLg.ForeColor = Color.DarkBlue;
                Tb_VGASet.Text = "0020";
                tb_SetIntPw.Text = "2.500";
                Tb_SetOffset.Text = "2.500";
                Tb_VPcon.Text = "0.000";
                Lbl_VGAval.ForeColor = Color.DarkBlue;
                Lbl_VGAval.Text = "0000";

                Tb_CalA_PwToADC.Text = "1.0000";
                Tb_CalB_PwToADC.Text = "0.0000";
                Tb_CalA_Pw.Text = "1.0000";
                Tb_CalB_Pw.Text = "0.0000";
                Tb_CalAcmdToPw.Text = "1.0000";
                Tb_CalBcmdToPw.Text = "0.0000";

                Prg_Bar01.Increment(10);
                bool test2 = await CreateRepFile();//generate new .txt file
                Prg_Bar01.Increment(10);
                test2 = await LoadGlobalTestArray(bulkSetLaserIO);
                Prg_Bar01.Increment(10);
                test2 = await LoadGlobalTestArray(bulkSetTEC);
                Prg_Bar01.Increment(10);
                test2 = await LoadGlobalTestArray(bulkSetVarialble);
                Prg_Bar01.Increment(10);

                test2 =await SetUsbInterface();
                //Bt_PM100.Enabled = true;
                test2 = await PM100Button();

                this.Cursor = Cursors.Default;
                bt_NewTest.BackColor = Color.LawnGreen;

                MessageBox.Show(" Button Disconnect USB \n Power Cycle laser \n Button Re-connect USB \n Wait for TEC lock LED \n Start 'Cal VGA' \n");

                Prg_Bar01.Value = 0;
            }
            else if (bt_NewTest.BackColor == Color.LawnGreen) {
                if (Bt_PM100.BackColor == Color.LawnGreen) { bool closePM = await PM100Button(); }
                bt_NewTest.BackColor = Color.Coral;
                MessageBox.Show("Click again to re-initialise test");
            }

            return true;
        }
        //======================================================================
        private void Bt_SetTempParam_Click(object sender, EventArgs e) { Task<bool> test2 = LoadGlobalTestArray(bulkSetTEC); }
        //======================================================================
        private void Bt_SetVGA_Click(object sender, EventArgs e) { Task<bool> setVGA = SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300, 9); }
        //======================================================================
        private void Bt_CalVGA_Click(object sender, EventArgs e) {
            //if green stop ramp....
            if (USB_Port_Open == true) { Task<bool> calvga = CalVGA(); }
            else MessageBox.Show("USB not connected");
        }
        //======================================================================
        public class Ref<T>
        {
            public Ref() { }
            public Ref(T value) { Value = value; }
            public T Value { get; set; }
 
            public override string ToString()
            {
                T value = Value;
                return value == null ? "" : value.ToString();
            }
            public static implicit operator T(Ref<T> r) { return r.Value; }
            public static implicit operator Ref<T>(T value) { return new Ref<T>(value); }
        }
        //Passing parameters by reference to an asynchronous method
        //======================================================================
        private async Task<bool> CalVGA()
        {
            if (Bt_CalVGA.BackColor == Color.Coral)
            {
                const double startRp = 0.000;
                const double stopRp = 5.000;
                const double stepRp = 0.020;//value 0.01..0.05..
                double calPower = 0;
                double setOffSet = 0;
                double setPower = Convert.ToDouble(Tb_minMaxPw.Text);

                string Pw_Pcon_0V = string.Empty;
                string Pw_Pcon_055V = string.Empty;
                string Pw_Pcon_500V = string.Empty;
                string Pw_EnOff = string.Empty;
                string Pw_05vPCON = string.Empty;

                string offset = string.Empty;
                bool initvga = false;
                bool goodOffset = false;

                Prg_Bar01.Maximum = 120;

                this.Cursor = Cursors.WaitCursor;

                initvga = await LoadGlobalTestArray(bulkSetLaserIO);
                Prg_Bar01.Increment(10);
                initvga = await LoadGlobalTestArray(bulkSetVga);
                Prg_Bar01.Increment(10);
                initvga = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300, 9);//default initial VGA gain 20
                Prg_Bar01.Increment(10);
                initvga = await SendToSerial(CmdSetOffstVolt, Tb_SetOffset.Text, 300, 9);//sefault initial offset 2.500V
                Prg_Bar01.Increment(10);
                Set_USB_Digit_Out(0, 0);//enable line
                Set_USB_Digit_Out(1, 0);//
                Tb_VPcon.Text = "00.000";
                WriteDAC(0, 0);

                if (Convert.ToBoolean(Read_USB_Digit_in(2)) == true)
                {//Laser OK //test
                    Tb_LaserOK.BackColor = Color.Green;

                    initvga = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);//Laser Enable
                    Set_USB_Digit_Out(0, 1);//Laser Enable

                    initvga = await ReadAllanlg(true);//test if OK
                    Prg_Bar01.Increment(10);
                    for (int i = 0; i <= 2; i++)//3 VGA set iteration //test
                    {
                        bool boolCalVGA1 = await RampDAC1(startRp, stopRp, stepRp, false);//set VGA MAX power
                        if (boolCalVGA1 == false) break;

                        for (int vgaVal = 20; vgaVal <= 80; vgaVal++)//Ramp and set VGA
                        {

                            if (vgaVal >= 80)
                            {
                                MessageBox.Show("Fault? MAX VGA 80");
                                break;
                            }
                            else
                            {
                                Tb_VGASet.Text = vgaVal.ToString("0000");
                                bool vgaset = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300, 9);
                                vgaset = await ReadAllanlg(false);
                                double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW

                                if (pm100Res >= setPower) break;//continue
                            }
                        }

                        WriteDAC(0.500, 0); //set to 0.5V PCON with above VGA
                        await Task.Delay(200);

                        for (int j = 0; j < 59; j++) //adjust V offset
                        {
                            bool vgaset02 = await ReadAllanlg(false);
                            calPower = Convert.ToDouble(Lbl_PM100rd.Text);//mW @ 0.1%
                            setOffSet = Convert.ToDouble(Tb_SetOffset.Text);//offset value re-initialise for new test

                            if (calPower > (setPower * 0.00130)) { setOffSet = setOffSet + 0.002; } //add offset...reduces power
                            else if (calPower < (setPower * 0.00070)) { setOffSet = setOffSet - 0.002; } //increase power
                            else { goodOffset = true; }

                            offset = setOffSet.ToString("0.000");//format string
                            Tb_SetOffset.Text = offset;
                            bool boolCalVGA5 = await SendToSerial(CmdSetOffstVolt, offset, 400, 9);//update offset

                            if (goodOffset == true) break;
                        }
                     Prg_Bar01.Increment(10);
                    }
                }
                else if (Convert.ToBoolean(Read_USB_Digit_in(2)) == false) {
                Tb_LaserOK.BackColor = Color.Red;
                MessageBox.Show("Laser NOT OK"); }

                WriteDAC(5, 0);
                await Task.Delay(300);
                initvga = await ReadAllanlg(false);
                Pw_Pcon_500V = Lbl_PM100rd.Text;
                Prg_Bar01.Increment(10);

                WriteDAC(0.55, 0);
                await Task.Delay(300);
                initvga = await ReadAllanlg(false);
                Pw_Pcon_055V = Lbl_PM100rd.Text;
                Prg_Bar01.Increment(10);

                WriteDAC(0, 0);
                await Task.Delay(300);
                initvga = await ReadAllanlg(false);
                Pw_Pcon_0V = Lbl_PM100rd.Text;
                Prg_Bar01.Increment(10);

                bool rampdac11 = await RampDAC1toPower(00.050, 0.450, 00.700, 0.005, false);//adjust PCON to MAX power
                initvga = await ReadAllanlg(false);
                Pw_05vPCON = Lbl_Vpcon.Text;
                Prg_Bar01.Increment(10);

                WriteDAC(0, 0);
                Set_USB_Digit_Out(0, 0); //Laser Disable
                await Task.Delay(300);
                initvga = await ReadAllanlg(false);
                Pw_EnOff = Lbl_PM100rd.Text;

                initvga = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9); //end VGA stop test
                
                Lbl_VGAval.Text = Tb_VGASet.Text; //actualise VGA value on TAB2
                Lbl_VGAval.ForeColor = Color.Green;
                Prg_Bar01.Value = 0;

                /*************************************************/
                try
                {
                    if (File.Exists(filePathRep))
                    {
                        using (StreamWriter fs = File.AppendText(filePathRep))
                        {
                            fs.WriteLine("DATE: " + dateTimePicker1.Value.ToString());
                            fs.WriteLine("User: " + Tb_User.Text);
                            fs.WriteLine("Work Order: " + Tb_WorkOrder.Text);
                            fs.WriteLine("Laser Assembly SN.: " + lbl_SerNbReadBack.Text);
                            fs.WriteLine("Laser Board SN.: " + Tb_LsBoardSn.Text);
                            fs.WriteLine("TEC Board SN: " + tb_TecSerNumb.Text);
                            fs.WriteLine("Firmware: " + lbl_SWLevel.Text);
                            fs.WriteLine("Wavelength: " + Lbl_WaveLg.Text);
                            fs.WriteLine("Software Nominal power: " + Tb_SoftNomPw.Text);
                            fs.WriteLine("Set Add.: " + lbl_RdAdd.Text);
                            fs.WriteLine("VGA value: " + Lbl_VGAval.Text);
                            fs.WriteLine("Offset value: " + Tb_SetOffset.Text);
                            fs.WriteLine("Power @ 5V Pcon: " + Pw_Pcon_500V);
                            fs.WriteLine("Power @ 0V Pcon: " + Pw_Pcon_0V);
                            fs.WriteLine("Power @ Enable Off: " + Pw_EnOff);
                            fs.WriteLine("PCON Voltage @ 0.1% power: " + Pw_05vPCON);
                        }
                    }
                }
                catch (Exception err1) { MessageBox.Show(err1.Message); }
                /*************************************************/

                this.Cursor = Cursors.Default;
                Bt_CalVGA.BackColor = Color.LawnGreen;
            }
            else if (Bt_CalVGA.BackColor == Color.LawnGreen)
            {
                Bt_CalVGA.BackColor = Color.Coral;
            }

            return true;
        }
        //======================================================================
        #region Current Zero 
        //======================================================================
        private void Bt_ReaduCcurrent_Click(object sender, EventArgs e) { Task<bool> readuCLsCurrent = SendToSerial(CmdCurrentRead, StrDisable, 300, 9); }
        //======================================================================
        private void Bt_ZeroI_Click(object sender, EventArgs e) { Task<bool> zeroI = ZerroCurrent(); }
        //======================================================================
        private async Task<bool> ZerroCurrent() {

            if (Bt_ZeroI.BackColor == Color.Coral) {
            this.Cursor = Cursors.WaitCursor;

            WriteDAC(0, 0);//Pcon Channel = 0V
            Set_USB_Digit_Out(0, 1);//Laser ON
            bool rdIcal = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);

            rdIcal = await ReadAllanlg(false);
 
            rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300, 9);//read current value from cpu displayed on label

            rdIcal = await SendToSerial(CmdSet0mA, StrDisable , 300, 9);//zero value cal
            rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300, 9);//recheck new cpu value...same voltage offset at Imon OUT
            rdIcal = await ReadAllanlg(false);

                /*************************************************/
                try { if (File.Exists(filePathRep)) { using (StreamWriter fs = File.AppendText(filePathRep)) { fs.WriteLine("Vout converted to mA Mon @ 0V Pcon: " + Lbl_Viout.Text); } } }
                catch (Exception err1) { MessageBox.Show(err1.Message); }
                /*************************************************/

                this.Cursor = Cursors.Default;
                Bt_ZeroI.BackColor = Color.LawnGreen;
                Bt_PwOutMonCal.Enabled = true;
            }

            else if (Bt_ZeroI.BackColor == Color.LawnGreen)
            {
                Bt_ZeroI.BackColor = Color.Coral;
                Bt_PwOutMonCal.BackColor = Color.Coral;
                Bt_PwOutMonCal.Enabled = false;
            }
            return true;
        }
        #endregion 
        //======================================================================
        private void Bt_pdCalibration_Click(object sender, EventArgs e) { Task<bool> pdcal = PD_Calibration(); }
        //======================================================================
        private async Task<bool> PD_Calibration() {

            if (Bt_pdCalibration.BackColor == Color.Coral)
            {
            bool pdCalTask = false;
            const double startRp = 0.600;
            const double stopRp = 4.900;
            const double stepRp = 0.100;
            int arrIndex1 = Convert.ToInt16((stopRp - startRp) / stepRp);
            double[] abResults = new double[2];

            this.Cursor = Cursors.WaitCursor;

            Set_USB_Digit_Out(0, 1);                                        //Enable laser  
            pdCalTask = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9); // 

            pdCalTask = await RampDACLI(startRp, stopRp, stepRp, true, false);//external PCON

            WriteDAC(0, 0);
            Set_USB_Digit_Out(0, 0);                    
            pdCalTask = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
            pdCalTask = await ReadAllanlg(false);

            abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 1, 0);
            Tb_CalA_Pw.Text = abResults[0].ToString("000.0000");
            Tb_CalB_Pw.Text = abResults[1].ToString("000.0000");

            abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 2, 0);
            Tb_CalA_PwToADC.Text = abResults[0].ToString("000.0000");
            Tb_CalB_PwToADC.Text = abResults[1].ToString("000.0000");


            Bt_pdCalibration.BackColor = Color.LawnGreen;
            this.Cursor = Cursors.Default;
            }
            
            else if (Bt_pdCalibration.BackColor == Color.LawnGreen) {
                MessageBox.Show("restart calibration");
                Bt_pdCalibration.BackColor = Color.Coral; ;
            }

            return true;
        }
        //======================================================================
        private void Bt_FinalLsSetup_Click(object sender, EventArgs e) { Task<bool> endSetup = LsFinalSet(); }
        //======================================================================
        private async Task<bool> LsFinalSet() //load data in laser
        {
            if (Bt_FinalLsSetup.BackColor == Color.Coral)
            {
                string chkBxStateExtPwCtrl = StrDisable;
                string chkBxStateEnblSet = StrDisable;
                string chkBxStateDigitModSet = StrDisable;
                string chkBxStateAnlgModSet = StrDisable;
                
                this.Cursor = Cursors.WaitCursor;

                if (ChkBx_ExtPwCtrl.Checked == true) { chkBxStateExtPwCtrl = StrEnable; }
                else { chkBxStateExtPwCtrl = StrDisable; }

                if (ChkBx_EnableSet.Checked == true) { chkBxStateEnblSet = StrEnable; }
                else { chkBxStateEnblSet = StrDisable; }

                if (ChkBx_DigitModSet.Checked == true) { chkBxStateDigitModSet = StrEnable; }
                else { chkBxStateDigitModSet = StrDisable; }

                if (ChkBx_AnlgModSet.Checked == true) { chkBxStateAnlgModSet = StrEnable; }
                else { chkBxStateAnlgModSet = StrDisable; }

                bool finalSet = await SendToSerial(CmdTestMode, chkBxStateExtPwCtrl, 300, 9);
                finalSet = await SendToSerial(CmdEnablLogicvIn, chkBxStateEnblSet, 300, 9);
                finalSet = await SendToSerial(CmdsetTTL, chkBxStateDigitModSet, 300, 9);
                finalSet = await SendToSerial(CmdAnalgInpt, chkBxStateAnlgModSet, 300, 9);

                finalSet = await LoadGlobalTestArray(bulkSetFinalSetup);

                try
                {
                    if (File.Exists(filePathRep))
                    {

                        using (StreamWriter fs = File.AppendText(filePathRep))
                        {
                            fs.WriteLine("A Pw Cal: " + Tb_CalA_Pw.Text);
                            fs.WriteLine("B Pw Cal: " + Tb_CalB_Pw.Text);
                            fs.WriteLine("A Pw to ADC in: " + Tb_CalAcmdToPw.Text);
                            fs.WriteLine("B Pw to ADC in: " + Tb_CalBcmdToPw.Text);
                            fs.WriteLine("A Pcon to Pw: " + Tb_CalA_PwToADC.Text);
                            fs.WriteLine("B Pcon to Pw: " + Tb_CalB_PwToADC.Text);
                            fs.WriteLine("Internal Power Control: " + chkBxStateExtPwCtrl);
                            fs.WriteLine("Analog Modulation set to Inverted: " + chkBxStateAnlgModSet);
                            fs.WriteLine("Digital Modulation set to Inverted: " + chkBxStateDigitModSet);
                            fs.WriteLine("Enable set to Inverted: " + chkBxStateEnblSet);
                        }
                    }
                }
                catch (Exception err1) { MessageBox.Show(err1.Message); }

                this.Cursor = Cursors.Default;
                Bt_FinalLsSetup.BackColor = Color.LawnGreen;
            }

            else if (Bt_FinalLsSetup.BackColor == Color.LawnGreen) { Bt_FinalLsSetup.BackColor = Color.Coral; }

            return true; 
        }
        //======================================================================
        private void Bt_LaserEn_Click_1(object sender, EventArgs e) {
            if (Bt_LaserEn.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdLaserEnable, StrEnable, 300, 9);
                Bt_LaserEn.BackColor = Color.LawnGreen;
            }
            else if (Bt_LaserEn.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                Bt_LaserEn.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_InvDigtMod_Click_1(object sender, EventArgs e) {
            if (Bt_InvDigtMod.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdsetTTL, StrEnable, 300, 9);
                Bt_InvDigtMod.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvDigtMod.BackColor == Color.LawnGreen){
                Task<bool> sdEnd = SendToSerial(CmdsetTTL, StrDisable, 300, 9);
                Bt_InvDigtMod.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_InvAnlg_Click_1(object sender, EventArgs e) {
            if (Bt_InvAnlg.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdAnalgInpt, StrEnable, 300, 9);
                Bt_InvAnlg.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvAnlg.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdAnalgInpt, StrDisable, 300, 9);
                Bt_InvAnlg.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_IntExtPw_Click_1(object sender, EventArgs e) {
            if (Bt_IntExtPw.BackColor == Color.SandyBrown)  {
                Task<bool> sdEne = SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300, 9);
                Bt_IntExtPw.BackColor = Color.LawnGreen;
            }
            else if (Bt_IntExtPw.BackColor == Color.LawnGreen) {
                Task<bool> sdEnd = SendToSerial(CmdSetInOutPwCtrl, StrDisable, 300, 9);
                Bt_IntExtPw.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        private void Bt_SetIntPwDAC_Click(object sender, EventArgs e) { Task<bool> sdEne = SendToSerial(CmdSetPwCtrlOut, tb_SetIntPw.Text, 600, 9); }
        //======================================================================
        private void Bt_EnableTest_Click(object sender, EventArgs e) {
            if (Bt_EnableTest.BackColor == Color.SandyBrown) {
                Bt_EnableTest.BackColor = Color.LawnGreen;
                Task<bool> sdEne = SendToSerial(CmdTestMode, StrEnable, 300, 9);
            }
            else if (Bt_EnableTest.BackColor == Color.LawnGreen) {
                Bt_EnableTest.BackColor = Color.SandyBrown;
                Task<bool> sdEnd = SendToSerial(CmdTestMode,StrDisable,300, 9);
            }
        }
        //======================================================================
        private void Bt_setOffDac_Click(object sender, EventArgs e) { Task<bool> sdEne = SendToSerial(CmdSetOffstVolt, Tb_SetOffset.Text, 600, 9); }
        //======================================================================
        private void Bt_SetPwDac_Click(object sender, EventArgs e) { Task<bool> sdPddac = SendToSerial(CmdSetPwMonOut, Tb_PwToVout.Text, 600, 9); }
        //======================================================================
        private void Bt_InvEnable_Click(object sender, EventArgs e)
        {
            if (Bt_InvEnable.BackColor == Color.SandyBrown)
            {
                Task<bool> sdEne = SendToSerial(CmdEnablLogicvIn, StrEnable, 300, 9);
                Bt_InvEnable.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvEnable.BackColor == Color.LawnGreen)
            {
                Task<bool> sdEnd = SendToSerial(CmdEnablLogicvIn, StrDisable, 30, 90);
                Bt_InvEnable.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        #region Calibrate Power Monitor Output
        private void Bt_PwOutMonCal_Click(object sender, EventArgs e) { Task<bool> runPwCal = PwMonOutCal(); }
        //======================================================================
        private async Task<bool> PwMonOutCal() {

            if (Bt_PwOutMonCal.BackColor == Color.Coral) {

                const double startRp = 00.000;
                const double stopRp = 5.000;
                const double stepRp = 0.020;
                string pmonVmax = Tb_PwToVcal.Text;//4V 
                double pmonVmaxDlb = Convert.ToDouble(Tb_PwToVcal.Text);

                double RatedPw = Convert.ToDouble(Tb_NomPw.Text);//in mW
 
                this.Cursor = Cursors.WaitCursor;

                bool sendCalPw = await SendToSerial(CmdTestMode, StrEnable, 300, 9);

                WriteDAC(00.000, 0);
                sendCalPw = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);
                Set_USB_Digit_Out(0, 1);

                bool rampdac1 = await RampDAC1toPower(RatedPw, startRp, stopRp, stepRp, false);//adjust PCON to MAX power

                sendCalPw = await SendToSerial(CmdSetPwtoVout, pmonVmax, 600, 9);
                sendCalPw = await ReadAllanlg(true);
                
                double pmonRd = Convert.ToDouble(Lbl_PwreadV.Text);
                if (pmonRd > pmonVmaxDlb + 0.2 || pmonRd < pmonVmaxDlb - 0.2) { MessageBox.Show("Pmon Calibration Error"); }

                /*************************************************/
                try { if (File.Exists(filePathRep)) { using (StreamWriter fs = File.AppendText(filePathRep)) { fs.WriteLine("Vout PD Mon @ Max. Pw: " + Lbl_PwreadV.Text); } } }
                catch (Exception err1) { MessageBox.Show(err1.Message); }
                /*************************************************/

                WriteDAC(00.000, 0);
                sendCalPw = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                Set_USB_Digit_Out(0, 0);
                sendCalPw = await ReadAllanlg(true);
                /*************************************************/
                try { if (File.Exists(filePathRep)) { using (StreamWriter fs = File.AppendText(filePathRep)) { fs.WriteLine("Vout PD Mon @ Min. Pw: " + Lbl_PwreadV.Text); } } }
                catch (Exception err1) { MessageBox.Show(err1.Message); }
                /*************************************************/

                this.Cursor = Cursors.Default;
                Bt_PwOutMonCal.BackColor = Color.LawnGreen;
                Bt_PwOutMonCal.Enabled = false;
            }
            
            else if (Bt_PwOutMonCal.BackColor == Color.LawnGreen) {
                MessageBox.Show("End cal.");
                Bt_PwOutMonCal.BackColor = Color.Coral; }

            return true;
        }
        #endregion
        //======================================================================
        private void Bt_LiPlot_Click(object sender, EventArgs e) { Task<bool> liplotseq = LIplot(); }
        //======================================================================
        private async Task<bool> LIplot()
        {

            Set_USB_Digit_Out(0, 0);//enable line
            Set_USB_Digit_Out(1, 0);//digital modulation line
            WriteDAC(0, 0);
    
            if (Bt_LiPlot.BackColor == Color.Coral) {

                this.Cursor = Cursors.WaitCursor;
                Array.Clear(dataADC, 0, dataADC.Length);//clear array with 0 to start with

                bool initvga = false;//async methods
                bool invRamp = false;

                bool iniLItest = await LoadGlobalTestArray(bulkSetVga);

                if (Convert.ToBoolean(Read_USB_Digit_in(2)) == true) //Laser OK 
                {
                    Tb_LaserOK.BackColor = Color.Green; 

                    if (ChkBx_IntDacPcon.Checked == true) //internal PCON DAC 2.500-4.000
                    {
                        double stepRp =     0.020;
                        double startRp =    2.500;
                        double stopRp =     4.000;

                        iniLItest = await SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300, 9);

                        tb_SetIntPw.Text = startRp.ToString();//reset internal DAC

                        Set_USB_Digit_Out(1, 1);//digital modulation line
                        initvga = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);//Laser Enable
                        Set_USB_Digit_Out(0, 1);//Laser Enable
                        initvga = await ReadAllanlg(true);//test if OK

                        MessageBox.Show("Enable Laser");

                        bool boolCalVGA2 = await RampDACint(startRp, stopRp, stepRp, true);

                        initvga = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                        Set_USB_Digit_Out(0, 0); //Laser Disable
                        Set_USB_Digit_Out(1, 0);//digital modulation line

                        tb_SetIntPw.Text = startRp.ToString();//reset internal DAC
                        boolCalVGA2 = await SendToSerial(CmdSetPwCtrlOut, tb_SetIntPw.Text, 300, 9);

                        bool rdAnlg = await ReadAllanlg(false);
                    }

                    else if (ChkBx_IntDacPcon.Checked == false) //external PCON 0-5
                    {
                        double stepRp = 0.050;
                        double startRp = 0.000;
                        double stopRp = 0.000;

                        if (ChkBx_InvExtPcon.Checked == false) //non inverted ramp
                        {
                            invRamp = false;
                            startRp = 0.000;
                            stopRp = 5.000;
                            iniLItest = await SendToSerial(CmdAnalgInpt, StrDisable, 300, 9); //Non Inv. PCON
                        }
                        else if (ChkBx_InvExtPcon.Checked == true) //inverted ramp
                        {
                            invRamp = true;
                            startRp = 5.000;
                            stopRp = 0.000;
                            iniLItest = await SendToSerial(CmdAnalgInpt, StrEnable, 300, 9); //Inv. PCON
                        }

                        iniLItest = await SendToSerial(CmdSetInOutPwCtrl, StrDisable, 300, 9); //External PCON
                        Tb_VPcon.Text = startRp.ToString();//reset external DAC

                        Set_USB_Digit_Out(1, 1);//digital modulation line
                        initvga = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);//Laser Enable
                        Set_USB_Digit_Out(0, 1);//Laser Enable
                        initvga = await ReadAllanlg(true);//test if OK

                        MessageBox.Show("Enable Laser");

                        bool boolCalVGA1 = await RampDACLI(startRp, stopRp, stepRp, true, invRamp);//an other ramp method to be on "safe side"...just not time now to consolidate

                        initvga = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9); //end VGA stop test
                        Set_USB_Digit_Out(0, 0); //Laser Disable
                        Set_USB_Digit_Out(1, 0);//digital modulation line
                        WriteDAC(0, 0);
                        bool rdAnlg = await ReadAllanlg(false);
                    }

                    /**********************************************************************************/

                    int indx1 = dataADC.GetLength(0);//arrays lenght 120
                    int indx = indx1;

                    for (int indx2 = 0; indx2 < indx1; indx2++) {//eliminates trailling 0 as it is cleared with 0 to start with ...?
                        if (dataADC[indx2, 3] == 0) {
                            indx = indx2;
                            break; }
                        else continue; }

                    Rt_ReceiveDataUSB.Clear();

                    for (int arrLp = 0; arrLp < indx; arrLp++) {
                        Rt_ReceiveDataUSB.AppendText(dataADC[arrLp, 3].ToString() + " " +
                                                     dataADC[arrLp, 0].ToString() + " " +
                                                     dataADC[arrLp, 2].ToString() + " " +
                                                     dataADC[arrLp, 1].ToString() + " " +
                                                     dataADC[arrLp, 4].ToString() + " " +
                                                     Footer); }

                    bool liFile = await CreateRepFileLI();

                    try
                    {
                        if (File.Exists(filePathLI))
                        {
                            using (StreamWriter fsLI = File.AppendText(filePathLI))
                            {
                                fsLI.WriteLine("\n");

                                for (int arrLp = 0; arrLp < indx; arrLp++)
                                {
                                    fsLI.WriteLine(dataADC[arrLp, 3].ToString() + " " +
                                                   dataADC[arrLp, 0].ToString() + " " +
                                                   dataADC[arrLp, 2].ToString() + " " +
                                                   dataADC[arrLp, 1].ToString() + " " +
                                                   dataADC[arrLp, 4].ToString());
                                }
                            }
                        }
                    }
                    catch (Exception err1) { MessageBox.Show(err1.Message); }
                    /**********************************************************************************/

                this.Cursor = Cursors.Default;
                Bt_LiPlot.BackColor = Color.LawnGreen;                
                Tb_LaserOK.BackColor = Color.Green;
                }

                else if (Convert.ToBoolean(Read_USB_Digit_in(2)) == false) { //Laser NOT OK
                Tb_LaserOK.BackColor = Color.Red;
                MessageBox.Show("Laser NOT OK");
                return false; }

        }//if green button

        else if (Bt_LiPlot.BackColor == Color.LawnGreen) { Bt_LiPlot.BackColor = Color.Coral; }
 
            return true;
        }
        //======================================================================
        private void Bt_SetIntPwCal_Click(object sender, EventArgs e) { Task<bool> calIntPw = CalIntPwSet(); }
        //======================================================================
        private async Task<bool> CalIntPwSet()
        {
            if (Bt_SetIntPwCal.BackColor == Color.Coral) {

                bool pdCalTask = false;
                const double startRp = 02.760;
                const double stopRp =  03.860;
                const double stepRp = 0.020;
                int arrIndex1 = Convert.ToInt16((stopRp - startRp) / stepRp);
                double[] abResults = new double[2];

                this.Cursor = Cursors.WaitCursor;

                Set_USB_Digit_Out(0, 1);                                            //Enable laser  
                pdCalTask = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);  // 

                pdCalTask = await RampDACint(startRp, stopRp, stepRp, true);

                Set_USB_Digit_Out(0, 0);
                pdCalTask = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                tb_SetIntPw.Text = "02.500";
                pdCalTask = await SendToSerial(CmdSetPwCtrlOut, tb_SetIntPw.Text, 300, 9);
                pdCalTask = await ReadAllanlg(false);

                abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 0, 2);

                Tb_CalAcmdToPw.Text = Convert.ToString(abResults[0]);
                Tb_CalBcmdToPw.Text = Convert.ToString(abResults[1]);

                Bt_SetIntPwCal.BackColor = Color.LawnGreen;
                this.Cursor = Cursors.Default;
            }

            else if (Bt_SetIntPwCal.BackColor == Color.LawnGreen)
            {
                MessageBox.Show("restart calibration");
                Bt_SetIntPwCal.BackColor = Color.Coral; ;
            }

            return true;
        }
        //======================================================================
        #region Temp Comp Base plate
        //======================================================================
        private void Bt_BasepltTemp_Click(object sender, EventArgs e) { Task<bool> readtempBplt = SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9); }
        //======================================================================
        private void Bt_BasePltTempComp_Click(object sender, EventArgs e) { Task<bool> setTcomp = CompBpltTemp(); }
        //======================================================================
        private async Task<bool> CompBpltTemp() {

            if (Bt_BasePltTempComp.BackColor == Color.Coral)
            {
            bool setCompT =     await SendToSerial(CmdSetBaseTempCal, "0000", 300, 9);                      //set init comp to 0000 remember to reset for next init.
            setCompT =          await SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9);                    //read initial value
            
            int measTemp =  ReadExtTemp();                                                                  //get user temp //wait
            int tempComp1 = Convert.ToInt16(Lbl_TempBplt.Text) - measTemp;
            setCompT = await SendToSerial(CmdSetBaseTempCal, tempComp1.ToString("0000"), 300, 9);           //set init comp to 0000 remember to reset for next init.
            setCompT = await SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9);                             //read comp data

            Bt_BasePltTempComp.BackColor = Color.LawnGreen;
            }
            else if (Bt_BasePltTempComp.BackColor == Color.LawnGreen)
            {
                MessageBox.Show("Base plate Cal.");
                Bt_BasePltTempComp.BackColor = Color.Coral;
            }

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
        private double[] FindLinearLeastSquaresFit(double[,] dataXy, int strtX, int endX, int xIndx, int yIndx)
        {
            double[] rtnAb = new double[2];
            double S1 = 0;
            double Sx = 0;
            double Sy = 0;
            double Sxx = 0;
            double Sxy = 0;

                for (int lp2 = strtX; lp2 < endX; lp2++)//changed to, avoid extended and get 0 value...?
                {
                    Sx += dataXy[lp2, xIndx];//x
                    Sy += dataXy[lp2, yIndx];//y

                    Sxx += dataXy[lp2, xIndx] * dataXy[lp2, xIndx];//xx
                    Sxy += dataXy[lp2, xIndx] * dataXy[lp2, yIndx];//xy

                    Rt_ReceiveDataUSB.AppendText(   Convert.ToString(dataXy[lp2, xIndx]) + " " +
                                                    Convert.ToString(dataXy[lp2, yIndx]) + " " +
                                                    Footer);
                    S1++;
                }

                double m = (Sxy * S1 - Sx * Sy) / (Sxx * S1 - Sx * Sx);
                double b = (Sxy * Sx - Sy * Sxx) / (Sx * Sx - S1 * Sxx);

                rtnAb[0] = m;
                rtnAb[1] = b;

            return rtnAb;
        }
        //======================================================================
        private async Task<bool> CreateRepFile()
        {
            string txtName = "Test Results " + Tb_SerNb.Text + ".txt";
            filePathRep = "C:\\Log_01\\" + txtName;
            Tb_txtFilePathRep.Text = filePathRep;

            await Task.Delay(1);

            try
            {
                using (FileStream fs = File.Create(filePathRep)) {
                    Byte[] info = new UTF8Encoding(true).GetBytes(txtName + Footer + Footer);
                    fs.Write(info, 0, info.Length);
                    return true; }
            }
            catch (Exception err) {
                MessageBox.Show(err.Message);
                return false; }
        }
        //======================================================================
        private async Task<bool> CreateRepFileLI()
        {
            string txtName = "LI PLOT " + Tb_SerNb.Text +" " + dateTimePicker1.Value.Date.ToString("ddMMyyyy") + ".txt "; ;
            filePathLI = "C:\\Log_01\\" + txtName;
            Tb_txtFilePathLI.Text = filePathLI;

            await Task.Delay(1);

            try
            {
                using (FileStream fsLI = File.Create(filePathLI))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(txtName + Footer + Footer);
                    fsLI.Write(info, 0, info.Length);
                    return true;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
        }
        //======================================================================
        private void Tb_LaserPN_KeyPress(object sender, KeyPressEventArgs e) { if (e.KeyChar==Convert.ToChar(Keys.Return)) { ReadDbs(); } }
        //======================================================================
        private void Bt_ShipState_Click(object sender, EventArgs e) { Task<bool> sendShpData = SendShpData(); }
        //======================================================================
        private async Task<bool> SendShpData()
        {
            if (Bt_ShipState.BackColor == Color.Coral) {

                string chkBxStateExtPwCtrl = string.Empty;//null
                string chkBxStateEnblSet = string.Empty;
                string chkBxStateDigitModSet = string.Empty;
                string chkBxStateAnlgModSet = string.Empty;

                this.Cursor = Cursors.WaitCursor;

                if (ChkBx_ExtPwCtrl.Checked == true) { chkBxStateExtPwCtrl = StrEnable; }
                else { chkBxStateExtPwCtrl = StrDisable; }

                if (ChkBx_EnableSet.Checked == true) { chkBxStateEnblSet = StrEnable; }
                else { chkBxStateEnblSet = StrDisable; }

                if (ChkBx_DigitModSet.Checked == true) { chkBxStateDigitModSet = StrEnable; }
                else { chkBxStateDigitModSet = StrDisable; }

                if (ChkBx_AnlgModSet.Checked == true) { chkBxStateAnlgModSet = StrEnable; }
                else { chkBxStateAnlgModSet = StrDisable; }

                bool finalSet = await SendToSerial(CmdTestMode, StrEnable, 300, 9);
                finalSet = await SendToSerial(CmdSetInOutPwCtrl, chkBxStateExtPwCtrl, 300, 9);
                finalSet = await SendToSerial(CmdEnablLogicvIn, chkBxStateEnblSet, 300, 9);
                finalSet = await SendToSerial(CmdsetTTL, chkBxStateDigitModSet, 300, 9);
                finalSet = await SendToSerial(CmdAnalgInpt, chkBxStateAnlgModSet, 300, 9);
                finalSet = await SendToSerial(CmdTestMode, StrDisable, 300, 9);

                this.Cursor = Cursors.Default;
                Bt_ShipState.BackColor = Color.LawnGreen;
            }

            else if (Bt_ShipState.BackColor==Color.LawnGreen) { Bt_ShipState.BackColor = Color.Coral; }

            return true;
        }
        //======================================================================
        private void button1_Click(object sender, EventArgs e) { /*ReadDbs();*/ }
        //======================================================================
        private void ReadDbs()
        {
            Rt_ReceiveDataUSB.Clear();
            int dbArrayIdx = 0;
            bool entryOK = false;
            string readstuff = string.Empty;
            string[] laserParameters = new string[30];
            
            try {
                con.Open();
                cmd = new SqlCommand("SELECT * FROM " + "Laser_Setup_Config", con);
                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read()) {
                dbArrayIdx++;
                readstuff = rdr["PartNumber"].ToString();

                    if (readstuff.Contains(Tb_LaserPN.Text)) {
                        entryOK = true;
                        Lbl_MdlName.Text = rdr["Description"].ToString();
                        Lbl_Wlgth1.Text =  rdr["Wavelength"].ToString().PadLeft(4,'0');
                        Tb_Wavelength.Text = Lbl_Wlgth1.Text;

                        Lbl_MonPowerDtbas.Text = rdr["NominalPower"].ToString().PadLeft(5,'0');//note power in mw ! and not 1/10mW
                        Tb_NomPw.Text = Lbl_MonPowerDtbas.Text;
                        Tb_maxMaxPw.Text =  rdr["MaxPower"].ToString().PadLeft(5, '0');
                        Tb_minMaxPw.Text =  rdr["MinPower"].ToString().PadLeft(5, '0');
                        Tb_SoftNomPw.Text = rdr["SoftwareNomPower"].ToString().PadLeft(5, '0');

                        Tb_SetAdd.Text =   rdr["LaserAddress"].ToString().PadLeft(2,'0');

                        double dummyTemp = (Convert.ToDouble(rdr["TEC_BlockTemperature"].ToString())) * 10;
                        Tb_TECpoint.Text = dummyTemp.ToString().PadLeft(4, '0');//note power in mw ! and not 1/10mW

                        Tb_PwToVcal.Text = rdr["PowerMonitorVoltage"].ToString().PadLeft(5, '0');

                        if ((rdr["PowerControlSource"].ToString()) == "Internal  ") { ChkBx_ExtPwCtrl.Checked = true; }
                        else { ChkBx_ExtPwCtrl.Checked = false; }

                        if ((rdr["EnableLine"].ToString()) == "Norm      ") { ChkBx_EnableSet.Checked = false; }
                        else { ChkBx_EnableSet.Checked = true; }//inverted

                        if ((rdr["DigitalModulation"].ToString()) == "Norm      ") { ChkBx_DigitModSet.Checked = false; }
                        else { ChkBx_DigitModSet.Checked = true; }//inverted

                        if ((rdr["AnalogueModulation"].ToString()) == "Norm      ") { ChkBx_AnlgModSet.Checked = false; }
                        else { ChkBx_AnlgModSet.Checked = true; }//inverted

                        MessageBox.Show("PN: " + readstuff + " @ " + dbArrayIdx.ToString());
                        Tb_LaserPN.ForeColor = Color.Green;

                        break;
                    }
                }
            }

            catch (Exception e) { MessageBox.Show("Dtb Read Error " + e.ToString()); }

            if (entryOK == false) { MessageBox.Show("No parts in db\nTry again or contact engineering\n"); }

            if (rdr != null) { rdr.Close(); }
            if (con != null) { con.Close(); }

        }
        //======================================================================
        private void button2_Click(object sender, EventArgs e) {
            double dummyTemp1 = (Convert.ToDouble(Tb_532TempSet.Text)) * 10;
            Tb_TECpoint.Text = dummyTemp1.ToString().PadLeft(4, '0'); //Converted from C to 1/10C
            Task<bool> test2 = LoadGlobalTestArray(bulkSetTEC532);//update TEC temp only
        }
        //======================================================================
        //======================================================================
    }
    //======================================================================
    //======================================================================
}
//======================================================================
//======================================================================