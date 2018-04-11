using System;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO.Ports;
using System.IO;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Thorlabs.PM100D_32.Interop;
using MccDaq;
using System.Management;

//iRIS Production 30058_04
//06/04/2018 ECNxxxxxx

namespace iRIS_CLM_GUI_TEST_04
{
    public partial class Form_iRIS_Clm_test_04 : Form
    {
        #region Constant Commands Definition
        const string rtnNull = "00";
        const string CmdLaserEnable = "02";
        const string CmdRdBplateTemp = "07";
        const string CmdSetUnitNo = "12";
        const string CmdRdMinLsPower = "13";    //minimum laser power 0mW MKT 
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
        const string CmdRdPnNb = "45";
        const string CmdRdCustomerPm = "46";//PCB serial number
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
        const string CmdSetDefEnbl = "85";
        const string CmdReadPSU = "86";
        const string CmdSetBaseTempCal = "87";
        const string CmdRdLaserType = "88"; //CLM = 41 / MKT = 19
        const string CmdSetLaserType =  "89";
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

        const int MKT_Ls = 41;
        const int CLM_Ls = 19;
        const int CCM_Ls = 64;
        #endregion
        //=================================================

        #region Test Sequence Definition
        //=================================================
        string[,] bulkSetLaserIO ;      
        string[,] bulkSetVarialble;
        string[,] bulkSetdefaultCtrl;    
        string[,] bulkSetFinalSetup ;
        string[,] bulkSetTEC ;
        string[,] bulkSetTEC532 ;
        string[,] bulkSetRstClk ;
        //=================================================
        string[,] bulkSetBurnin ;
        //=================================================
        string[,] bulkSetVga ;    //Inv. TTL line in
        //=================================================
        string[,] analogRead ;
        string[,] analogRead2;
        string[,] setLaserType ;
        #endregion
        //=================================================

        #region Variable
        string CmdRdSerialNo    = null;  //cmd 10 MKT
        string CmdRdFirmware    = null;  //cmd 09 MKT
        string CmdRdWavelen     = null;  //cmd 11 MKT
        string CmdSetLsPw       = null;  //cmd 04 MKT
        string CmdRdLaserPow    = null;  //cmd 05 MKT 
        string CmdRatedPower    = null;  //cmd 14 MKT
        string CmdLaserStatus   = null;  //cmd 06 MKT bit 5 
        //=================================================
        string[]  testStringArr = new string[2]; //used to load commands in bulk send
        //=================================================
        string indata_USB       = null;
        string indata_RS232     = null;
        string cmdTrack         = null;

        string  rtnHeader       = null;
        string  rtnCmd          = null;
        string  rtnValue        = null;

        string dataBaseName     = null;
        string filePathRep      = null;

        byte[] byteArrayToTest1 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest2 = new byte[8];//reads back "bits"
        byte[] byteArrayToTest3 = new byte[8];//reads back "bits"

        double[,] dataADC = new double[1200, 5];
        double[] dataSet1 = new double[32];
        double maxPw = 0;
        double maxCurr = 0;

        bool USB_Port_Open =    false;
        bool RS232_Port_Open =  false;
        bool testMode =         false;
        bool boardFound =       false;//moved to global as it is set in one place only and read others
        bool stopLoop =         true; //stops loop test/cal

        int arrayLgth   = 0;
        int arrIndex    = 0;
        int laserType   = 0;
        #endregion
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
        //======================================================================

        #region Setting ADCDAC IO USB Interface
        public MccDaq.DaqDeviceDescriptor[] inventory;
        public MccDaq.MccBoard DaqBoard;
        public MccDaq.ErrorInfo ULStat;
        public MccDaq.Range Range;
        public MccDaq.Range AORange;
        public MccDaq.Range RangeSelected;
        
        Int32 numchannels = 0;
        int   nudAInChannel = 0;
        //======================================================================
        public const String AllowedCharactersInt = "0123456789";
        public const String AllowedCharactersFloat = "0123456789.";
        //======================================================================
        #endregion
        //======================================================================
         public Form_iRIS_Clm_test_04()
        {
            InitializeComponent();
            Getportnames();

            USB_CDC.DataReceived += new SerialDataReceivedEventHandler(CDCDataReceivedHandler);
            RS232.DataReceived   += new SerialDataReceivedEventHandler(RS232DataReceivedHandler);
            tabControl1.TabPages[1].Enabled = false;
            //open sql connection when loading the form
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

            if (rtnCmd == cmdTrack)//what was sent is received in the right order or received...
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

                    case CmdLaserEnable://uses serial send to set
                        if (rtnValue == StrDisable) { Bt_LaserEn.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_LaserEn.BackColor = Color.LawnGreen; }
                        break;

                    case CmdRdBplateTemp://used for MKT and CLM some difference...
                        double rdBkTempDbl = (Convert.ToDouble(rtnValue.PadLeft(4, '0')))/10;
                        Lbl_TempBplt.Text = rdBkTempDbl.ToString("00.0");
                        break;

                    case CmdRdTecTemprt://used for MKT set to 25.0 C
                        //lbl_TecTemp.Text = rtnValue.PadLeft(5,'0');
                        lbl_TecTemp.Text = Tb_532TempSet.Text;
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
                        Lbl_TmrVal.Text = rtnValue;
                        break;

                    case CmdRdSummary:
                        break;

                    case CmdSetInOutPwCtrl:
                        if (rtnValue == StrDisable) { Bt_IntExtPw.BackColor = Color.SandyBrown; }
                        else if (rtnValue == StrEnable) { Bt_IntExtPw.BackColor = Color.LawnGreen; }
                        break;

                    case CmdSet0mA:
                        break;

                    case CmdSetStramind:
                        break;

                    case CmdRdCmdStautus2://cmd 34 note the array starts from 7 to 0....
                        byteArrayToTest3 = ConvertToByteArr(rtnValue);
                        break;

                    case CmdManufDate:
                        break;

                    case CmdRdInitCurrent:
                        break;

                    case CmdRdModelName:
                        break;

                    case CmdRdPnNb:
                        break;

                    case CmdRdCustomerPm:
                        lbl_SerNbReadBack.ForeColor = Color.Green;
                        lbl_SerNbReadBack.Text = rtnValue.PadLeft(16, ' ');
                        break;

                    case CmdCurrentRead:
 
                        if (testMode == true) {
                            Lbl_uClsCurrent.Text = rtnValue.PadLeft(5, '0');
                            Lbl_MaOrBits.Text = "uC Ls. bits"; }
                        else if (testMode == false) {
                            Lbl_uClsCurrent.Text = rtnValue.PadLeft(4, '0');
                            Lbl_MaOrBits.Text = "uC Laser I-mA Read"; }

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

                    case CmdRdPwSetPcon:
                            lbl_ADCpconRd.Text = rtnValue.PadLeft(5, '0');
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

                    //*********** MKT cmd ************//
                    case CmdSetLaserType:
                        break;

                    case CmdRdLaserType://only here the label is updated with a value
                        Int16 rtnLsType = Convert.ToInt16(rtnValue);
                        if (rtnLsType == 41) { Lbl_LsType.Text = "MKT"; }
                        else if (rtnLsType == 19) { Lbl_LsType.Text = "CLM"; }
                        else if (rtnLsType == 64) { Lbl_LsType.Text = "CCM"; }
                        else { Lbl_LsType.Text = "ERR"; }                        
                        break;

                    case CmdRdMinLsPower:
                        break;

                    case CmdSetDefEnbl:
                        break;

                    //********************************//

                    default:

                        if (rtnCmd == CmdRdSerialNo) {
                            lbl_SerNbReadBack.Text = rtnValue.PadLeft(8, '0');
                            lbl_SerNbReadBack.ForeColor = Color.Green;
                        }
                        else if (rtnCmd == CmdRdFirmware) {
                            lbl_SWLevel.Text = rtnValue.PadLeft(8, '0');
                            lbl_SWLevel.ForeColor = Color.Green;
                        }
                        else if (rtnCmd == CmdRdWavelen) {
                            Lbl_WaveLg.Text = rtnValue;
                            Lbl_WaveLg.ForeColor = Color.Green;
                        }
                        else if (rtnCmd == CmdSetLsPw) {
                        }
                        else if (rtnCmd == CmdRdLaserPow) {
                            lbl_LaserPD.Text = rtnValue.PadLeft(5, '0');
                        }
                        else if (rtnCmd==CmdRdMinLsPower)
                        {

                        } 
                        else if (rtnCmd == CmdLaserStatus) {
                        if (rtnValue == "0000") { Tb_MKTLasEnable.BackColor = Color.Red; }
                        else if (rtnValue == "0001") { Tb_MKTLasEnable.BackColor = Color.Green; }
                        }
                        else { MessageBox.Show("Case Default Receive"); }

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
            string cmdToTest = string.Empty;
            string dataToAppd = string.Empty;
            int sndDl = 300;
            int comThresh = 9;

            cmdToTest  = strCmd[0];
            dataToAppd = strCmd[1]; //if anything different will be changed in the case below

            switch (cmdToTest)//cases for CLM and partially MKT (set commands)
            {
                case CmdRdUnitNo:// = "99" this is a label cannot be changed
                    break;

                case CmdSetUnitNo:// = "12"
                    dataToAppd = "00" + Tb_SetAdd.Text;//0002 i.e.
                    break;

                case CmdSetPartNumber:
                    dataToAppd = Tb_LaserPN.Text;
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdLaserEnable://uses serial send to set
                    break;

                case CmdRdBplateTemp:
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
                    sndDl = 600;
                    comThresh = 14;
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

                case CmdRdInitCurrent:
                    break;

                case CmdRdModelName:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdPnNb:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                case CmdRdCustomerPm:
                    sndDl = 600;
                    comThresh = 14;
                    break;

                 case CmdCurrentRead:
                    if(testMode == true) comThresh = 10;
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
                    //dataToAppd = Tb_MaxLsCurrent.Text;//set in private async Task<bool> PwMonOutCal() {
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

                case CmdRdPwSetPcon:
                        sndDl = 300;
                    break;

                case CmdSetBaseTempCal:
                    //compensation value for temperature
                    break;

                case CmdSetTECTemp:
                    double sendTemp = Convert.ToDouble(Tb_TECpoint.Text)*10;//0000 format
                    dataToAppd = sendTemp.ToString("0000");
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

                //*********** MKT cmd ************//
                case CmdSetLaserType:
                    dataToAppd = laserType.ToString();
                    sndDl = 300;
                    break;

                case CmdRdLaserType://send query only
                    sndDl = 300;
                    break;

                case CmdRdMinLsPower://send query only
                    sndDl = 300;
                    break;

                case CmdSetDefEnbl://send enable on default 
                    if (ChkBx_EnableSet.Checked==true) { dataToAppd = "0001"; }
                    else if (ChkBx_EnableSet.Checked == false) { dataToAppd = "0000"; }
                    sndDl = 300;//slow bd rate
                    break;

                //********************************//
                default:

                    if (cmdToTest == CmdRdSerialNo)//updated commands
                    {
                        sndDl = 600;//slow bd rate
                        comThresh = 14;
                    }
                    else if (cmdToTest == CmdRdFirmware)
                    {
                        sndDl = 600;
                        comThresh = 14;
                    }
                    else if (cmdToTest == CmdRdWavelen)
                    {
                        sndDl =300;
                    }
                    else if (cmdToTest == CmdSetLsPw)
                    {
                        sndDl = 300;
                    }
                    else if (cmdToTest == CmdRdLaserPow)
                    {
                        sndDl = 300;
                    }
                    else if (rtnCmd == CmdRdMinLsPower)
                    {

                    }
                    else if (rtnCmd == CmdLaserStatus)
                    {
                        if (rtnValue == "0000") { Tb_MKTLasEnable.BackColor = Color.Red; }
                        else if (rtnValue == "0001") { Tb_MKTLasEnable.BackColor = Color.Green; }
                    }
                    else { MessageBox.Show("Case Default Send"); }
 
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
        private byte[] ConvertToByteArr(string str)
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
        //private void Bt_USB_Click(object sender, EventArgs e)
        private void Bt_USB_Click(object sender, EventArgs e) { Task<bool> setUsbCom = SetComsUSB(); }
        //======================================================================
        private async Task<bool> SetComsUSB()
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
                Bt_SetAddr.Enabled = false;
                Tb_EepromGood.BackColor = Color.Red;
                lbl_RdAdd.Text = "00";
                lbl_SerNbReadBack.Text = "0000000000000000";
                lbl_SWLevel.Text = "00000000";
                Bt_SetLsType.Enabled = false;
            }
            else
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
                }

                try
                {
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

                    bool usbadd1 = await SetAddress();
                    usbadd1 = await CheckLaserType();
                    SetLaserType(laserType);
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
                    lbl_SerNbReadBack.Text = "0000000000000000";
                    lbl_SWLevel.Text = "00000000";
                    Bt_SetLsType.Enabled = false;
                    MessageBox.Show("USB_Port_Open COM Error");
                    return false;
                }
            }
            this.Cursor = Cursors.Default;
            return true;
        }
        //======================================================================
        //private void Bt_RS232_Click(object sender, EventArgs e)
        private void Bt_RS232_Click(object sender, EventArgs e) { SetComUART(); }
        //======================================================================
        private void SetComUART()
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
                lbl_RdAdd.Text = "00";
                lbl_SerNbReadBack.Text = "0000000000000000";
                lbl_SWLevel.Text = "00000000";
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

                    if (laserType == MKT_Ls) { RS232.BaudRate = 19200; }
                    else { RS232.BaudRate = 115200; }

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
                    Task<bool> setaddr = SetAddress();
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
            this.Cursor = Cursors.Default;
        }
        //======================================================================
        private void Bt_Scan_Click(object sender, EventArgs e) { Getportnames(); }
        //======================================================================
        private void Getportnames()
        {
            string[] rtnComstr = new string[10];
            rtnComstr =  GetUSBDevices();
            //foreach (string strCom in rtnComstr) { if (!String.IsNullOrEmpty(strCom)) { Rtb_ComList.AppendText(strCom.ToString() + "\r\n"); } }
            //if (strCom.Contains("FTDI")) MessageBox.Show(strCom.ToString()); }
            //*****************************************************
            string[] portnames = SerialPort.GetPortNames();
            Cb_USB.Items.Clear(); //combo box ComConnect

            foreach (string s in portnames) Cb_USB.Items.Add(s);

            if (Cb_USB.Items.Count > 0) Cb_USB.SelectedIndex = 0;
            else Cb_USB.Text = "No COM port";
            //******************************************************
        }
        //======================================================================
        private void Bt_SetAddr_Click(object sender, EventArgs e) { Task<bool> setadd = SetAddress(); }
        //======================================================================
        private async Task<bool> CheckLaserType()
        {
            bool setad1 = await SendToSerial(CmdTestMode, StrEnable, 300, 9);
            setad1 = await SendToSerial(CmdRdLaserType, StrDisable, 300, 9);
            setad1 = await SendToSerial(CmdTestMode, StrDisable, 300, 9);
            if (Lbl_LsType.Text != Lbl_LaserType.Text) {
                Bt_SetLsType.Enabled = true;
                MessageBox.Show("Set Laser Type");
            }
            return true;
        }
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
            Task<bool> exitAll = ExitPgm();
            Application.Exit();
        }
        //======================================================================
        private async Task<bool> ExitPgm()
        {
            //***********************************************************************
            Properties.Settings.Default.PM100string = Tb_PM100str.Text;
            Properties.Settings.Default.DefaultUser = Tb_User.Text;
            Properties.Settings.Default.RootFolder = Tb_FolderLoc.Text;
            Properties.Settings.Default.WOrder = Tb_WorkOrder.Text;
            Properties.Settings.Default.TableResults = Tb_DatabaseWrt.Text;
            Properties.Settings.Default.TableLoad = Tb_DatabaseString.Text;
            Properties.Settings.Default.Server = Tb_ServerName.Text;
            Properties.Settings.Default.Database01 = Tb_InitialCatalog.Text;
            Properties.Settings.Default.DbUser = Tb_User1.Text;
            Properties.Settings.Default.DbPw = Tb_Pw1.Text;
            Properties.Settings.Default.Save();
            //***********************************************************************
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
            //***********************************************************************
            this.Tb_PM100str.Text = Properties.Settings.Default.PM100string;
            this.Tb_User.Text = Properties.Settings.Default.DefaultUser;
            this.Tb_FolderLoc.Text = Properties.Settings.Default.RootFolder;
            this.Tb_WorkOrder.Text = Properties.Settings.Default.WOrder;
            this.Tb_DatabaseWrt.Text = Properties.Settings.Default.TableResults;
            this.Tb_DatabaseString.Text = Properties.Settings.Default.TableLoad;
            this.Tb_ServerName.Text = Properties.Settings.Default.Server;
            this.Tb_InitialCatalog.Text = Properties.Settings.Default.Database01;
            this.Tb_User1.Text = Properties.Settings.Default.DbUser;
            this.Tb_Pw1.Text = Properties.Settings.Default.DbPw;
            //***********************************************************************
            OpenSqlConnection();
            //************************************************************************
            this.Cursor = Cursors.Default;
        }
        //======================================================================
        private string[] GetUSBDevices()
        {
            string[] comId = new string[10];
            comId.Initialize();
            int inxComId = 0;

            ManagementObjectCollection collection = null;
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            collection = searcher.Get();

            foreach (ManagementObject mo in collection)
            {
                Rtb_ComList.AppendText(mo["Caption"].ToString() + "\n");
                Rtb_ComList.AppendText(mo["Manufacturer"].ToString() + "\n");
                Rtb_ComList.AppendText("\n");
                if(inxComId > 9) { break; }
                else { comId[inxComId] = mo["Caption"].ToString() + "  " + mo["Manufacturer"].ToString()+ "  "; }
                inxComId++;
            }

            collection.Dispose();

            return comId;
        }
        //======================================================================
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
            string pm100InitString = Tb_PM100str.Text;

            try
            {
                pm = new PM100D(pm100InitString, false, true);
                pm.findRsrc(out rsrc);

                if (rsrc != 0)
                {
                    //MessageBox.Show("Thorlab resources " + Convert.ToString(rsrc));
                    //Int16 rdWavelgth = Convert.ToInt16(Tb_Wavelength.Text);
                    if (int.TryParse(Tb_Wavelength.Text, out int rdWavelgth32) == true)
                    {
                        Int16 rdWavelgth = Convert.ToInt16(rdWavelgth32);

                        if (rdWavelgth <= 1800 && rdWavelgth >= 400)//Wavelength OK
                        {
                            pm.setWavelength(rdWavelgth);
                            //Set Zero / Dark adjustment
                            //pm.setPowerAutoRange(true);
                            pm100cnt = true;
                        }
                        else //Wavelength error
                        {
                            return pm100cnt = false;
                        }
                    }
                }
                else
                {
                    pm100cnt = false;
                    MessageBox.Show("No PM100D or Wlgth Error");
                }
            }
            catch (Exception e) { MessageBox.Show("PM100 Error" + e.ToString()); }

            await Task.Delay(2);
            this.Cursor = Cursors.Default;
            return pm100cnt;
        }
        //======================================================================
        private void Bt_RdPM100_Click(object sender, EventArgs e) { ReadPM100();  }
        //======================================================================
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
        //======================================================================
        #endregion
        //======================================================================
        #region // USB Interface Code...
        private void Bt_USBinterf_Click(object sender, EventArgs e) { Task usbInter = SetUsbInterface(); }
        //======================================================================
        private async Task<bool> SetUsbInterface()
        {
            if (Bt_USBinterf.BackColor == Color.Coral)
            {
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
                        catch (MccDaq.ULException ule) { System.Windows.Forms.MessageBox.Show(ule.Message, "No USB-Interface found in system. Run InstaCal"); }
                    }
                }
                else {
                    System.Windows.Forms.MessageBox.Show("No USB-Interface detected");
                    this.Close(); }//close application

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
                     //MessageBox.Show(mystring + " " + "");
                }
                Bt_RsLaserOk.Enabled = true;
                Bt_USBinterf.BackColor = Color.LawnGreen;
                Bt_USBinterf.Text = "USB Interface Connected";
            }

            else if (Bt_USBinterf.BackColor == Color.LawnGreen) {
                MccDaq.DaqDeviceManager.ReleaseDaqDevice(DaqBoard);
                boardFound = false;
                Bt_USBinterf.BackColor = Color.Coral;
                Bt_USBinterf.Text = "USB Interface Connect"; }

            await Task.Delay(1);
            return true;
        }
        //======================================================================
        private void Set_USB_Digit_Out(sbyte portNb, sbyte state)
        {
            if (boardFound == true)
            {
                if (state == 1)
                {
                    if (portNb == 0) { Bt_LsEnable.BackColor = Color.Plum; }
                    else if (portNb == 1) { Bt_DigMod.BackColor = Color.Plum; }
                    ULStat = DaqBoard.DBitOut(DigitalPortType.AuxPort, portNb, DigitalLogicState.High);
                }

                else if (state == 0)
                {
                    if (portNb == 0) { Bt_LsEnable.BackColor = Color.PeachPuff; }
                    else if (portNb == 1) { Bt_DigMod.BackColor = Color.PeachPuff; }
                    ULStat = DaqBoard.DBitOut(DigitalPortType.AuxPort, portNb, DigitalLogicState.Low);
                }
            }
            else { MessageBox.Show("USB interface not connected"); }
        }
        //======================================================================
        private sbyte Read_USB_Digit_in(sbyte portNb)//0 to 7
        {
            if (boardFound==true)
            {
                ULStat = DaqBoard.DBitIn(DigitalPortType.AuxPort, portNb, out DigitalLogicState digiIn);
                return ((sbyte)digiIn);
            }
            else
            {
                MessageBox.Show("USB interface not connected");
                return 0;
            }
        }
        //======================================================================
        private void Bt_LsEnable_Click(object sender, EventArgs e)
        {
            if (Bt_LsEnable.BackColor == Color.PeachPuff) Set_USB_Digit_Out(0, 1);
            else Set_USB_Digit_Out(0, 0);
        }
        //======================================================================
        private void Bt_DigMod_Click(object sender, EventArgs e) // not wired
        {
            if (Bt_DigMod.BackColor==Color.PeachPuff) Set_USB_Digit_Out(1, 1);
            else Set_USB_Digit_Out(1, 0);
        }
        //======================================================================
        private void Bt_RsLaserOk_Click(object sender, EventArgs e) {
            if (boardFound == true) {
                sbyte lsOK = 0;
                lsOK = Read_USB_Digit_in(2);
                if (lsOK == 1) Tb_LaserOK.BackColor = Color.LawnGreen;
                else Tb_LaserOK.BackColor = Color.Red;
            }
            else MessageBox.Show("USB I/O not connected");
        }
        //======================================================================
        private void WriteDAC(double dacValue, int dacChannel)//value in volts
        {
            if (boardFound == true)
            {
                double dacValue1 = 0;
                dacValue1 = dacValue * 0.5;//unipolar convertion
                ULStat = DaqBoard.VOut(dacChannel, RangeSelected, (float)dacValue1, MccDaq.VOutOptions.Default);
                if (dacChannel == 0) { Tb_VPcon.Text = dacValue.ToString("00.000"); }
            }
            else { MessageBox.Show("USB interface not connected"); }
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
            if (boardFound == true)
            { 
            Range = MccDaq.Range.Bip10Volts;//connect ch low to AGND
            ULStat = DaqBoard.VIn32(Convert.ToInt16(adcChannel), Range, out double VInVolts, MccDaq.VInOptions.Default);
            return VInVolts;
            }
            else 
            {
                MessageBox.Show("USB interface not connected");
                return 0;
            }
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
            bool rampState = false;
            arrIndex = 0;

            for (double startRpLp = startRp; (startRpLp <= stopRp) && (stopLoop == false); startRpLp = startRpLp + stepRp)
            {

                    WriteDAC(startRpLp, 0);
                    rampDAC1task = await ReadAllanlg(rdIntADC);//displays current in bits

                if (rampDAC1task == false) { rampState = false; }            

                if (rdIntADC == true) {
                    dataADC[arrIndex, 0] = Convert.ToDouble(Lbl_PM100rd.Text)*10; ;
                    dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                    dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                    dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                    dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);
                    arrIndex++; }
            }
            if (stopLoop == true) { rampState = false; }
            else if (stopLoop == false) { rampState = true; }
            return rampState;
        }
        //======================================================================
        private async Task<bool> RampDAC1toPower(double toPower, double startRp, double stopRp, double stepRp, bool rdIntADC)//external PCON can be simplified
        {
            bool rampDAC1task = false;
            bool rampState = false;
            arrIndex = 0;

            for (double startRpLp = startRp; (startRpLp <= stopRp) && (stopLoop == false); startRpLp = startRpLp + stepRp)
            {
                WriteDAC(startRpLp, 0);
                rampDAC1task = await ReadAllanlg(rdIntADC);//displays current in bits //check for overcurrent
                if (rampDAC1task == false) { rampState = false; } // current error
                else
                {
                    double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW from Analog read
                    if (rdIntADC == true)
                    {
                        dataADC[arrIndex, 0] = pm100Res * 10;
                        dataADC[arrIndex, 1] = Convert.ToDouble(lbl_LaserPD.Text);
                        dataADC[arrIndex, 2] = Convert.ToDouble(lbl_ADCpconRd.Text);
                        dataADC[arrIndex, 3] = Convert.ToDouble(Lbl_Viout.Text);
                        dataADC[arrIndex, 4] = Convert.ToDouble(Lbl_Vpcon.Text);
                        arrIndex++;
                    }

                    if (pm100Res >= toPower) { return true; }
                }
             }
            if (stopLoop == true) { rampState = false; }
            else if (stopLoop == false) { rampState = true; }
            return rampState;
        }
        //======================================================================
        private async Task<bool> RampDACint(double startRp, double stopRp, double stepRp, bool rdIntADC)//can be simplified
        {
            string intPwVolt = string.Empty;
            bool rmpInt = false;
            bool rampState = false;
            arrIndex = 0;

            rmpInt = await SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300, 9);//set internal power

            for (double startRpLp = startRp; (startRpLp <= stopRp) && (stopLoop == false); startRpLp = startRpLp + stepRp) {//in volts

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

            if(stopLoop == true) { rampState = false; }
            else if (stopLoop == false) { rampState = true; }
            return rampState;
        }
        //======================================================================
        private async Task<bool> ReadAllanlg(bool fullRd) {//reads all data

            double pwrRead = 0;    //pm100
            double pconRead = 0;   //PCON feedback
            double lsrPwRead = 0;  //PD Vout

            await Task.Delay(2);
            pconRead =  ReadADC(0);   //PCON feedback
            lsrPwRead = ReadADC(1);  //PD Vout

            double lsrCurrRead = ((ReadADC(2)/5.01)*1000); ; //Current Vout converted to mA compatible with laser setup data
            if (lsrCurrRead > maxCurr)
            {
                Set_USB_Digit_Out(0, 0);//Laser disable
                WriteDAC(0, 0);
                bool maxCurrStop = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                MessageBox.Show("Current Error");
                return false;
            }

            if (pm100ok == true) //PM100 Connected
            {
                pwrRead = ReadPM100();

                if (pwrRead > maxPw)
                {
                    Set_USB_Digit_Out(0, 0);//Laser disable
                    WriteDAC(0, 0);
                    bool maxPwStop = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                    MessageBox.Show("Power Error");
                    return false;
                }
                Lbl_PM100rd.Text = pwrRead.ToString("00.000"); //update label in mW
            } 
            else if (pm100ok == false) {
                MessageBox.Show("PM100 not connected");
                return false; }

            double tempSens = ReadADC(3) * 100;
            Lbl_Vpcon.Text = pconRead.ToString("00.000");
            Lbl_PwreadV.Text = lsrPwRead.ToString("00.000");//*294.12 -- 0 to 4V third tab
            Lbl_Viout.Text = lsrCurrRead.ToString("000.0");//already converted to mA
            Lbl_TempSens.Text = tempSens.ToString("000.00");

            Lbl_PDmonUser.Text = Lbl_PwreadV.Text;//second tab
            Lbl_ImonUser.Text = Lbl_Viout.Text;
            Lbl_PM100user.Text = Lbl_PM100rd.Text;

            if (fullRd == true) { bool readAdc = await LoadGlobalTestArray(analogRead); }//internal uCadc

            return true;
        }
        //======================================================================
        private void Bt_NewTest_Click(object sender, EventArgs e) {
            if (USB_Port_Open == true) { Task<bool> test2 = FirtInit(); }
            else MessageBox.Show("USB not connected");
        }
        //======================================================================
        private bool LoadCurAndPwLimits()
        {
            try {
                maxPw = Convert.ToDouble(Tb_maxMaxPw.Text);//global values...therefore carefully set...
                maxCurr = Convert.ToDouble(Tb_MaxLsCurrent.Text);//to be reviewed at some point...database entry...
            }
            catch (FormatException) {
                MessageBox.Show("Current Format Error");
                return false; }

            if (maxCurr < 10 || maxPw < 1) { return false; }
            else { return true; }
        }
        //======================================================================
        private async Task<bool> FirtInit()
        {
            if (bt_NewTest.BackColor == Color.Coral)
            {
                bool vGood = false;
                Prg_Bar01.Maximum = 60;

                if (vGood = LoadCurAndPwLimits() == true)//Current and Power values OK...
                {
                    this.Cursor = Cursors.WaitCursor;

                    Bt_CalVGA.BackColor = Color.Coral;
                    Bt_ZroCurr_PMonOutCal.BackColor = Color.Coral;
                    Bt_BasePltTempComp.BackColor = Color.Coral;
                    Bt_pdCalibration.BackColor = Color.Coral;
                    Bt_SetIntPwCal.BackColor = Color.Coral;
                    Bt_FinalLsSetup.BackColor = Color.Coral;
                    Bt_RstClk.BackColor = Color.Coral;
                    Bt_LiPlot.BackColor = Color.Coral;
                    Bt_ShipState.BackColor = Color.Aqua;
                    Bt_SetBurnin.BackColor = Color.Coral;

                    Prg_Bar01.Increment(10);
                    Lbl_VGAval.Text = "0000";
                    Lbl_VGAval.ForeColor = Color.DarkBlue;
                    Lbl_WaveLg.Text = "0000";
                    Lbl_WaveLg.ForeColor = Color.DarkBlue;
 
                    Tb_VGASet.Text = "0020";
                    tb_SetIntPw.Text = "2.500";
                    Tb_SetOffset.Text = "2.500";
                    Tb_VPcon.Text = "0.000";

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
                    test2 = await SendToSerial(CmdRdCustomerPm, StrDisable, 300, 9);
                    test2 = await SendToSerial(CmdRdFirmware, StrDisable, 300, 9);
                    test2 = await LoadGlobalTestArray(analogRead2);//added 10-04-2018

                    test2 = await SetUsbInterface();
                    test2 = await PM100Button();

                    this.Cursor = Cursors.Default;

                    test2 = await RESETclk();

                    bt_NewTest.BackColor = Color.LawnGreen;
                    //MessageBox.Show(" Open Laser Shutter \n Click Button Disconnect USB \n Power Cycle laser \n Click Button Re-connect USB \n Wait for TEC lock LED Click 'Rd Laser OK' \n Start 'Cal VGA' \n");
                }

                else MessageBox.Show("value error");
            }
            else if (bt_NewTest.BackColor == Color.LawnGreen) {
                if (Bt_PM100.BackColor == Color.LawnGreen) { bool closePM = await PM100Button(); }
                if (Bt_USBinterf.BackColor == Color.LawnGreen) { bool closeUSBint = await SetUsbInterface(); }
                if (Bt_USBcom.BackColor == Color.LawnGreen) { bool rsetUsb = await SetComsUSB(); }
                bt_NewTest.BackColor = Color.Coral;
                Lbl_LsType.Text = "000";
                MessageBox.Show("Click again to re-initialise test");
                Prg_Bar01.Value = 0;
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
            if ((USB_Port_Open == true)&&(boardFound == true)) { Task<bool> calvga = CalVGA(); }
            else MessageBox.Show("USB / IO USB not connected");
        }
        //======================================================================
        //Passing parameters by reference to an asynchronous method...
        //======================================================================
        private async Task<bool> CalVGA()
        {
            if (Bt_CalVGA.BackColor == Color.Coral)
            {
                Prg_Bar01.Maximum = 120;
                #region set some variable
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
                bool boolCalVGA1 = false;
                bool goodOffset = false;
                stopLoop = false;//can run loop
                int vgaVal = 20;
                #endregion set some variable

                this.Cursor = Cursors.WaitCursor;

                #region init and load test
                initvga = await LoadGlobalTestArray(bulkSetLaserIO);
                Prg_Bar01.Increment(10);
                initvga = await LoadGlobalTestArray(bulkSetVga);
                initvga = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300, 9);//default initial VGA gain 20
                initvga = await SendToSerial(CmdSetOffstVolt, Tb_SetOffset.Text, 300, 9);//sefault initial offset 2.500V
                Prg_Bar01.Increment(10);
                Set_USB_Digit_Out(0, 0);//enable line
                Set_USB_Digit_Out(1, 0);//
                Tb_VPcon.Text = "00.000";
                WriteDAC(0, 0);
                #endregion init and load test

                if ((Convert.ToBoolean(Read_USB_Digit_in(2)) == true) &&  (stopLoop == false)) //Laser OK //test
                {
                    Tb_LaserOK.BackColor = Color.LawnGreen;

                    initvga = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);//Laser Enable
                    Set_USB_Digit_Out(0, 1);//Laser Enable

                    initvga = await ReadAllanlg(true);//test if OK
                    Prg_Bar01.Increment(10);

                    for (int i = 0; (i <= 2)&&(stopLoop == false); i++)//3 VGA set iteration //test
                    {
                        boolCalVGA1 = await RampDAC1(startRp, stopRp, stepRp, false);//set VGA MAX power
                        if (boolCalVGA1 == false)
                        {
                            stopLoop = true;
                            break;//stop test
                        }
                        else
                        {
                            for (vgaVal = 20; (vgaVal <= 80)&&(stopLoop == false); vgaVal++)//Ramp and set VGA
                            {
                                if (vgaVal >= 80)
                                {
                                    stopLoop = true;
                                    MessageBox.Show("Fault? MAX VGA 80");
                                    break;
                                }
                                else
                                {
                                    Tb_VGASet.Text = vgaVal.ToString("0000");
                                    boolCalVGA1 = await SendToSerial(CmdSetVgaGain, Tb_VGASet.Text, 300, 9);
                                    boolCalVGA1 = await ReadAllanlg(false);//if error

                                    if (boolCalVGA1 == false)
                                    {
                                        stopLoop = true;
                                        break;
                                    }
                                    else
                                    {
                                        double pm100Res = Convert.ToDouble(Lbl_PM100rd.Text);//mW
                                        //if (pm100Res >= setPower) break;
                                        if (pm100Res > setPower) break; // added 11/04/2018 the power should be a least above 
                                    }
                                }
                                Prg_Bar01.Increment(10);
                            }

                            WriteDAC(0.500, 0); //set to 0.5V PCON with above VGA
                            await Task.Delay(200);

                            for (int j = 0; (j < 59) && (stopLoop == false); j++) //adjust V offset
                            {
                                boolCalVGA1 = await ReadAllanlg(false);

                                if (boolCalVGA1 == true)
                                {
                                    calPower = Convert.ToDouble(Lbl_PM100rd.Text);//mW @ 0.1%
                                    setOffSet = Convert.ToDouble(Tb_SetOffset.Text);//offset value re-initialise for new test

                                    if (calPower > (setPower * 0.00130)) { setOffSet = setOffSet + 0.002; } //add offset...reduces power
                                    else if (calPower < (setPower * 0.00070)) { setOffSet = setOffSet - 0.002; } //increase power
                                    else { goodOffset = true; }

                                    offset = setOffSet.ToString("0.000");//format string
                                    Tb_SetOffset.Text = offset;
                                    bool boolCalVGA5 = await SendToSerial(CmdSetOffstVolt, offset, 400, 9);//update offset

                                    if (goodOffset == true) break;
                                    Prg_Bar01.Increment(10);
                                }
                            }
                            Prg_Bar01.Value = 0;
                        }// 3 VGA iterations...
                    }
                }

                else if (Convert.ToBoolean(Read_USB_Digit_in(2)) == false) {
                Tb_LaserOK.BackColor = Color.Red;
                MessageBox.Show("Laser NOT OK"); }//Error

               //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                #region set final test
                if (boolCalVGA1 == true || stopLoop == false) // all good
                {
                    Prg_Bar01.Maximum = 120;
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

                    if (vgaVal <= 40)
                    {   Lbl_VGAval.ForeColor = Color.Red;
                        Tb_VGASet.ForeColor = Color.Red;
                        MessageBox.Show("VGA <= 40");
                    }
                    else
                    {
                        Lbl_VGAval.ForeColor = Color.Green;
                        Tb_VGASet.ForeColor = Color.Green;
                    }
                    #endregion set final test
                    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
                    #region set file and record
                    try
                    {
                        if (File.Exists(filePathRep))
                        {
                            using (StreamWriter fs = File.AppendText(filePathRep))
                            {
                                fs.WriteLine("DATE: " + dateTimePicker1.Value.ToString());
                                fs.WriteLine("User: " + Tb_User.Text);
                                fs.WriteLine("Laser Part Number: " + Tb_LaserPN.Text);
                                fs.WriteLine("Laser Assembly Serial Number: " + Tb_SerNb.Text);
                                fs.WriteLine("14284 TEC PCB Serial Number: " + Tb_TecSerNumb.Text);
                                fs.WriteLine("14264 Laser PCB Serial Number: " + lbl_SerNbReadBack.Text);
                                fs.WriteLine("Firmware: " + lbl_SWLevel.Text);
                                fs.WriteLine("EEPROM Wavelength: " + Lbl_WaveLg.Text);
                                fs.WriteLine("Diode Wavelength: " + Tb_Wavelength.Text);
                                fs.WriteLine("Software Nominal power: " + Tb_SoftNomPw.Text);
                                fs.WriteLine("Set Add.: " + lbl_RdAdd.Text);
                                fs.WriteLine("VGA value: " + Lbl_VGAval.Text);
                                fs.WriteLine("Offset value: " + Tb_SetOffset.Text);
                                fs.WriteLine("Power @ 5V Pcon: " + Pw_Pcon_500V);
                                fs.WriteLine("Power @ 0V Pcon: " + Pw_Pcon_0V);
                                fs.WriteLine("Power @ Enable Off: " + Pw_EnOff);
                                fs.WriteLine("PCON Voltage @ 0.1% power: " + Pw_05vPCON);
                            }
                            dataSet1[0] = Convert.ToDouble(Pw_Pcon_500V) ;
                            dataSet1[1] = Convert.ToDouble(Pw_Pcon_0V);
                            dataSet1[2] = Convert.ToDouble(Pw_EnOff);
                            dataSet1[3] = Convert.ToDouble(Pw_05vPCON);
                        }
                    }
                    catch (Exception err1) { MessageBox.Show(err1.Message); }
                    #endregion set file and record
                    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//

                        if (stopLoop == true) { bool endLoop = await ResetAll(); }
                        else {
                        this.Cursor = Cursors.Default;
                        Bt_CalVGA.BackColor = Color.LawnGreen; }
                }
                else
                {
                    Prg_Bar01.Value = 0;
                    stopLoop = true;
                    initvga = await ResetAll();
                    this.Cursor = Cursors.Default;
                }
            }

            else if (Bt_CalVGA.BackColor == Color.LawnGreen)
            {
                Bt_CalVGA.BackColor = Color.Coral;
                Prg_Bar01.Value = 0;
            }
            return true;
        }
        //======================================================================
        private async Task<bool> ResetAll()
        {
            WriteDAC(0, 0);
            Set_USB_Digit_Out(0, 0); //Laser Disable
            bool sendStop = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
            sendStop = await ReadAllanlg(false);
            Prg_Bar01.Value = 0;
            Bt_CalVGA.BackColor = Color.Coral;
            Tb_LaserOK.BackColor = Color.Red;
            Lbl_VGAval.ForeColor = Color.Red;
            Tb_VGASet.ForeColor = Color.Red;
            this.Cursor = Cursors.Default;
            MessageBox.Show("VGA Cal Abort");
            return true;
        }
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
            pdCalTask = await ReadAllanlg(false);//refresh/reset label
            pdCalTask = await LoadGlobalTestArray(analogRead);//refresh/reset labels 

            abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 1, 0);
            Tb_CalA_Pw.Text = abResults[0].ToString("0000.0000");
            Tb_CalB_Pw.Text = abResults[1].ToString("0000.0000");

            abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 2, 0);
            Tb_CalA_PwToADC.Text = abResults[0].ToString("0000.0000");
            Tb_CalB_PwToADC.Text = abResults[1].ToString("0000.0000");


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

                finalSet = await WriteResToDb(0);

                this.Cursor = Cursors.Default;
                Bt_FinalLsSetup.BackColor = Color.LawnGreen;
            }

            else if (Bt_FinalLsSetup.BackColor == Color.LawnGreen) { Bt_FinalLsSetup.BackColor = Color.Coral; }

            return true; 
        }
        //======================================================================
        private async Task<bool> WriteResToDb(sbyte saveData)
        {
            //string cmdString = "INSERT INTO " + Tb_DatabaseWrt.Text + " (PartNumber, LaserAddress) VALUES (@val1, @val2)";

            string cmdString0 = "INSERT INTO " + Tb_DatabaseWrt.Text +
                " ( PartNumber, LaserAddress, Laser_Assy_Sn, Laser_Board_Sn, TEC_Board_Sn, SwLevel, TestDate, Wavelength, SoftwareNomPower, TecName, VGA_Value, V_Offset, "+
                " Pw_at_5VPCON, Pw_at_0VPCON, Pw_at_Enbl_Off, VPCON_at_01pc, I_OUT_at_0VPCON, I_OUT_at_Nom_Pw, V_OUT_PD_at_Nom_Pw, V_OUT_PD_at_Min_Pw, "+
                " TEC_BlockTemperature, A_Pw_Cal ,B_Pw_Cal, A_Pw_to_ADC_Cal, B_Pw_to_ADC_Cal, A_PCON_to_Pw_Cal, B_PCON_to_Pw_Cal, Diode_Wavelength )" +
                " VALUES ( @val1, @val2, @val3, @val4, @val5, @val6, @val7, @val8, @val9, @val10, @val11, @val12, " +
                " @val13, @val14, @val15, @val16, @val17, @val18, @val19, @val20, @val21 , @val22, @val23, @val24, @val25, @val26, @val27, @val32)";

            //command.CommandText = "UPDATE Student SET Address = @add, City = @cit Where FirstName = @fn and LastName = @add";

            string cmdString1 = "UPDATE "+ Tb_DatabaseWrt.Text +
                " SET PowerControlSource = @val28, EnableLine = @val29, DigitalModulation = @val30, AnalogModulation = @val31 "+
                " WHERE  Laser_Assy_Sn = "+ Tb_SerNb.Text; //assuming there is only one laser serial number equivalent (double tested?)

            try
            {
                con.Open();
                cmd.Parameters.Clear();

                if (saveData == 0) {
                    cmd = new SqlCommand(cmdString0, con);
                    //************************************************************************//
                    cmd.Parameters.AddWithValue("@val1", Tb_LaserPN.Text);
                    cmd.Parameters.AddWithValue("@val2", Tb_SetAdd.Text);
                    cmd.Parameters.AddWithValue("@val3", Tb_SerNb.Text);
                    cmd.Parameters.AddWithValue("@val4", lbl_SerNbReadBack.Text);
                    cmd.Parameters.AddWithValue("@val5", Tb_TecSerNumb.Text);
                    cmd.Parameters.AddWithValue("@val6", lbl_SWLevel.Text);
                    cmd.Parameters.AddWithValue("@val7", dateTimePicker1.Text);
                    cmd.Parameters.AddWithValue("@val8", Lbl_WaveLg.Text);
                    cmd.Parameters.AddWithValue("@val9", Tb_SoftNomPw.Text);
                    cmd.Parameters.AddWithValue("@val10", Tb_User.Text);
                    cmd.Parameters.AddWithValue("@val11", Lbl_VGAval.Text);
                    cmd.Parameters.AddWithValue("@val12", Tb_SetOffset.Text);

                    cmd.Parameters.AddWithValue("@val13", dataSet1[0]);
                    cmd.Parameters.AddWithValue("@val14", dataSet1[1]);
                    cmd.Parameters.AddWithValue("@val15", dataSet1[2]);
                    cmd.Parameters.AddWithValue("@val16", dataSet1[3]);
                    cmd.Parameters.AddWithValue("@val17", dataSet1[4]);
                    cmd.Parameters.AddWithValue("@val18", dataSet1[5]);
                    cmd.Parameters.AddWithValue("@val19", dataSet1[6]);
                    cmd.Parameters.AddWithValue("@val20", dataSet1[7]);

                    cmd.Parameters.AddWithValue("@val21", Tb_532TempSet.Text);
                    cmd.Parameters.AddWithValue("@val22", Tb_CalA_Pw.Text);
                    cmd.Parameters.AddWithValue("@val23", Tb_CalB_Pw.Text);

                    cmd.Parameters.AddWithValue("@val24", Tb_CalA_PwToADC.Text);
                    cmd.Parameters.AddWithValue("@val25", Tb_CalB_PwToADC.Text);
                    cmd.Parameters.AddWithValue("@val26", Tb_CalAcmdToPw.Text);
                    cmd.Parameters.AddWithValue("@val27", Tb_CalBcmdToPw.Text);

                    cmd.Parameters.AddWithValue("@val32", Tb_Wavelength.Text);

                }

                if (saveData == 1) {
                    cmd = new SqlCommand(cmdString1, con);
                    //************************************************************************//
                    if (ChkBx_ExtPwCtrl.Checked == true) { cmd.Parameters.AddWithValue("@val28", "Internal"); }
                    else { cmd.Parameters.AddWithValue("@val28", "External"); }

                    if (ChkBx_EnableSet.Checked == true) { cmd.Parameters.AddWithValue("@val29", "Inverted"); }
                    else { cmd.Parameters.AddWithValue("@val29", "Norm"); }

                    if (ChkBx_DigitModSet.Checked == true) { cmd.Parameters.AddWithValue("@val30", "Inverted"); }
                    else { cmd.Parameters.AddWithValue("@val30", "Norm"); }

                    if (ChkBx_AnlgModSet.Checked == true) { cmd.Parameters.AddWithValue("@val31", "Inverted"); }
                    else { cmd.Parameters.AddWithValue("@val31", "Norm"); }
                }
                 //************************************************************************//
                cmd.ExecuteNonQuery();
            }

            catch (Exception) { MessageBox.Show("Write Table Error"); }
            finally { con.Close(); }
            await Task.Delay(2);

            return true;
        }
        //======================================================================
        private void Bt_LaserEn_Click_1(object sender, EventArgs e) { Task<bool> sdEnd = SoftLaserEnable(); }
        //======================================================================
        private async Task<bool> SoftLaserEnable()
        {
            bool sdEne = false;

            if (Bt_LaserEn.BackColor == Color.SandyBrown)
            {
                sdEne = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);
                Bt_LaserEn.BackColor = Color.LawnGreen;
            }
            else if (Bt_LaserEn.BackColor == Color.LawnGreen)
            {
                sdEne = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                Bt_LaserEn.BackColor = Color.SandyBrown;
            }

        sdEne = await SendToSerial(CmdLaserStatus, StrDisable, 300, 9);
        return sdEne;
        }
        //======================================================================
        private void Bt_InvDigtMod_Click_1(object sender, EventArgs e) {
            if (Bt_InvDigtMod.BackColor == Color.SandyBrown) {
                Task<bool> sdEne = SendToSerial(CmdsetTTL, StrEnable, 300, 9); //invert set to 1 xor 0 = 1
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
                Task<bool> sdEne = SendToSerial(CmdEnablLogicvIn, StrEnable, 300, 9); // invert enabled 
                Bt_InvEnable.BackColor = Color.LawnGreen;
            }
            else if (Bt_InvEnable.BackColor == Color.LawnGreen)
            {
                Task<bool> sdEnd = SendToSerial(CmdEnablLogicvIn, StrDisable, 30, 9); // invert disable 0 xor 1/0
                Bt_InvEnable.BackColor = Color.SandyBrown;
            }
        }
        //======================================================================
        //======================================================================
        #region Calibrate Power Monitor Output
        private void Bt_ZroCurr_PMonOutCal_Click(object sender, EventArgs e) { Task<bool> zeroIPcal = ZeroMaxIPcal(); }
        //======================================================================
        private async Task<bool> ZerroCurrent()
        {
            
            WriteDAC(0, 0);//Pcon Channel = 0V
            Set_USB_Digit_Out(0, 1);//Laser ON
            bool rdIcal = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);

            rdIcal = await ReadAllanlg(false);
            rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300, 9);//read current value from cpu displayed on label
            rdIcal = await SendToSerial(CmdSet0mA, StrDisable, 300, 9);//zero value cal
            rdIcal = await SendToSerial(CmdCurrentRead, StrDisable, 300, 9);//recheck new cpu value...same voltage offset at Imon OUT
            rdIcal = await ReadAllanlg(false);

            /*************************************************/
            try { if (File.Exists(filePathRep)) {
                using (StreamWriter fs = File.AppendText(filePathRep)) { fs.WriteLine("V_Iout converted to mA Mon @ 0V Pcon: " + Lbl_Viout.Text); } }
                dataSet1[4] = Convert.ToDouble(Lbl_Viout.Text);
            }
            catch (Exception err1) { MessageBox.Show(err1.Message); }
            /*************************************************/
            return true;
        }
        //======================================================================
        private async Task<bool> ZeroMaxIPcal()
        {
            if (Bt_ZroCurr_PMonOutCal.BackColor == Color.Coral)
            {
                this.Cursor = Cursors.WaitCursor;
                
                bool sendzrcal = await ZerroCurrent();
                Prg_Bar01.Increment(10);
                sendzrcal = await PwMonOutCal();
                Prg_Bar01.Increment(10);

                Bt_ZroCurr_PMonOutCal.BackColor = Color.LawnGreen;
                Prg_Bar01.Value = 0;
                this.Cursor = Cursors.Default;
            }
            else if (Bt_ZroCurr_PMonOutCal.BackColor == Color.LawnGreen)
            {
                Bt_ZroCurr_PMonOutCal.BackColor = Color.Coral;
                MessageBox.Show("End cal.");
            }
                return true;
        }
        //======================================================================
        private async Task<bool> PwMonOutCal()
        {
                const double startRp = 00.000;
                const double stopRp = 5.000;
                const double stepRp = 0.020;
                string iopNomPw  = string.Empty;
                string voutPDmax = string.Empty;
                string voutPDmin = string.Empty;
                string pmonVmax = Tb_PwToVcal.Text;//4V 
                double pmonVmaxDlb = Convert.ToDouble(Tb_PwToVcal.Text);
                double RatedPw = Convert.ToDouble(Tb_NomPw.Text);//in mW i.e. 50mW

                this.Cursor = Cursors.WaitCursor;

                bool sendCalPw = await SendToSerial(CmdTestMode, StrEnable, 300, 9);

                WriteDAC(00.000, 0);
                sendCalPw = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);
                Set_USB_Digit_Out(0, 1);

                bool rampdac1 = await RampDAC1toPower(RatedPw, startRp, stopRp, stepRp, false);//adjust PCON to NomPw

                if (rampdac1 == false) { return false; }
                else
                {
                    sendCalPw = await ReadAllanlg(true);
                    iopNomPw =  Lbl_Viout.Text; //current 
                    Prg_Bar01.Increment(10);

                    sendCalPw = await SendToSerial(CmdSetPwtoVout, pmonVmax, 600, 9);//send 4V
                    sendCalPw = await SendToSerial(CmdSetMaxIop, iopNomPw, 600, 9);//save the nominal current for nominal power

                    sendCalPw = await ReadAllanlg(true);
                    voutPDmax = Lbl_PwreadV.Text; //nominal power recorded
                    double pmonRd = Convert.ToDouble(voutPDmax);//check calibration for Pmon
                    if (pmonRd > pmonVmaxDlb + 0.1 || pmonRd < pmonVmaxDlb - 0.1) { MessageBox.Show("Pmon Calibration Error"); }

                    WriteDAC(00.000, 0);//reset ramp
                    Set_USB_Digit_Out(0, 0);
                    sendCalPw = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
                    sendCalPw = await ReadAllanlg(true);
                    /*************************************************/
                    try { if (File.Exists(filePathRep)) { using (StreamWriter fs = File.AppendText(filePathRep)) {
                                fs.WriteLine("V_Iout mon. converted to mA @ Nominal Power: " + iopNomPw);
                                fs.WriteLine("Vout PD Mon @ Nominal Pw: " + voutPDmax);
                                fs.WriteLine("Vout PD Mon @ Min. Pw: " + Lbl_PwreadV.Text); }
                        dataSet1[5] = Convert.ToDouble(iopNomPw);
                        dataSet1[6] = Convert.ToDouble(voutPDmax);
                        dataSet1[7] = Convert.ToDouble(Lbl_PwreadV.Text);
                    }
                }//Pmon for minimal power
                    catch (Exception err1) { MessageBox.Show(err1.Message); }
                    /*************************************************/
                }
            Prg_Bar01.Increment(10);
            return true;
        }
        #endregion
        //======================================================================
        //======================================================================
        private void Bt_LiPlot_Click(object sender, EventArgs e) {  Task<bool> liplotseq = MaxLsPowerSet(); }//ramp
        //======================================================================
        private async Task<bool> MaxLsPowerSet()
        {
            Set_USB_Digit_Out(0, 0);//enable line off
            Set_USB_Digit_Out(1, 0);//digital modulation line
            WriteDAC(0, 0);
            bool setPwCheck = false; //async methods
            double dacPCONSet = 0;

            if (Bt_LiPlot.BackColor == Color.Coral)//ready for test
            {
                this.Cursor = Cursors.WaitCursor;
                bool iniLItest = await LoadGlobalTestArray(bulkSetVga); //initialise IO and test mode

            if (Convert.ToBoolean(Read_USB_Digit_in(2)) == true) //Laser OK 
                {
                    Tb_LaserOK.BackColor = Color.LawnGreen; 

                    if (ChkBx_ExtPwCtrl.Checked == true) //internal PCON - Calibrated now, then the power setting in 1/10mW should be OK
                    {
                        setPwCheck = await SendToSerial(CmdSetInOutPwCtrl, StrEnable, 300, 9); //internal Pw Ctrl in 1/10mW
                        setPwCheck = await SendToSerial(CmdTestMode, StrDisable, 300, 9); //run mode

                        double finalPowerDbl = Convert.ToDouble(Tb_SoftNomPw.Text) * 10;  //nominal power from database
                        //double finalPowerDbl = Convert.ToDouble(Tb_SoftNomPw.Text) * 9;  //nominal power from database

                        string finalPower = finalPowerDbl.ToString("0000");
                        Tb_SetPower03.Text = finalPower;
                        setPwCheck = await SendToSerial(CmdSetLsPw, finalPower, 300, 9); //set power //ready to test
                    }
                     else if (ChkBx_ExtPwCtrl.Checked == false) //external PCON 0V or 5V -- I/O set in bulkSetVga sequence 
                    {
                        setPwCheck = await SendToSerial(CmdSetInOutPwCtrl, StrDisable, 300, 9); //External Pw Ctrl in 1/10mW
                        setPwCheck = await SendToSerial(CmdTestMode, StrEnable, 300, 9); //test mode
                        if (ChkBx_AnlgModSet.Checked == false) //non inverted ramp
                        {
                            dacPCONSet = 5.000;
                            //dacPCONSet = 4.850;

                            iniLItest = await SendToSerial(CmdAnalgInpt, StrDisable, 300, 9); //Non Inv. PCON
                        }
                        else if (ChkBx_AnlgModSet.Checked == true) //inverted ramp
                        {
                            dacPCONSet = 0.000;
                            //dacPCONSet = 0.150;

                            iniLItest = await SendToSerial(CmdAnalgInpt, StrEnable, 300, 9); //Inv. PCON
                        }
                        Tb_VPcon.Text = dacPCONSet.ToString();//set and update external DAC start value
                        setPwCheck = await SendToSerial(CmdTestMode, StrDisable, 300, 9); //run mode mode
                    }

                    Set_USB_Digit_Out(1, 1);//digital modulation line
                    Set_USB_Digit_Out(0, 1);//Laser Enable
                    WriteDAC(dacPCONSet, 0);//stays on no effect if internal PCON
                    setPwCheck = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);//Laser Enable
                    setPwCheck = await SendToSerial(CmdOperatingHr, StrDisable, 600, 9); //read timer
                    setPwCheck = await ReadAllanlg(true);//test if OK 

                    this.Cursor = Cursors.Default;
                    Bt_LiPlot.BackColor = Color.LawnGreen;
                    Bt_LiPlot.Text = "Click To Reset Laser Power";    
                }
                else if (Convert.ToBoolean(Read_USB_Digit_in(2)) == false) //Laser NOT OK
                {
                setPwCheck = await ResetLaserPCON();//reset laser IO
                Tb_LaserOK.BackColor = Color.Red;
                MessageBox.Show("Laser NOT OK");
                return false;
                }
        }//if green button

        else if (Bt_LiPlot.BackColor == Color.LawnGreen)
            {
                this.Cursor = Cursors.WaitCursor;

                Set_USB_Digit_Out(1, 0);//digital modulation line
                Set_USB_Digit_Out(0, 0);//Laser Enable
                WriteDAC(0, 0);//stays on
                Tb_VPcon.Text = "00.000";
                Tb_SetPower03.Text = "0000";
                setPwCheck = await SendToSerial(CmdSetLsPw, "0000", 300, 9); //set power

                setPwCheck = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);//Laser Enable
                setPwCheck = await SendToSerial(CmdTestMode, StrEnable, 300, 9); //Test mode
                setPwCheck = await ResetLaserPCON();//reset laser IO and read analog
                setPwCheck = await SendToSerial(CmdOperatingHr, StrDisable, 600, 9); //read timer

                Bt_LiPlot.BackColor = Color.Coral;
                Bt_LiPlot.Text = "Set to Nominal Laser Power";
                this.Cursor = Cursors.Default;
            }
            setPwCheck = await SendToSerial(CmdLaserStatus, StrDisable, 300, 9);
            return true;
        }
        //======================================================================
        private async Task<bool> ResetLaserPCON()
        {
            Set_USB_Digit_Out(0, 0); //Laser Disable
            Set_USB_Digit_Out(1, 0);//digital modulation line
            WriteDAC(0, 0);
            bool rstLas = await SendToSerial(CmdLaserEnable, StrDisable, 300, 9);
            //tb_SetIntPw.Text = startRp.ToString();//reset internal DAC
            tb_SetIntPw.Text = "2.500";//reset internal DAC
            rstLas = await SendToSerial(CmdSetPwCtrlOut, tb_SetIntPw.Text, 300, 9);
            rstLas = await ReadAllanlg(true);
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
                pdCalTask = await LoadGlobalTestArray(analogRead);//refresh/reset labels 

                abResults = FindLinearLeastSquaresFit(dataADC, 0, arrIndex1, 0, 2);

                Tb_CalAcmdToPw.Text = (abResults[0]).ToString("00000.0000");
                Tb_CalBcmdToPw.Text = (abResults[1]).ToString("00000.0000");

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
        private void Bt_BasepltTemp_Click(object sender, EventArgs e) { Task<bool> readtempBplt = ReadBpTemp(); }
        //======================================================================
        private async Task<bool> ReadBpTemp()
        {
            bool readtempBplt1 = false;
            if (laserType == MKT_Ls ) {
                readtempBplt1 = await SendToSerial(CmdTestMode, StrDisable, 300, 9);
                readtempBplt1 = await SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9);
                readtempBplt1 = await SendToSerial(CmdTestMode, StrEnable, 300, 9);
            }
            else { readtempBplt1 = await SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9); }
            return true;
        } 
        //======================================================================
        private void Bt_BasePltTempComp_Click(object sender, EventArgs e) { Task<bool> setTcomp = CompBpltTemp(); }
        //======================================================================
        private async Task<bool> CompBpltTemp() {

            if (Bt_BasePltTempComp.BackColor == Color.Coral)
            {
            bool setCompT =     await SendToSerial(CmdSetBaseTempCal, "0000", 300, 9);                      //set init comp to 0000 remember to reset for next init.
            setCompT =          await SendToSerial(CmdRdBplateTemp, StrDisable, 300, 9);                    //read initial value
            
            //double measTemp =  ReadExtTemp();                                                               //get user temp //wait
            double measTemp = ReadExtTempLM35();//if LM35 implemented
            double tempComp1 = (Convert.ToDouble(Lbl_TempBplt.Text) - measTemp)*10;
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
        private double ReadExtTemp()//user thermometer reading
        {
            double strpopupInt = 0;
            Getitright:
            string strpopup = Microsoft.VisualBasic.Interaction.InputBox(" Enter Temperature in C \n", "Base Plate Temperature Compensation", "00.0");

            bool t = Information.IsNumeric(strpopup);
            int strLgh = strpopup.Length;

            if (t == true) {
                if (strLgh < 5) { strpopupInt = Convert.ToDouble(strpopup); }
                else {
                    MessageBox.Show("Format 00.0");
                    goto Getitright; }
            }
            else {
                MessageBox.Show("Numerical only");
                goto Getitright;
            }

            return strpopupInt;
        }
        //======================================================================
        private double ReadExtTempLM35()//10mV/C
        {
            double lm30Vread = (ReadADC(3))*100;//convert to 1/10 C 25C = (0.01x25)*1000
            return lm30Vread;
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
            string rootPath = Tb_FolderLoc.Text + @"\" + Tb_LaserPN.Text;//production data + laser folder alreary set i.e. 015335
            string folderName = @"\" + Tb_WorkOrder.Text + @"\" + Tb_SerNb.Text;//work order/ipo and laser assembly serial number added now
            string filePathFold = rootPath + folderName;
            string txtName =    @"\" + Tb_SerNb.Text + ".txt";//file .txt name

            filePathRep = filePathFold + txtName; // \\officeserver\Production Test Data\iFLEX IRIS Test Data\015335\IPO..........\0052.....

            await Task.Delay(1);
            // Create folder
            try
            {
                if (Directory.Exists(filePathFold)) { MessageBox.Show("Folder Exist"); }
                else { DirectoryInfo di = Directory.CreateDirectory(filePathFold); }
            }
            catch (Exception err) { MessageBox.Show(err.Message); }
            finally { }

            //create file
            try
            {
                using (FileStream fs = File.Create(filePathRep))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(txtName + Footer + Footer);
                    fs.Write(info, 0, info.Length);
                    Tb_txtFilePathRep.Text = filePathRep;
                    return true;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
            }
            finally { }
        }
        //======================================================================
        private void Tb_LaserPN_Leave(object sender, EventArgs e) { ReadDbs(); }
        //======================================================================
        private void Bt_SetBurnin_Click(object sender, EventArgs e) { Task<bool> burninData = SendBurninData(); }
        //======================================================================
        private async Task<bool> SendBurninData()
        {
            double brninPwDbl = (Convert.ToDouble(Tb_minMaxPw.Text))*10;//50.1 >> 501
            string brninPw = brninPwDbl.ToString("0000");

            if (Bt_SetBurnin.BackColor == Color.Coral)
            {
                string chkBxStateExtPwCtrlBrin      = null;
                string chkBxStateEnblSetBrin        = null;
                string chkBxStateDigitModSetBrin    = null;
                string chkBxStateAnlgModSetBrin     = null;

                this.Cursor = Cursors.WaitCursor;

                if (ChkBx_IntPwCtrlBurn.Checked == true) { chkBxStateExtPwCtrlBrin = StrEnable; }
                else { chkBxStateExtPwCtrlBrin = StrDisable; }

                if (ChkBx_EnInvBurn.Checked == true) { chkBxStateEnblSetBrin = StrEnable; }
                else { chkBxStateEnblSetBrin = StrDisable; }

                if (ChkBx_DiditModInvBurn.Checked == true) { chkBxStateDigitModSetBrin = StrEnable; }
                else { chkBxStateDigitModSetBrin = StrDisable; }

                if (ChkBx_AlgModInvBurn.Checked == true) { chkBxStateAnlgModSetBrin = StrEnable; }
                else { chkBxStateAnlgModSetBrin = StrDisable; }

                bool finalSet = await SendToSerial(CmdTestMode, StrEnable, 300, 9);
                finalSet = await SendToSerial(CmdSetInOutPwCtrl, chkBxStateExtPwCtrlBrin, 300, 9);
                finalSet = await SendToSerial(CmdEnablLogicvIn, chkBxStateEnblSetBrin, 300, 9);
                finalSet = await SendToSerial(CmdsetTTL, chkBxStateDigitModSetBrin, 300, 9);
                finalSet = await SendToSerial(CmdAnalgInpt, chkBxStateAnlgModSetBrin, 300, 9);
                finalSet = await SendToSerial(CmdLaserEnable, StrEnable, 300, 9);
                finalSet = await SendToSerial(CmdTestMode, StrDisable, 300, 9);//run mode
                finalSet = await SendToSerial(CmdSetLsPw, brninPw, 300, 9);//set internal power but not used if set to external

                this.Cursor = Cursors.Default;
                Bt_SetBurnin.BackColor = Color.LawnGreen;
            }

            else if (Bt_SetBurnin.BackColor == Color.LawnGreen) { Bt_SetBurnin.BackColor = Color.Coral; }
            return true;
        }
        //======================================================================
        private void Bt_ShipState_Click(object sender, EventArgs e) { Task<bool> sendShpData = SendShpData(); }
        //======================================================================
        private async Task<bool> SendShpData()
        {
            bool finalSet = false;

            if (Bt_ShipState.BackColor == Color.Aqua) {

                string chkBxStateExtPwCtrl = string.Empty;//null
                string chkBxStateEnblSet = string.Empty;
                string chkBxStateDigitModSet = string.Empty;
                string chkBxStateAnlgModSet = string.Empty;
                string chkBxStateOnAtPon = string.Empty;
                double finalPowerDbl = Convert.ToDouble(Tb_SoftNomPw.Text) * 10;
                string finalPower = finalPowerDbl.ToString("0000");

                this.Cursor = Cursors.WaitCursor;

                if (ChkBx_ExtPwCtrl.Checked == true) { chkBxStateExtPwCtrl = StrEnable; }
                else { chkBxStateExtPwCtrl = StrDisable; }

                if (ChkBx_EnableSet.Checked == true) { chkBxStateEnblSet = StrEnable; }
                else { chkBxStateEnblSet = StrDisable; }

                if (ChkBx_DigitModSet.Checked == true) { chkBxStateDigitModSet = StrEnable; }
                else { chkBxStateDigitModSet = StrDisable; }

                if (ChkBx_AnlgModSet.Checked == true) { chkBxStateAnlgModSet = StrEnable; }
                else { chkBxStateAnlgModSet = StrDisable; }

                if (ChkBx_SoftEnStart.Checked == true) { chkBxStateOnAtPon = StrEnable;  }
                else { chkBxStateOnAtPon = StrDisable; }

                finalSet = await SendToSerial(CmdTestMode, StrEnable, 300, 9);
                finalSet = await SendToSerial(CmdSetInOutPwCtrl, chkBxStateExtPwCtrl, 300, 9);
                finalSet = await SendToSerial(CmdEnablLogicvIn, chkBxStateEnblSet, 300, 9);
                finalSet = await SendToSerial(CmdsetTTL, chkBxStateDigitModSet, 300, 9);
                finalSet = await SendToSerial(CmdAnalgInpt, chkBxStateAnlgModSet, 300, 9);

                finalSet = await SendToSerial(CmdTestMode, StrDisable, 300, 9);
                finalSet = await SendToSerial(CmdSetLsPw, finalPower, 300, 9);

                finalSet = await WriteResToDb(1);

                this.Cursor = Cursors.Default;
                Bt_ShipState.BackColor = Color.LawnGreen;
            }
            else if (Bt_ShipState.BackColor==Color.LawnGreen) { Bt_ShipState.BackColor = Color.Aqua; }
            finalSet = await SendToSerial(CmdLaserStatus, StrDisable, 300, 9);
            return true;
        }
        //======================================================================
        private void ReadDbs()
        {
            Rt_ReceiveDataUSB.Clear();
            //int dbArrayIdx = 0; //just running index to locate the line on table...not used 
            bool entryOK = false;
            string readstuff = string.Empty;
            string[] laserParameters = new string[30];
            
            try {
                con.Open();
                cmd = new SqlCommand("SELECT * FROM " + "Laser_Setup_Config", con);
                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read()) {
                //dbArrayIdx++;
                readstuff = rdr["PartNumber"].ToString();//read each partnumber

                    if (readstuff.Contains(Tb_LaserPN.Text)) {//if match
                        entryOK = true;

                        Lbl_MdlName.Text = rdr["Description"].ToString();
                        Lbl_MdlName.ForeColor = Color.Green;
                        Lbl_Wlgth1.Text = rdr["Wavelength"].ToString().PadLeft(4, '0').TrimStart('0');
                        Lbl_Wlgth1.ForeColor = Color.Green;
                        Tb_Wavelength.Text = Lbl_Wlgth1.Text;
        
                        Lbl_NomPowerDtbas.Text = rdr["NominalPower"].ToString().PadLeft(5, '0').TrimStart('0');//note power in mw ! and not 1/10mW
                        Lbl_NomPowerDtbas.ForeColor = Color.Green;
                        Tb_NomPw.Text = Lbl_NomPowerDtbas.Text;
                        Tb_maxMaxPw.Text = rdr["MaxPower"].ToString().PadLeft(5, '0').TrimStart('0');
                        Tb_minMaxPw.Text = rdr["MinPower"].ToString().PadLeft(5, '0').TrimStart('0');
                        Tb_SoftNomPw.Text = rdr["SoftwareNomPower"].ToString().PadLeft(5, '0').TrimStart('0');

                        Tb_SetAdd.Text = rdr["LaserAddress"].ToString().PadLeft(2, '0');

                        Tb_TECpoint.Text = rdr["TEC_BlockTemperature"].ToString();

                        Tb_PwToVcal.Text = rdr["PowerMonitorVoltage"].ToString().PadLeft(5, '0').TrimStart('0');

                        if ((rdr["PowerControlSource"].ToString()) == "Internal  ") { ChkBx_ExtPwCtrl.Checked = true; }
                        else { ChkBx_ExtPwCtrl.Checked = false; }

                        if ((rdr["EnableLine"].ToString()) == "Norm      ") { ChkBx_EnableSet.Checked = false; }
                        else { ChkBx_EnableSet.Checked = true; }//inverted

                        if ((rdr["DigitalModulation"].ToString()) == "Norm      ") { ChkBx_DigitModSet.Checked = false; }
                        else { ChkBx_DigitModSet.Checked = true; }//inverted

                        if ((rdr["AnalogueModulation"].ToString()) == "Norm      ") { ChkBx_AnlgModSet.Checked = false; }
                        else { ChkBx_AnlgModSet.Checked = true; }//inverted

                        if ((rdr["SoftwareEnableStartup"].ToString()) == "ON") { ChkBx_SoftEnStart.Checked = true; }
                        else { ChkBx_SoftEnStart.Checked = false; }

                        Lbl_LaserType.Text = rdr["LaserType"].ToString().TrimEnd(' ');
                        Lbl_LaserType.ForeColor = Color.Green;

                        if (rdr["LaserType"].ToString()         == "CLM       ") { laserType = CLM_Ls; }
                        else if (rdr["LaserType"].ToString()    == "MKT       ") { laserType = MKT_Ls; }
                        else if (rdr["LaserType"].ToString()    == "CCM       ") { laserType = CCM_Ls; }
                                                
                        //MessageBox.Show("PN: " + readstuff + " @ " + dbArrayIdx.ToString() + "\n\n" + "Enter 'Diode Max. Current Limit' Value now");
                        Tb_LaserPN.ForeColor = Color.Green;
                        Tb_MaxILimit.Focus();
                        Tb_MaxILimit.Clear();

                        break;
                    }
                }
                con.Close();
            }
            catch (Exception e)   { MessageBox.Show("Dtb Read Error " + e.ToString()); }
            if (entryOK == false) { MessageBox.Show("No parts in db\nTry again or contact engineering\n"); }

            //if (rdr != null) { rdr.Close(); }
            //if (con != null) { con.Close(); }
        }
        //======================================================================
        private void SetLaserType(int lsType)//used for send and receive
        {
            switch (lsType) {

                case MKT_Ls:
                    CmdRdSerialNo   = "10";
                    CmdRdFirmware   = "09";  
                    CmdRdWavelen    = "11";  
                    CmdSetLsPw      = "04";  
                    CmdRdLaserPow   = "05";  
                    CmdRatedPower   = "14";  
                    CmdLaserStatus  = "06";
                    Bt_BasePltTempComp.Enabled = false;
                    label28.Visible = true;
                    Tb_MKTLasEnable.Visible = true;
                    break;

                case CLM_Ls:
                    CmdRdSerialNo   = "04";  //cmd 10 MKT
                    CmdRdFirmware   = "06";  //cmd 09 MKT
                    CmdRdWavelen    = "08";  //cmd 11 MKT
                    CmdSetLsPw      = "03";  //cmd 04 MKT
                    CmdRdLaserPow   = "44";  //cmd 05 MKT 
                    CmdRatedPower   = "47";  //cmd 14 MKT
                    CmdLaserStatus  = "14";  //cmd 06 MKT bit 5 
                    Bt_BasePltTempComp.Enabled = true;
                    label28.Visible = false;
                    Tb_MKTLasEnable.Visible = false;
                    break;

                case CCM_Ls:
                    CmdRdSerialNo   = "04";  
                    CmdRdFirmware   = "06";  
                    CmdRdWavelen    = "08";  
                    CmdSetLsPw      = "03";
                    CmdRdLaserPow   = "44";  
                    CmdRatedPower   = "47";  
                    CmdLaserStatus  = "14";
                    Bt_BasePltTempComp.Enabled = true;
                    label28.Visible = false;
                    Tb_MKTLasEnable.Visible = false;
                    break;

                default:
                    MessageBox.Show("No Laser Type Loaded");
                    break;
            }

            #region test arrays

            bulkSetLaserIO = new string[7, 2] {   //the rest of the string is build with case...
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetTECena_dis,     StrDisable},
            { CmdSetInOutPwCtrl,    StrDisable },       //external PCON
            { CmdAnalgInpt,         StrDisable },       //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },       //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable } };      //Inv. TTL line in nothing connected

            bulkSetVarialble = new string[14, 2] {
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

            bulkSetdefaultCtrl = new string[6, 2] {
            {CmdTestMode,       StrEnable  },
            {CmdRatedPower,     StrDisable },
            {CmdSetPwMonOut,    StrDisable },
            {CmdSetVgaGain,     StrDisable },
            {CmdSetOffstVolt,   StrDisable },      //Offset 2.500V
            {CmdSetPwCtrlOut,   StrDisable } };    //Internal PCON 2.500V
            
            bulkSetFinalSetup = new string[7, 2] {
            {CmdTestMode,           StrEnable},
            {CmdSetCalAPw,          StrEnable},
            {CmdSetCalBPw,          StrDisable},
            {CmdSetCalAPwtoVint,    StrEnable },
            {CmdSetCalBPwtoVint,    StrDisable},
            {CmdSetCalAVtoPw,       StrEnable},
            {CmdSetCalBVtoPw,       StrDisable} };

            bulkSetTEC = new string[6, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetTECTemp,        StrDisable },
            { CmdSetTECkp,          StrDisable },
            { CmdSetTECki,          StrDisable },
            { CmdSetTECsmpTime,     StrDisable },
            { CmdSetTECena_dis,     StrEnable} };

            bulkSetTEC532 = new string[2, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetTECTemp,        StrDisable } };

            bulkSetRstClk = new string[5, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdOperatingHr,       StrDisable },
            { CmdRstTime,           StrDisable },
            { CmdTestMode,          StrDisable },
            { CmdOperatingHr,       StrDisable } };
            //=================================================
            bulkSetBurnin = new string[7, 2] {
            { CmdTestMode,          StrEnable  },
            { CmdSetInOutPwCtrl,    StrDisable },     //external PCON
            { CmdAnalgInpt,         StrDisable },     //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },     //Non Inv. Laser Enable 0 >> 0/1 enable
            { CmdsetTTL,            StrEnable  },     //Inv. TTL line in 1  line to 0
            { CmdLaserEnable,       StrEnable  },
            { CmdTestMode,          StrDisable } };
            //=================================================
            bulkSetVga = new string[6, 2] {
            { CmdLaserEnable,       StrDisable },
            { CmdTestMode,          StrEnable  },
            { CmdSetInOutPwCtrl,    StrDisable },     //external PCON
            { CmdAnalgInpt,         StrDisable },     //Non Inv. PCON
            { CmdEnablLogicvIn,     StrDisable },     //Non Inv. Laser Enable
            { CmdsetTTL,            StrEnable } };    //Inv. TTL line in
                                                      //=================================================
            analogRead = new string[4, 2] {//read analog inputs
            { CmdTestMode,          StrEnable },
            { CmdRdLaserPow,       StrDisable },
            { CmdRdPwSetPcon,       StrDisable },
            { CmdCurrentRead,       StrDisable } };

            analogRead2 = new string[6, 2] {//read all analog inputs
            { CmdTestMode,          StrEnable },
            { CmdRdCmdStautus2,     StrDisable },
            { CmdRdLaserPow,        StrDisable },
            { CmdRdPwSetPcon,       StrDisable },
            { CmdCurrentRead,       StrDisable },
            { CmdRdTecTemprt,       StrDisable } };
            /*
            analogRead2 = new string[7, 2] {//read all analog inputs
            { CmdTestMode,          StrEnable },
            { CmdRdCmdStautus2,     StrDisable },
            { CmdRdLaserPow,        StrDisable },
            { CmdRdPwSetPcon,       StrDisable },
            { CmdCurrentRead,       StrDisable },
            { CmdRdTecTemprt,       StrDisable },
            { CmdRdBplateTemp,      StrDisable } };
            */
            setLaserType = new string[4, 2] {
            { CmdTestMode,          StrEnable },
            { CmdSetLaserType,      StrDisable },
            { CmdRdLaserType,       StrDisable },
            { CmdTestMode,          StrDisable } };

            #endregion

        }
        //======================================================================
        private void Bt_RstClk_Click(object sender, EventArgs e) { Task<bool> rstclk = RESETclk(); }
        //======================================================================
        private async Task<bool> RESETclk() {
            this.Cursor = Cursors.WaitCursor;
            Bt_RstClk.BackColor = Color.LawnGreen;
            bool lsrst = await LoadGlobalTestArray(bulkSetRstClk);
            Bt_RstClk.BackColor = Color.Coral;
            this.Cursor = Cursors.Default;
            return true; }
        //======================================================================
        private void Bt_Set532Temp_Click(object sender, EventArgs e) { Task<bool> setTmp5352 = LoadGlobalTestArray(bulkSetTEC532); }
        //======================================================================
        private void Bt_SetFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog { ShowNewFolderButton = true };
            // Show the FolderBrowserDialog. 
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                Tb_FolderLoc.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
                // Get Program Files location.
                string programfiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                // Get Common Program Files location.
                string commonProgramfiles = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles, Environment.SpecialFolderOption.None);
            }
        }
        //======================================================================
        private void Tb_MaxILimit_Leave(object sender, EventArgs e) {
 
            if (int.TryParse(Tb_MaxILimit.Text, out int dummyInt) == true)
            {
                Tb_MaxLsCurrent.Text = Tb_MaxILimit.Text;

                if ((LoadCurAndPwLimits()) == true)//reset at init...?
                {
                    Tb_MaxILimit.ForeColor = Color.Green;
                    tabControl1.TabPages[1].Enabled = true;
                }
                else
                {
                    Tb_MaxILimit.ForeColor = Color.OrangeRed;
                    MessageBox.Show("incorrect value");
                }
            }
            else { MessageBox.Show("Enter numerical current"); }
        }
        //======================================================================
        private void Rtb_ComList_DoubleClick(object sender, EventArgs e) { Rtb_ComList.Clear(); }
        //======================================================================
        private void Bt_SetPower03_Click(object sender, EventArgs e) { Task<bool>  finalSet = SendToSerial(CmdSetLsPw, Tb_SetPower03.Text, 300, 9); }
        //======================================================================
        private void Tb_User_Click(object sender, EventArgs e) { Tb_User.Clear(); }
        private void Tb_WorkOrder_Click(object sender, EventArgs e) { Tb_WorkOrder.Clear(); }
        private void Tb_SerNb_Click(object sender, EventArgs e) { Tb_SerNb.Clear(); }
        private void Tb_TecSerNumb_Click(object sender, EventArgs e) { Tb_TecSerNumb.Clear(); }
        private void Tb_LaserPN_Click(object sender, EventArgs e) { Tb_LaserPN.Clear(); }
        private void Tb_MaxILimit_Click(object sender, EventArgs e) { Tb_MaxILimit.Clear(); }
        //======================================================================
        private void Tb_WorkOrder_Leave(object sender, EventArgs e) { Tb_WorkOrder.ForeColor = Color.Green; }
        //======================================================================
        private void Tb_SerNb_Leave(object sender, EventArgs e) { Tb_SerNb.ForeColor = Color.Green; }
        //======================================================================
        private void Tb_TecSerNumb_Leave(object sender, EventArgs e) { Tb_TecSerNumb.ForeColor = Color.Green; }
        //======================================================================
        private void Tb_User_KeyDown(object sender, EventArgs e) { Tb_User.ForeColor = Color.Green; }
        //======================================================================
        private void Tb_LaserPN_MouseLeave(object sender, EventArgs e) { MessageBox.Show("\n\n" + "Enter 'Diode Max. Current Limit' Value now"); }
        //not used
        //======================================================================
        private void Bt_StopTest_Click(object sender, EventArgs e) {
            stopLoop = true;
            Rtb_ComList.AppendText("true\n");
        }
        //======================================================================
        private void Bt_EnableIObox_Click(object sender, EventArgs e)
        {
                if (Bt_EnableIObox.BackColor == Color.SandyBrown)
                {
                Bt_EnableIObox.BackColor = Color.LawnGreen;
                GrBx_CutomIO.Enabled = true;
                GrBx_532Tset.Enabled = true;
                    
                }
                else if (Bt_EnableIObox.BackColor == Color.LawnGreen)
                {
                Bt_EnableIObox.BackColor = Color.SandyBrown;
                GrBx_CutomIO.Enabled = false;
                GrBx_532Tset.Enabled = false;
            }
        }
        //======================================================================
        private void Bt_EnableDBstring_Click_1(object sender, EventArgs e) { GrBx_DatabaseString.Visible = true; }
        //======================================================================
        private void Bt_SetLsType_Click(object sender, EventArgs e) { Task<bool> usbadd = SetLaserFirm(); }
        //======================================================================
        private async Task<bool> SetLaserFirm() { //decouple to get async, async can be accessible via event....
            Lbl_LsType.Text = "000";
            bool setad = await LoadGlobalTestArray(setLaserType);//read and set laser type
            setad = await SetComsUSB();//should disconnect usb laser needs to restart
            return true;
        }
        //======================================================================
        private void Tb_Wavelength_Click(object sender, EventArgs e) { Tb_Wavelength.Clear(); }
        //======================================================================
        private void Tb_Wavelength_Leave(object sender, EventArgs e)
        {
            int tbWavelngth = Convert.ToInt16(Lbl_Wlgth1.Text);

            if (int.TryParse(Tb_Wavelength.Text, out int dummyInt) == true)
            {
                if (dummyInt <= (tbWavelngth + 5) && dummyInt >= (tbWavelngth - 5))
                {
                    Tb_Wavelength.ForeColor = Color.Green;
                }
                else
                {
                    Tb_Wavelength.ForeColor = Color.OrangeRed;
                    Tb_Wavelength.Text = "0000";
                    MessageBox.Show("Out of range value");
                }
            }
            else { MessageBox.Show("Enter numerical integer Wavelength"); }
        }
        //======================================================================
        //======================================================================
    }
    //======================================================================
    //======================================================================
}
//======================================================================
//======================================================================

