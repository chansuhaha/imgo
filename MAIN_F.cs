using C1.Win.C1FlexGrid;
using IMCC_ALL_OF_POE_GET.Adaptors;
using IMCC_ALL_OF_POE_GET.Model;
using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace IMCC_ALL_OF_POE_GET
{
    public partial class MAIN_F : Form
    {
        private const int CP_NOCLOSE_BUTTON = 0x200;

        //public string sDetector = "";
        private bool IsDeveloper = false;
        int multi_Channel_Count = 4;
        
        private const string sMulti_Channel_Detector = "GASTRON_GTM_1000G";
        
        // open 소스 형상관리 확인
        //DataTable dtxml;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }

        private string[] sDectors;
        private string sDetectors_Combolist;
        private string currentPath;
        private string currentPathLog;
        private ListViewItem li;
        private Thread tdMain;
        private Thread tdCount;
        //private DateTime dtLogDeletePre;
        private string sBasePath = AppDomain.CurrentDomain.BaseDirectory;
        public static int thCount = 0;
        public static string sPOE_Start_IP = "";
        public static ManualResetEvent _doneEvent = new ManualResetEvent(false);
        Stopwatch swUpload = new Stopwatch();
        private int iLogSaveDays = 7;


        public MAIN_F()
        {
            InitializeComponent();

            
        }

        private void killProcess()
        {
            Process[] process = null;
            Process process_Curr = null;

            string sAppName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;

            process_Curr = Process.GetCurrentProcess();
            process = Process.GetProcessesByName(sAppName);

            if (process.Count() > 1)
            {
                for (int i = 0; i < process.Count(); i++)
                {
                    if (process[i].Id != process_Curr.Id)
                    {
                        process[i].Kill();
                    }
                }
            }

            process = null;
            process_Curr = null;
        }

        private void MAIN_F_Load(object sender, EventArgs e)
        {
#if false
            //this.txtImsi.Visible = true;
#else
            //this.txtImsi.Visible = false;
#endif
            killProcess();

            #region ��ư ��Ʈ�� ����
            this.Start_Gubun("I");

            #endregion

            #region �αװ�� Ȯ��
            currentPath = Environment.CurrentDirectory;
            currentPathLog = Environment.CurrentDirectory + @"\Log";

            DirectoryInfo di = new DirectoryInfo(currentPathLog);

            if (!di.Exists)
            {
                Directory.CreateDirectory(currentPathLog);
            }
            #endregion

            #region �������� �ҷ�����
            string tempCheckSetting = "";

            tempCheckSetting = ConfigurationManager.AppSettings["CTL_NODE"].ToString();

            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Control Node Information does not setting", "Warning!!");
                Close();
                return;
            }

            this.txtCtlNode.Text = tempCheckSetting;
            this.txtCtlNode.ReadOnly = true;



            tempCheckSetting = ConfigurationManager.AppSettings["ConnectionString"].ToString();

            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Connection String not Set.", "Warning!!");
                Close();
                return;
            }

            CommonData.sConnectionString = tempCheckSetting;



            tempCheckSetting = ConfigurationManager.AppSettings["Title"].ToString();
            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Title not Set.", "Warning!!");
                Close();
                return;
            }
            this.Text = tempCheckSetting;



            tempCheckSetting = ConfigurationManager.AppSettings["Models"].ToString();
            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Detector Models does not exist in Config.", "Warning!!");
                Close();
                return;
            }

            sDetectors_Combolist = tempCheckSetting;
            sDectors = sDetectors_Combolist.Split(new char[] { '|' });



            tempCheckSetting = ConfigurationManager.AppSettings["Channel_Count"].ToString();
            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Detector Channel Count does not set.", "Warning!!");
                Close();
                return;
            }

            try
            {
                multi_Channel_Count = int.Parse(tempCheckSetting);
            }
            catch
            {
            }

            

            tempCheckSetting = ConfigurationManager.AppSettings["IsDeveloper"].ToString();
            if (tempCheckSetting == null || tempCheckSetting.Length < 1 || tempCheckSetting.Equals("0"))
            {
                IsDeveloper = false;
            }
            else
            {
                IsDeveloper = true;
            }

            tempCheckSetting = ConfigurationManager.AppSettings["LogSaveDays"].ToString();

            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Log Save Days does not exit", "Warning!!");
                Close();
                return;
            }
            try
            {
                iLogSaveDays = int.Parse(tempCheckSetting);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Log Save Days setting does not correct format. it's numeric only", "Warning!!");
                Close();
                return;
            }

            tempCheckSetting = ConfigurationManager.AppSettings["POE_Start_IP"].ToString();

            if (tempCheckSetting == null || tempCheckSetting.Length < 1)
            {
                MessageBox.Show("Start IP Address does not exist", "Warning!!");
                Close();
                return;
            }
            try
            {
                sPOE_Start_IP = tempCheckSetting;
                IPAddress address;
                if (IPAddress.TryParse(sPOE_Start_IP, out address) == false) throw new Exception("Start IP Address Format does not collect.");
                this.txtIPAddress.Text = sPOE_Start_IP;
                this.txtIPAddress.ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Close();
                return;
            }






            #endregion �������� �ҷ�����

            #region ������ �� ���� �ε� �ʱ�ȭ
            this.DectorModelSetting();

            //sDetector = this.DetectorName;
            //this.Text = string.Format("{0}  {1}",sDetector , this.Text);
            #endregion

            #region ���� ����� Ŭ���� �ε�
            CommonData.CommonDataInit(this);
            #endregion

            #region �α� ����Ʈ ��Ʈ�� ����
            this.LogList_Init();
            #endregion




            #region ����͸��� �׸��� �ʱ�ȭ
            try
            {
                this.POE_NodeDataInit(this.grdView);
                this.GridGroupSetting(this.grdView);
            }
            catch { }
            #endregion

            #region map ������ �׸��� �ʱ�ȭ
            try { 
                this.POE_NodeDataInit(this.grdMap);
                this.GridGroupSetting(this.grdMap);
            }
            catch { }
            #endregion

            #region �׸��� ��Ÿ�� ������
            this.LoadGridStyle();
            #endregion

            #region �۾������� Count ǥ��
            this.ThreadCountLable();
            #endregion

            #region ������ ���
            if (IsDeveloper == false) mainNodeTab.TabPages.Remove(this.tabStatus);
            #endregion

            this.WRITE("F", 0, this.txtCtlNode.Text, "0", "0", string.Format(" Application started..."));
        }

        private void ThreadCountLable()
        {
            this.lblAvailable.Text = "- ��밡���� �񵿱� I/O ������ �� : ";
            this.lblMax.Text = "- ��밡���� �۾������� ��: ";
            this.lblUsedThread.Text = "- ������� ������ �� :";
            this.lblWorkerThread.Text = "- �ִ� �۾��� ������� :";
            this.lblCompletitionThread.Text = "- �ִ� �񵿱� I/O ������� :";
        }

        #region �׸��� Style ���� ����
        private void LoadGridStyle()
        {

            this.grdView.Styles["Normal"].Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.Flat;
            this.grdMap.Styles["Normal"].Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.Flat;

            CellStyle csRed = this.grdView.Styles.Add("redStyle");
            csRed.Font = new Font("Tahoma", 10, FontStyle.Bold);
            csRed.ForeColor = Color.Red;

            CellStyle csBlack = this.grdView.Styles.Add("blackStyle");
            csBlack.Font = new Font("Tahoma", 10, FontStyle.Bold);
            csBlack.ForeColor = Color.Black;

            CellStyle csyellow = this.grdView.Styles.Add("yellowStyle");
            csyellow.Font = new Font("Tahoma", 10, FontStyle.Bold);
            csyellow.ForeColor = Color.Yellow;

        }
        #endregion �׸��� Style ���� ����



        private bool isIPAddress(string addrString)
        {
            bool retVal = false;
            IPAddress address;
            if (IPAddress.TryParse(addrString, out address))
            {
                retVal = true;
            }
            return retVal;
        }



        #region ������ Model ����
        private void DectorModelSetting()
        {
            int forcnt = sDectors.Length;
            for (int i = 0; i < forcnt; i++)
            {
                this.cboDetector.Items.Add(sDectors[i]);
            }
            cboDetector.SelectedIndex = 0;
            //cboDetector.Enabled = false;
        }
        #endregion ������ Model ����

        #region ��� ����, �����ư �̺�Ʈ ó��
        private void btnExit_Click(object sender, EventArgs e)
        {
            StopComm();

            Thread.Sleep(100);

            this.Close();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.grdView.Rows.Count > 1)
                {
                    Start_Gubun("S");

                    StartComm();

                    // �̺�Ʈ ����
                    this.mainNodeTab.SelectedIndexChanged -= this.mainNodeTab_SelectedIndexChanged;

                    this.WRITE("F", 0, this.txtCtlNode.Text, "0", "0", string.Format(" Application Communication started..."));

                }
                else
                {
                    MessageBox.Show("Point information does not exists");
                }
            }
            catch { }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            StopComm();

            this.mainNodeTab.SelectedIndexChanged += this.mainNodeTab_SelectedIndexChanged;

            Thread.Sleep(500);

            try
            {
                Start_Gubun("T");
            }
            catch (Exception es)
            {
                MessageBox.Show(es.Message.ToString());
            }
            this.WRITE("F", 0, this.txtCtlNode.Text, "0", "0", string.Format(" Application Communication stopped..."));

        }

        private void btnNodeSave_Click(object sender, EventArgs e)
        {
            try
            {
                //NodeDataChangeSave();
                SetMidasInfo();
                this.POE_NodeDataInit(this.grdMap);
                this.GridGroupSetting(this.grdMap);
                MessageBox.Show("Saved Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error [Invalid entry : {0}]", ex.Message));
            }

        }

        #endregion ��� ����, �����ư �̺�Ʈ ó��

        #region ��� ���� �� ��Ʈ�� ó�� ( ������ ����)
        private void Start_Gubun(string sGubun)
        {
            //this.menuStrip1.Visible = false;
            

            if (sGubun.Equals("S"))
            {
                this.btnStart.Enabled = false;
                this.btnStop.Enabled = true;
                this.btnExit.Enabled = false;
                this.btnNodeSave.Enabled = false;
                this.grdMap.AllowEditing = false;
                this.btnPointAdd.Enabled = false;
                this.btnPointDelete.Enabled = false;

                this.mainNodeTab.SelectedIndex = 0;
                this.mainNodeTab.TabPages.Remove(this.tabMapManage);
                this.pictureBox3.Enabled = true;
                this.pictureBox3.Visible = true;
            }
            else if (sGubun.Equals("T"))
            {
                this.btnStart.Enabled = true;
                this.btnStop.Enabled = false;
                this.btnExit.Enabled = true;
                this.btnNodeSave.Enabled = true;
                this.grdMap.AllowEditing = true;
                this.btnPointAdd.Enabled = true;
                this.btnPointDelete.Enabled = true;
                this.mainNodeTab.TabPages.Add(this.tabMapManage);
                this.pictureBox3.Enabled = false;
                this.pictureBox3.Visible = false;

            }
            else if (sGubun.Equals("I"))
            {
                this.btnStart.Enabled = true;
                this.btnStop.Enabled = false;
                this.btnExit.Enabled = true;
                this.btnNodeSave.Enabled = true;
                this.grdMap.AllowEditing = true;
                this.btnPointAdd.Enabled = true;
                this.btnPointDelete.Enabled = true;
                this.pictureBox3.Enabled = false;
                this.pictureBox3.Visible = false;

            }
            this.btnUp.Visible = false;
            this.btnDown.Visible = false;

            this.txtPort.Text = "502";
            this.cboDetector.DropDownStyle = ComboBoxStyle.DropDownList;
            
        }
        #endregion ��� ���� �� ��Ʈ�� ó�� ( ������ ����)


        private void LogList_Init()
        {
            /*
            _lv.Scrollable = true;
            _lv.View = View.Details;
            ColumnHeader header = new ColumnHeader();
            header.Text = "";
            header.Name = "col1";
            _lv.Columns.Add(header);
            _lv.Columns[0].Width = _lv.Width - 20;
            */
        }

        private void POE_NodeDataInit(C1FlexGrid grd)
        {
            try
            {
                string strPointPath = AppDomain.CurrentDomain.BaseDirectory + string.Format(@"\Config\Point_List.xml", "");
                FileInfo fi = new FileInfo(strPointPath);

                bool bSave = false;


                grd.Cols["DETECTOR"].ComboList = this.sDetectors_Combolist;
                grd.Rows[0].StyleNew.Font = new Font(Font, FontStyle.Bold);
                grd.Rows.Count = 1;


                if (fi.Exists)
                {

                    XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<POE_Node>));

                    using (StreamReader rd = new StreamReader(strPointPath))
                    {
                        poe_list = xs.Deserialize(rd) as ObservableCollection<POE_Node>;
                    }


                    for (int k = 0; k < poe_list.Count; k++)
                    {
                        grd.Rows.Add();
                        grd.Rows[k + 1]["SEQ"] = poe_list[k].SEQ;
                        grd.Rows[k + 1]["CHANNEL"] = poe_list[k].Channel;
                        grd.Rows[k + 1]["POINT_NM"] = poe_list[k].POINT_NM;

                        grd.Rows[k + 1]["DETECTOR"] = poe_list[k].DETECTOR;
                        grd.Rows[k + 1]["POE_IP"] = poe_list[k].POE_IP;
                        grd.Rows[k + 1]["POE_PORT"] = poe_list[k].POE_PORT;

                        // ctl ��� ������ �ܺο��� ���� �Ǿ��� ��� ó��
                        if (bSave ==false && this.txtCtlNode.Text.Equals(poe_list[k].CTL_NODE) == false)
                        {
                            bSave = true;
                        }

                        grd.Rows[k + 1]["CTL_NODE"] = this.txtCtlNode.Text;
                        grd.Rows[k + 1]["F_ADDRESS"] = poe_list[k].F_ADDRESS;
                        grd.Rows[k + 1]["G_ADDRESS"] = poe_list[k].G_ADDRESS;
                        grd.Rows[k + 1]["VALID_YN"] = poe_list[k].VALID_YN;
                    }

                    if (grd.Equals(this.grdMap))
                    {
                        grd.AllowEditing = true;
                        grd.Cols["SEQ"].AllowEditing = false;
                        grd.Cols["CHANNEL"].AllowEditing = false;
                        grd.Cols["DETECTOR"].AllowEditing = true;
                        //grd.Cols["CTL_NODE"].AllowEditing = false;

                        // ctl ��� ������ �ܺο��� ���� �Ǿ��� ��� ����
                        if (bSave)
                        {
                            SetMidasInfo();
                        }
                    }
                    else
                    {
                        grd.AllowEditing = false;
                    }
                }
            }
            catch (Exception es)
            {
                throw es;
            }
        }


        private void DectctorListSetting()
        {
            string[] sClassName = Type.GetType(string.Format("{0}.{1}", "IMCC_ALL_OF_POE_GET.Adaptors.Detector", "*")).GetEnumNames();

            for (int i = 0; i < sClassName.Length; i++)
            {
                this.cboDetector.Items.Add(new ListViewItem(sClassName));
            }
        }


        private string GetClassNameList()
        {
            string [] sClassName = Type.GetType(string.Format("{0}.{1}", "IMCC_ALL_OF_POE_GET.Adaptors.Detector", "*")).GetEnumNames();
            string sRetVal = "";
            for(int i = 0; i< sClassName.Length; i++)
            {
                if (i == sClassName.Length - 1)
                {
                    sRetVal += string.Format("{0}", sClassName);
                }
                else
                {
                    sRetVal += string.Format("{0}|", sClassName);
                }
            }

            return sRetVal;
        }



        #region Point ������ �����Ѵ�.

        ObservableCollection<POE_Node> poe_list = null;

        ObservableCollection<POE_Node> run_detectors = null;
        public bool SetMidasInfo()
        {

            bool bResult = false;


            string strPointPath = AppDomain.CurrentDomain.BaseDirectory + string.Format(@"\Config\Point_List.xml");

            poe_list.Clear();


            POE_Node point = null;
            for(int i = 1; i< grdMap.Rows.Count; i++)
            {

                string pointName = string.Format("{0}_{1}_{2}", this.txtCtlNode.Text, i, Convert.ToUInt16(this.grdMap.Rows[i]["CHANNEL"]));
                point = new POE_Node();
                point.SEQ = i;
                point.Channel   = Convert.ToUInt16( this.grdMap.Rows[i]["CHANNEL"]);
                point.POINT_NM  = this.grdMap.Rows[i]["POINT_NM"]==null ? pointName : this.grdMap.Rows[i]["POINT_NM"].ToString();
                point.VALID_YN  = this.grdMap.Rows[i]["VALID_YN"].ToString();
                point.DETECTOR  = this.grdMap.Rows[i]["DETECTOR"].ToString();
                point.POE_IP    = this.grdMap.Rows[i]["POE_IP"].ToString();
                point.POE_PORT  = this.grdMap.Rows[i]["POE_PORT"].ToString();
                point.G_ADDRESS = this.grdMap.Rows[i]["G_ADDRESS"].ToString();
                point.F_ADDRESS = this.grdMap.Rows[i]["F_ADDRESS"].ToString();
                point.CTL_NODE  = this.txtCtlNode.Text;
                //point.CTL_NODE  = this.grdMap.Rows[i]["CTL_NODE"].ToString();


                this.poe_list.Add(point);

            }
            /*
            foreach (POE_Node point in this.poe_list as ObservableCollection<POE_Node>)
            {
                point.POINT_NM = this.grdMap.Rows[i + 1]["POINT_NM"].ToString();
                point.VALID_YN = this.grdMap.Rows[i + 1]["VALID_YN"].ToString();
                point.POE_IP = this.grdMap.Rows[i + 1]["POE_IP"].ToString();
                point.POE_PORT = this.grdMap.Rows[i + 1]["POE_PORT"].ToString();
                point.G_ADDRESS = this.grdMap.Rows[i + 1]["G_ADDRESS"].ToString();
                point.F_ADDRESS = this.grdMap.Rows[i + 1]["F_ADDRESS"].ToString();
                //point.MODEL         = this.grdMap.Rows[i + 1]["MODEL"].ToString();
                point.CTL_NODE = this.grdMap.Rows[i + 1]["CTL_NODE"].ToString();

                i++;
            }
            */


            StreamWriter wr = new StreamWriter(strPointPath, false, Encoding.UTF8);
            XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<POE_Node>));
            xs.Serialize(wr, poe_list);

            bResult = true;
            wr.Flush();
            wr.Close();

            return bResult;


        }

        private bool LoadMidasInfo()
        {
            try
            {
                string strPointPath = AppDomain.CurrentDomain.BaseDirectory + string.Format(@"\Config\Point_List.xml");


                FileInfo fi = new FileInfo(strPointPath);

                if (fi.Exists == false) { fi.Create(); }
                XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<POE_Node>));
                using (StreamReader rd = new StreamReader(strPointPath))
                {
                    poe_list = xs.Deserialize(rd) as ObservableCollection<POE_Node>;
                }

                StreamWriter wr = new StreamWriter(strPointPath, false, Encoding.UTF8);
                xs.Serialize(wr, poe_list);

                wr.Flush();
                wr.Close();


                fi = null;

            }
            catch (Exception es)
            {
                return false;
            }
            return true;
        }
        #endregion Point ������ �����Ѵ�.





        #region ��Ž��� �Լ�
        private void StartComm()
        {
            try
            {
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                btnExit.Enabled = false;
                btnNodeSave.Enabled = false;
                //SaveCommData[] savecomm = new SaveCommData[];


                CommonData.boProgramRun = true;
                tdMain = new Thread(new ThreadStart(THREAD_COMM_START));
                tdMain.Start();


                tdCount = new Thread(new ThreadStart(ThreadCountDisplay));
                tdCount.Start();


                return; // Thread Start 
            }
            catch { }
        }
        #endregion ��Ž��� �Լ�


        #region �����忡�� �󺧿� �� ����
        private delegate void SetTextBoxCallback(Label lbl, String str);
        private void SetTextBox(Label lbl, String str)
        {
            try
            {
                if (lbl.InvokeRequired)
                {
                    SetTextBoxCallback setTextBoxCallback = new SetTextBoxCallback(SetTextBox);
                    this.Invoke(setTextBoxCallback, new object[] { lbl, str });
                }
                else
                {
                    lbl.Text = str;
                }
            }
            catch { }
        }
        #endregion �����忡�� �󺧿� �� ����


        #region ��ŷ���
        Object[] objArray = null;
        // One event is used for each Fibonacci object.

        int iWorkth = 0;
        int iComP = 0;
        int iWorkerThread = 0;
        int iCompletionThread = 0;

        // ���α׷� ������ Thread 
        public void THREAD_COMM_START()
        {
            try
            {
                int SlipmilliSecond = 51;
                DateTime dtLogDeleteCheck = DateTime.Now;
                int iUploadCheck = 0;


                run_detectors = new ObservableCollection<POE_Node>(poe_list.Where(x => x.VALID_YN == "Y" || x.DETECTOR == sMulti_Channel_Detector));
                //run_detectors = new ObservableCollection<POE_Node>(poe_list);


                thCount = run_detectors.Count;
                // Array ��ü ������ŭ ����
                objArray = new Object[thCount];





                int k = 0;
                ObservableCollection<POE_Node> paramGTM = null;
                for (int i = 0; i < run_detectors.Count; i++)
                {
                    int iCh_Count = this.getChannelCount(run_detectors[i].DETECTOR);
                    int iCul_Channel = run_detectors[i].Channel;

                    if (iCul_Channel == 1)
                    {
                        paramGTM = new ObservableCollection<POE_Node>();
                    }


                    if (iCh_Count == iCul_Channel)
                    {

                        paramGTM.Add(run_detectors[i]);
                        objArray[k] = (AdaptorInf)Activator.CreateInstance(Type.GetType(string.Format("{0}.{1}", "IMCC_ALL_OF_POE_GET.Adaptors.Detector", run_detectors[i].DETECTOR)), this, paramGTM);
                        k++;

                    }
                    else
                    {
                        paramGTM.Add(run_detectors[i]);
                    }

                }

                // ���� ������ ��ü ����ŭ �������� �Ѵ�.
                Array.Resize(ref objArray, k);
                iUploadCheck = objArray.Length;
                thCount = iUploadCheck;

                // �׸��� ����ó��
                foreach (POE_Node node in poe_list.Where(x => x.VALID_YN != "Y"))
                {
                    GRID_WRITE_MON(node.SEQ, node.CTL_NODE.ToString(), 9); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER
                }

                swUpload.Start();


                int iwork = 0;
                int iIOCom = 0;
                ThreadPool.GetMaxThreads(out iwork, out iIOCom);
                // MaxThreads ������ �ý����� �� �ǵ��� ����!!!
                //ThreadPool.SetMaxThreads(150, iIOCom);
                ThreadPool.SetMinThreads(iUploadCheck / 3, iIOCom);

                while (CommonData.boProgramRun)
                {
                    Thread.Sleep(SlipmilliSecond);

                    for (int j = 0; j < objArray.Length; j++)
                    {
                        ThreadPool.QueueUserWorkItem(((AdaptorInf)objArray[j]).Process, j);
                    }


                    if (dtLogDeleteCheck < DateTime.Now.AddSeconds(-1 * 60 * 10))
                    {
                        WRITE_LOG_DELETE();
                        dtLogDeleteCheck = DateTime.Now;
                    }


                    _doneEvent.WaitOne();
                    thCount = iUploadCheck;
                    _doneEvent.Reset();

                    if (iUploadCheck > 0)
                    {
                        if (swUpload.ElapsedMilliseconds >= 2000 * 60)
                        {
                            string sCtl_Node = run_detectors[0].CTL_NODE;
                            ((AdaptorInf)objArray[0]).DB_UPLOAD_CTL(sCtl_Node);
                            swUpload.Reset();
                            swUpload.Start();
                            this.WRITE("F", 0, sCtl_Node, "0", "0", string.Format(" SP_CTL_PROCESS '{0}' Uploaded...", sCtl_Node));
                        }
                    }

                    Thread.Sleep(SlipmilliSecond);

                } // Thread While

                // grid color reset ;
                foreach (POE_Node node in poe_list)
                {
                    GRID_WRITE_MON(node.SEQ, node.CTL_NODE, 0); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER
                }

                // ������ �迭 ��ü ���� 
                objArray = null;
            }
            catch { }


        }
        #endregion �����忡�� �󺧿� �� ����



        private void ThreadCountDisplay()
        {
            try
            {
                while (CommonData.boProgramRun)
                {
                    try
                    {
                        ThreadPool.GetMaxThreads(out iWorkerThread, out iCompletionThread);
                        SetTextBox(this.lblWorkerThread, "- �ִ� �۾��� ������� :  " + iWorkerThread.ToString());
                        SetTextBox(this.lblCompletitionThread, "- �ִ� �񵿱� I/O ������� :  " + iCompletionThread.ToString());

                        ThreadPool.GetAvailableThreads(out iWorkth, out iComP);

                        SetTextBox(this.lblMax, "- ��밡���� �۾������� ��: " + iWorkth.ToString());
                        SetTextBox(this.lblAvailable, "- ��밡���� �񵿱� I/O ������ �� :  " + iComP.ToString());

                        SetTextBox(this.lblUsedThread, "- ������� ������ �� :  " + (iWorkerThread - iWorkth).ToString());

                        Thread.Sleep(2000);

                        if (bStopCount) break;
                    }
                    catch (Exception ex)
                    {
                        SetTextBox(this.lblWorkerThread, "�����߻� :  " + ex.Message);
                    }
                }
            }
            catch { }

        }


        public void StopComm()
        {
            CommonData.boProgramRun = false;
            Thread.Sleep(1000);
        }


        public bool WRITE(string paramDirection, int paramSeq, string paramCtlNode, string paramPointNm, string paramType, string paramMessage)
        {
            bool boRtn = false;
            try
            {
                

                switch (paramDirection)
                {
                    //case "D": WRITE_DISPLAY(paramSeq, paramCtlNode, paramPointNm, paramType, paramMessage); break;
                    case "F": WRITE_LOG(1, paramCtlNode, paramPointNm, paramType, paramMessage); break;
                    default:
                        //WRITE_DISPLAY(paramSeq, paramCtlNode, paramPointNm, paramType, paramMessage);
                        WRITE_LOG(1, paramCtlNode, paramPointNm, paramType, paramMessage);
                        break;
                }
                boRtn = true;


            }
            catch { }
            return boRtn;
        }


        public delegate void delegateWRITE_ModbusDISPLAY(string sText);
        public void WRITE_ModbusDISPLAY(string sText)
        {

            try
            {
                if (!this.txtImsi.InvokeRequired)
                {

                    this.txtImsi.Text = sText;
                }
                else
                {
                    delegateWRITE_ModbusDISPLAY _delegate = new delegateWRITE_ModbusDISPLAY(WRITE_ModbusDISPLAY);
                    this.Invoke(_delegate, new object[] { sText });
                }
            }
            catch
            {

            }


        } //public bool DISPLAY_WRITE(string paramFIN_NODE, string paramType, string paramMessage)


        public delegate bool delegateWRITE_DISPLAY(int paramSeq, string paramCtlNode, string paramPointNm, string paramType, string paramMessage);
        public bool WRITE_DISPLAY(int paramSeq, string paramCtlNode, string paramPointNm, string paramType, string paramMessage)
        {
            bool boRtn = false;

            paramType = paramType.ToUpper();

            try
            {
                /*
                if (!_lv.InvokeRequired)
                {
                    if (_lv.Items.Count > 200)
                    {
                        _lv.Items.RemoveAt(200);
                    }

                    li = new ListViewItem(string.Format("{0:HH:mm:ss} [{1:D2}-{2}-{3}][{4}] {5}", DateTime.Now, paramSeq, paramCtlNode, paramPointNm, paramType, paramMessage));

                    if (paramType == "ERROR")
                    {
                        li.BackColor = Color.Red;
                        li.ForeColor = Color.Yellow;
                    }
                    else if (paramType == "SQL")
                    {
                        li.BackColor = Color.WhiteSmoke;
                        li.ForeColor = Color.Blue;
                    }
                    else if (paramType == "PLC")
                    {
                        li.BackColor = Color.WhiteSmoke;
                        li.ForeColor = Color.DarkGreen;
                    }
                    else if (paramType == "SCAN")
                    {
                        li.BackColor = Color.WhiteSmoke;
                        li.ForeColor = Color.DarkGreen;
                    }
                    else if (paramType == "INFO")
                    {
                        li.BackColor = Color.DarkGray;
                        li.ForeColor = Color.WhiteSmoke;
                    }
                    else if (paramType == "SYSTEM")
                    {
                        li.BackColor = Color.White;
                        li.ForeColor = Color.DarkGray;
                    }
                    else
                    {
                        li.BackColor = Color.LightGray;
                        li.ForeColor = Color.Black;
                    }

                    _lv.Items.Insert(0, li);

                    //WRITE_LOG(paramCommType, paramFIN_NODE, paramType, paramMessage);
                }
                else
                {
                    delegateWRITE_DISPLAY _delegate = new delegateWRITE_DISPLAY(WRITE_DISPLAY);
                    this.Invoke(_delegate, new object[] { paramSeq, paramCtlNode, paramPointNm, paramType, paramMessage });
                }
                */
                boRtn = true;

            }
            catch
            {
                boRtn = false;
            }

            return boRtn;
        } //public bool DISPLAY_WRITE(string paramFIN_NODE, string paramType, string paramMessage)


        public bool WRITE_LOG(int paramSeq, string paramCtlNode, string paramPointNm, string paramType, string paramMessage)
        {
            bool boRtn = false;

            StreamWriter _SW;
            try
            {
                if (paramType == "RAW")
                {
                    _SW = new StreamWriter(currentPathLog + @"\" + string.Format("L{0:yyyyMMddHH}[{1:D3}-{2}-{3}]", DateTime.Now, paramSeq, paramCtlNode, paramPointNm) + ".log", true);
                }
                else
                {
                    _SW = new StreamWriter(currentPathLog + @"\" + string.Format("L{0:yyyyMMdd}[{1:D3}-{2}-{3}]", DateTime.Now, paramSeq, paramCtlNode, paramPointNm) + ".log", true);
                }

                _SW.WriteLine(string.Format("{0:HH:mm:ss} {1}", DateTime.Now, string.Format("{0} {1}",paramType, paramMessage)));
                _SW.Flush();
                _SW.Close();

                boRtn = true;
            }
            catch
            {
                boRtn = false;
            }

            _SW = null;

            return boRtn;
        } //public bool LOG_WRITE(string paramFIN_NODE, string paramType, string paramMessage)


        public void WRITE_LOG_DELETE()
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(currentPathLog);
                DateTime dtCheck = DateTime.Now.AddDays(iLogSaveDays * -1);

                foreach (FileInfo fi in di.GetFiles())
                {
                    if (fi.CreationTime < dtCheck)
                    {
                        try
                        {
                            fi.Delete();
                        }
                        catch (Exception es)
                        {
                            WRITE("", 999, "", "", "ERROR", es.Message.ToString());
                            continue;
                        }
                    }
                }

            }
            catch (Exception es)
            {
                WRITE("", 999, "", "", "ERROR", es.Message.ToString());
            }

        }

        public delegate bool delegateGRID_WRITE_MON(int paramSeq, string paramCtlNode, int paramMode);
        public bool GRID_WRITE_MON(int paramSeq, string paramCtlNode, int paramMode)
        {
            bool boRtn = false;

            try
            {
                if (!this.grdView.InvokeRequired)
                {
                    int rowIndex = -1;

                    Row drRow = grdView.Rows.Cast<Row>().Where(r =>
                        (

                            r["CTL_NODE"].ToString().Equals(paramCtlNode)
                            && r["SEQ"].ToString().Equals(paramSeq.ToString())
                        )
                                                               ).First();
                    rowIndex = drRow.Index;




                    var tempRows = from tempRow in CommonData.saveDt
                                   where tempRow.SEQ == paramSeq && tempRow.CTL_NODE == paramCtlNode
                                   select tempRow;

                    //tempRows.First().FAULT_YN = paramFaultYn;
                    //tempRows.First().FAULT_CODE = paramFaultCode;
                    //tempRows.First().ChangeDt = DateTime.Now;

                    float fValue = 0; //tempRows.First().LEAK_VAL;
                    string sFault = ""; //tempRows.First().FAULT_CODE;
                    try
                    {
                        fValue = tempRows.First().LEAK_VAL;
                        sFault = tempRows.First().FAULT_CODE;
                    }
                    catch { }


                    //int iColIndex = 0;
                    if (paramMode == 0)
                    {
                        // Stop 
                        grdView.Rows[rowIndex].StyleNew.BackColor = Color.White;
                        fValue = 0;
                        sFault = "";

                    }
                    else if (paramMode == 1)
                    {
                        // success
                        grdView.Rows[rowIndex].StyleNew.BackColor = Color.LightGreen;
                        //fValue = tempRows.First().LEAK_VAL;
                        //sFault = tempRows.First().FAULT_CODE;


                    }
                    else if (paramMode == 2)
                    {
                        // error
                        grdView.Rows[rowIndex].StyleNew.BackColor = Color.Red;
                        fValue = 0;
                        //sFault = "";

                    }
                    else if (paramMode == 3)
                    {
                        grdView.Rows[rowIndex].StyleNew.BackColor = Color.LightBlue;
                        fValue = 0;
                        sFault = "";

                    }
                    else
                    {
                        grdView.Rows[rowIndex].StyleNew.BackColor = Color.DarkGray;
                        fValue = 0;
                        sFault = "";

                    }


                    string sStyle = "";
                    int icolIndex = 0;

                    icolIndex = grdView.Cols.Count - 1;
                    grdView.Rows[rowIndex][icolIndex] = fValue;
                    sStyle = fValue > 0 ? "redStyle" : "blackStyle";
                    this.grdView.SetCellStyle(rowIndex, icolIndex, sStyle);

                    icolIndex = grdView.Cols.Count - 2;
                    grdView.Rows[rowIndex][icolIndex] = sFault;
                    if (sFault.Equals("0000"))
                    {
                        sStyle = "blackStyle";
                    }
                    else if (sFault.Equals("5003"))
                    {
                        sStyle = "yellowStyle";
                    }
                    else
                    {
                        sStyle = "redStyle";
                    }
                    this.grdView.SetCellStyle(rowIndex, icolIndex, sStyle);


                }
                else
                {
                    delegateGRID_WRITE_MON _delegate = new delegateGRID_WRITE_MON(GRID_WRITE_MON);
                    this.Invoke(_delegate, new object[] { paramSeq, paramCtlNode, paramMode });
                }

                boRtn = true;

            }
            catch (Exception ex)
            {

                boRtn = false;
            }

            return boRtn;

        } //public bool DISPLAY_WRITE(string paramFIN_NODE, string paramType, string paramMessage)







        private void MAIN_F_SizeChanged(object sender, EventArgs e)
        {
            try
            {
                //_lv.Columns[0].Width = _lv.Width - 20;
            }
            catch
            {
            }
        }

        public string TraceIP
        {
            get
            {
                return this.txtTraceIpAddress.Text;
            }
        }

        

        //string sClassName = "";




        private void btnPointAdd_Click(object sender, EventArgs e)
        {
            //this.grdMap.Rows[iSel].Selected = false;

            for (int i = 0; i < this.grdMap.Rows.Count; i++)
            {
                this.grdMap.Rows[i].Selected = false;
            }

            string sF_Add = "";
            int iF_Add = 0;
            string sG_Add = "";
            int iG_Add = 0;
            if (this.grdMap.Rows.Count > 1)
            {
                sF_Add = this.grdMap.Rows[this.grdMap.Rows.Count - 1]["F_ADDRESS"].ToString();
                iF_Add = CommonData.OnlyNumeric(sF_Add);

                sG_Add = this.grdMap.Rows[this.grdMap.Rows.Count - 1]["G_ADDRESS"].ToString();
                iG_Add = CommonData.OnlyNumeric(sG_Add);


            }
            else
            {
                iF_Add = 0;
                iG_Add = 0;
            }


            //this.grdMap.Rows[0]. .Styles["Normal"].Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.Dotted;
            CellStyle cs = grdMap.Styles.Add("red");
            cs.BackColor = Color.LightGray;

            string sDetector = this.cboDetector.Items[this.cboDetector.SelectedIndex].ToString();
            int pointCnt = GetPointPerDetector(); // ������� Addres ������ ��
            int iUseCnt = this.getChannelCount(sDetector); // ���� ��� ����Ʈ ��
            string sIP_Address = GetLastIP();

            if (sIP_Address.Equals(""))
            {
                MessageBox.Show("IP Address Full, you cannot Add Point anymore");
                return;
            }

            //int pointCnt = this.ChannelCount;
            for (int i = 0; i < pointCnt; i++)
            {
                this.grdMap.Rows.Add();

                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["SEQ"] = this.grdMap.Rows.Count -1 ;
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["CHANNEL"] = i+1;
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["CTL_NODE"] = this.txtCtlNode.Text;
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["POINT_NM"] = string.Format("{0}-{1:D3}", this.txtCtlNode.Text, this.grdMap.Rows.Count - 1);

                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["DETECTOR"] = cboDetector.Items[cboDetector.SelectedIndex].ToString();
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["VALID_YN"] = iUseCnt >= i+1 ? "Y" : "N";
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["POE_PORT"] = this.txtPort.Text;
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["POE_IP"] = sIP_Address;
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["F_ADDRESS"] = string.Format("F{0:D3}", ++iF_Add);
                this.grdMap.Rows[this.grdMap.Rows.Count - 1]["G_ADDRESS"] = string.Format("G{0:D3}", ++iG_Add);
                //this.grdMap.Height = this.grdMap.Rows[this.grdMap.Rows.Count - 1].Bottom + 2;
            }

            this.GridGroupSetting(grdMap);
            // assign bold style to a cell range

            this.grdMap.ShowCell(this.grdMap.Rows.Count - 1, 1);
        }

        private string GetLastIP()
        {
            string sRetIP = "";
            if(this.grdMap.Rows.Count == 1)
            {
                sRetIP = this.txtIPAddress.Text;
            }
            else if(this.grdMap.Rows.Count > 1)
            {
                sRetIP = this.grdMap.Rows[this.grdMap.Rows.Count - 2]["POE_IP"].ToString();
                sRetIP = Increase_IP(sRetIP);
            }
            return sRetIP;
        }

        private string Increase_IP(string curr_ip)
        {
            string[] sIP = curr_ip.Split('.');
            int retIP;
            retIP = int.Parse(sIP[3]) + 1;

            string sCheckIP = string.Format("{0}.{1}.{2}.{3}", sIP[0], sIP[1], sIP[2], retIP.ToString());
            IPAddress ip;
            if (IPAddress.TryParse(sCheckIP,out ip) == false)
            {
                sCheckIP = "";
            }
            return sCheckIP;
        }

        private string Increase_IP(string curr_ip,int idx)
        {
            string[] sIP = curr_ip.Split('.');
            int retIP;
            retIP = int.Parse(sIP[3]) + idx;

            string sCheckIP = string.Format("{0}.{1}.{2}.{3}", sIP[0], sIP[1], sIP[2], retIP.ToString());
            IPAddress ip;
            if (IPAddress.TryParse(sCheckIP, out ip) == false)
            {
                sCheckIP = "";
            }
            return sCheckIP;

        }


        private void GridGroupSetting(C1FlexGrid grd)
        {

            CellStyle csgrdMapBackLightGray = grd.Styles.Add("Back_LightGray");
            csgrdMapBackLightGray.BackColor = Color.LightGray;

            CellStyle csgrdMapBackWhite = grd.Styles.Add("Back_White");
            csgrdMapBackWhite.BackColor = Color.White;

            int iKColorSperate = 0;
            for (int i = 1; i < grd.Rows.Count; i++)
            {
                int iCh_count = GetPointPerDetector();
                int icul_channel = int.Parse(grd.Cols["CHANNEL"][i].ToString());
                if (icul_channel == 1 && grd.Equals(grdMap))
                {
                    CellRange rg = grd.GetCellRange(i, 1, i + iCh_count - 1, grd.Cols.Count - 1);

                    if (iKColorSperate % 2 == 0 && icul_channel == 1)
                    {
                        rg.Style = grd.Styles["Back_LightGray"];
                    }
                    else
                    {
                        rg.Style = grd.Styles["Back_White"];
                    }
                    iKColorSperate++;
                }

                if (icul_channel ==1 &&  iCh_count > 1 && grd.Equals(grdView))
                {
                    grd.AllowMerging = AllowMergingEnum.Custom;
                    //CellRange crCellRange = grd.GetCellRange(i - (iCh_count - 1), grd.Cols["POE_IP"].Index, i , grd.Cols["POE_IP"].Index);
                    CellRange crCellRangeIP = grd.GetCellRange(i , grd.Cols["POE_IP"].Index, i + iCh_count - 1, grd.Cols["POE_IP"].Index);
                    grd.MergedRanges.Add(crCellRangeIP);

                    CellRange crCellRangePort = grd.GetCellRange(i , grd.Cols["POE_PORT"].Index, i + iCh_count - 1, grd.Cols["POE_PORT"].Index);
                    grd.MergedRanges.Add(crCellRangePort);

                    CellRange crCellRangeDetector = grd.GetCellRange(i , grd.Cols["DETECTOR"].Index, i + iCh_count - 1, grd.Cols["DETECTOR"].Index);
                    grd.MergedRanges.Add(crCellRangeDetector);
                }
            }
        }


        private int getChannelCount(string sDetectorName)
        {
            // ������ ��� ������ ä���� 4���� �����Ѵ�.
            return sDetectorName.Equals(sMulti_Channel_Detector) ? this.multi_Channel_Count : 1;
        }
        private int GetPointPerDetector()
        {
            return 4;
        }

        private void GridGroupBoldSetting(C1FlexGrid grd)
        {

            CellRange rg = grd.GetCellRange(1, 1, 4, grd.Cols.Count - 1);
            rg.Style = grd.Styles["bold"];

            
            /*
            for (int i = 1; i < grd.Rows.Count; i++)
            {
                grd.Rows[i]["SEQ"] = i.ToString();

                if (i == this.ChannelCount)
                {
                    CellRange rg = grd.GetCellRange(i, 1, 4, grd.Cols.Count - 1);

                    if (((i - 1) / this.ChannelCount) % 2 == 0)
                    {
                        rg.Style = grd.Styles["bold"];
                    }
                }
            }
            */
        }


        private void btnPointDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.grdMap.Rows.Selected != null)
                {
                    iSel = grdMap.RowSel;
                    //if (grdMap.Rows.Count-1 == this.iSel) this.iSel =this.iSel-1;
                    int iCh_Count = this.GetPointPerDetector();
                    int icur_ch = int.Parse(grdMap.Cols["CHANNEL"][iSel].ToString());
                    int irow_idx = 0;
                    /*
                    if (iCh_Count == 1)
                    {
                        int ipcnt = this.ChannelCount;
                        int icul = (this.iSel - 1) / ipcnt;
                        int i = icul * ipcnt;
                        Row ro = grdMap.Rows[i + 1];
                        irow_idx = ro.Index;
                    }
                    else
                    {
                        Row ro = grdMap.Rows[iSel];
                        irow_idx = ro.Index - (icur_ch - 1);
                        
                    }
                    */
                    Row ro = grdMap.Rows[iSel];
                    irow_idx = ro.Index - (icur_ch - 1);

                    this.grdMap.Rows.RemoveRange(irow_idx, iCh_Count);


                }
            }
            catch { }

            this.GridGroupSetting(grdMap);
        }



        private void GridSequenceSetting()
        {
            for (int i = 0; i < this.grdView.Rows.Count; i++)
            {
                this.grdView.Rows[i + 1]["SEQ"] = (i + 1).ToString();
            }
        }

        int iSel = 1;
        private void grdMap_Click(object sender, EventArgs e)
        {
            iSel = ((C1FlexGrid)sender).RowSel;
        }

        private void btnMapInitilized_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtCtlNode.Text))
            {
                MessageBox.Show("Input Control node.");
                txtCtlNode.Focus();
                return;
            }

            if (string.IsNullOrEmpty(this.txtIPAddress.Text))
            {
                MessageBox.Show("Input Midas Ip Address.");
                txtIPAddress.Focus();
                return;
            }

            if (string.IsNullOrEmpty(this.txtPort.Text))
            {
                MessageBox.Show("Input Midas Port(numeric).");
                txtPort.Focus();
                return;
            }

            try
            {
                int.Parse(this.txtPort.Text);
            }
            catch
            {
                MessageBox.Show("POE Port is numeric only).");
                txtPort.Focus();
                return;

            }


            if (string.IsNullOrEmpty(this.txtCount.Text))
            {
                MessageBox.Show("Input Midas Creation Count(numeric).");
                txtCount.Focus();
                return;
            }


            try
            {
                int.Parse(this.txtCount.Text);
            }
            catch
            {
                MessageBox.Show("Midas Creation Count is numeric only).");
                txtCount.Focus();
                return;

            }


            try
            {
                if(int.Parse(this.txtCount.Text)%4 != 0)
                {
                    MessageBox.Show("Only multiples of 4 are allowed.");
                    txtCount.Focus();
                    return;
                }

            }
            catch
            {
                MessageBox.Show("Midas Creation Count is numeric only).");
                txtCount.Focus();
                return;

            }


            if (MessageBox.Show("All of Midas List will be initilized Continue?", "Confirm", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                this.grdMap.Rows.RemoveRange(1, this.grdMap.Rows.Count - 1);

                Map_Initilized();

                this.POE_NodeDataInit(this.grdMap);
                this.GridGroupSetting(this.grdMap);
            }

        }
        



        private bool Map_Initilized()
        {
            bool bResult = false;
            

            try
            {
                int iCount = int.Parse(this.txtCount.Text);

                
                string strPointPath = AppDomain.CurrentDomain.BaseDirectory + @"\Config";
                if (Directory.Exists(strPointPath) == false)
                {
                    Directory.CreateDirectory(strPointPath);
                }

                strPointPath = strPointPath + string.Format(@"\Point_List.xml");

                ObservableCollection<POE_Node> midasList = new ObservableCollection<POE_Node>();

                string sDetector = this.cboDetector.Items[this.cboDetector.SelectedIndex].ToString();
                int iUseCnt = this.getChannelCount(sDetector); // ���� ��� ����Ʈ ��
                int pointPerDetector = this.GetPointPerDetector();

                POE_Node node;
                string ipAddress = "";
                for (int i = 0; i < iCount ; i++)
                {
                    node = new POE_Node();
                    node.SEQ = (i + 1);
                    node.Channel = (i + 1) % this.GetPointPerDetector() == 0 ? 4 : (i + 1) % this.GetPointPerDetector();
                    //node.Channel = this.ChannelCount;
                    node.POINT_NM = string.Format("{0}-{1:D3}", this.txtCtlNode.Text, i + 1);
                    // �������� 4ä���� ���� Y, 1ä���� ù��°���� Y �� ���� 
                    node.VALID_YN = iUseCnt >= node.Channel ? "Y" : "N";

                    // 2019.07.29 == ��û���� �ݿ� ���� ������ 'N'
                    // node.VALID_YN = "N";
                    node.DETECTOR = cboDetector.Items[cboDetector.SelectedIndex].ToString();
                    ipAddress = Increase_IP(this.txtIPAddress.Text, i / this.GetPointPerDetector());
                    node.POE_IP = ipAddress;
                    //int increment = 0;
                    node.POE_PORT = this.txtPort.Text;
                    node.G_ADDRESS = string.Format("G{1:D3}", this.txtCtlNode.Text, (i + 1));
                    node.F_ADDRESS = string.Format("F{1:D3}", this.txtCtlNode.Text, (i + 1));
                    node.CTL_NODE = this.txtCtlNode.Text;
                    //node.MODEL = "MIDAS";
                    midasList.Add(node);
                }

                StreamWriter wr = new StreamWriter(strPointPath, false, Encoding.UTF8);
                XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<POE_Node>));
                xs.Serialize(wr, midasList);

                wr.Flush();
                wr.Close();
                bResult = true;

            }
            catch (Exception ex)
            {

            }
            return bResult;
        }

        private void mainNodeTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.mainNodeTab.SelectedIndex == 0)
            {
                this.POE_NodeDataInit(this.grdView);
                this.GridGroupSetting(this.grdView);
            }
            else if (this.mainNodeTab.SelectedIndex == 1)
            {
                
                this.POE_NodeDataInit(this.grdMap);
                this.GridGroupSetting(this.grdMap);

            }
        }

        bool bStopCount = false;
        private void MAIN_F_FormClosing(object sender, FormClosingEventArgs e)
        {
            bStopCount = true;
            this.WRITE("F", 0, this.txtCtlNode.Text, "0", "0", string.Format(" Application Exit..."));
        }

        private void cboDetector_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sDetectorName = ((ComboBox)sender).Items[((ComboBox)sender).SelectedIndex].ToString();
            
        }

        private void btnUpDown_Click(object sender, EventArgs e)
        {
            try
            {
                int iadddel = 1;

                if (this.grdMap.Rows.Selected != null)
                {
                    iSel = grdMap.RowSel;
                    //if (grdMap.Rows.Count-1 == this.iSel) this.iSel =this.iSel-1;
                    int iCh_Count = this.GetPointPerDetector();
                    int iMv_Ch_Count = 0; //this.getChannelCount(grdMap.Cols["DETECTOR"][iSel].ToString());
                    int icur_ch = int.Parse(grdMap.Cols["CHANNEL"][iSel].ToString());




                    if ((Button)sender == this.btnUp)
                    {
                        // �� ����Row ���� Up �Ұ��
                        if (iSel == 1) return;
                        iadddel = iadddel * -1;
                        iMv_Ch_Count = icur_ch * iadddel;
                    }
                    else
                    {
                        // �� �Ʒ��� Row ���� Down �Ұ��
                        if (iSel == grdMap.Rows.Count) return;

                        iMv_Ch_Count = iCh_Count - icur_ch + 1;
                    }
                    iMv_Ch_Count = this.GetPointPerDetector();



                    int irow_idx = 0;
                    /*
                    if (iCh_Count == 1)
                    {
                        int ipcnt = this.ChannelCount;
                        int icul = (this.iSel - 1) / ipcnt;
                        int i = icul * ipcnt;
                        Row ro = grdMap.Rows[i + 1];
                        irow_idx = ro.Index;
                    }
                    else
                    {
                        Row ro = grdMap.Rows[iSel];
                        irow_idx = ro.Index - (icur_ch - 1);

                    }
                    */
                    Row ro = grdMap.Rows[iSel];
                    irow_idx = ro.Index - (icur_ch - 1);

                    if (irow_idx + iMv_Ch_Count * iadddel == 0) return;
                    this.grdMap.Rows.MoveRange( irow_idx, iCh_Count, irow_idx + iMv_Ch_Count * iadddel);
                }
            }
            catch { }

            this.GridGroupSetting(grdMap);
        }



        private void grd_AfterEdit(object sender, RowColEventArgs e)
        {
            int iRow = e.Row;
            int icol = e.Col;
            try
            {
                if (grdMap.Cols[icol].Name.Equals("POE_IP") || grdMap.Cols[icol].Name.Equals("POE_PORT") || grdMap.Cols[icol].Name.Equals("DETECTOR"))
                //if (grdMap.Cols[icol].Name.Equals("DETECTOR"))
                {
                    string sValue = grdMap.Cols[icol][iRow].ToString();
                    int iCh_count = GetPointPerDetector();
                    int iUse_Channel = this.getChannelCount(grdMap.Cols["DETECTOR"][iRow].ToString());
                    int icul_channel = int.Parse(grdMap.Cols["CHANNEL"][iRow].ToString());

                    if (iCh_count > 1)
                    {
                        for (int i = 1; i <= iCh_count; i++)
                        {
                            grdMap.Cols[icol][iRow - (icul_channel) + i] = sValue;

                            if (grdMap.Cols[icol].Name.Equals("DETECTOR"))
                            {
                                grdMap.Cols["VALID_YN"][iRow - (icul_channel) + i] = i <= iUse_Channel ? "Y" : "N";
                            }
                        }
                    }
                }
            }
            catch { }
        }

    } // Class



} // NameSpace