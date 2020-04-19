using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Sockets;
using IMCC_MGMService.Com;
using IMCC_MGMService.Model;

using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;
using System.Configuration;
using System.Diagnostics;
//using IMCC_LEDLAMP_Adapter;

namespace IMCC_MGMService.WCF
{
    public static class COSMOS_Protocol
    {

        #region 전역변수
        // 통신용 Main 쓰레드 선언
        static Thread comThreadMain;

        // 로컬 디비 저장용 
        static Thread thrLocalLeak;

        // 로컬 디비 삭제용
        static Thread thrLocalDBDelete;

        // 외부 object 변수
        public static ControlNode ctlNode;
        // 포인트 정보
        public static ObservableCollection<PointInfo> pointInf;
        // 디비 연결정보
        public static DBHelper dbexe = null; // new DBHelper(ctlNode);

        // Etos 연결 소켓 선언
        static Socket socketLoop;
        // 통신 지속 여부 확인 변수
        public static bool isCommRun = true;

        // Read 메모리 개수
        public static int iReadAddSize = 25; // 64point 씩처리
        public static int iLocalDB_Saved_Days = 365;

        // SMCS DataBase 통신 상태
        public static bool bSMCS_DB_Status = false;

        // MiddleWare (Etos 장비) 통신상태
        public static bool bEtos_COMM_Status = false;

        #endregion 전역변수


        #region 통신 시작 부분
        public static void Main_Communication_Working()
        {
            // control 정보 로드
            COSMOS_Protocol.LoadControlNode();

            // 노드 정보로 디비 생성
            dbexe = new DBHelper(ctlNode);


            // 포인트정보 로드
            if (LoadPointInfo() == false) return;


            // 포인트 정보 초기화
            if (PointsInitilize() == false) return;

            // LED 램프 초기화
            LedLampInitialize();

            ////////////////////////////////////////
            ////////// 메인 Thread 호출 /////////////
            ////////////////////////////////////////
            Communication_Start();

        }
        #endregion 통신 시작 부분

        #region Node 정보를 불러온다.
        /// <summary>
        /// control 노드 정보를 읽어온다. ( from xml )
        /// </summary>
        public static bool LoadControlNode()
        {


            try
            {
                string strNodePath = AppDomain.CurrentDomain.BaseDirectory + @"\Config\CTL_NODE.xml";
                FileInfo fi = new FileInfo(strNodePath);

                if (fi.Exists)
                {
                    XmlSerializer xs = new XmlSerializer(typeof(ControlNode));

                    using (StreamReader rd = new StreamReader(strNodePath))
                    {
                        ctlNode = xs.Deserialize(rd) as ControlNode;
                    }
                    COMLogger.SetIMCCLog("Info", "Control Node Information successfully loaded");
                }
                fi = null;

            }
            catch (Exception es)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LoadControlNode = {0}", es.Message));
                return false;
            }
            return true;
        }
        #endregion

        #region Point 정보를 불러온다.
        /// <summary>
        /// Point  정보를 읽어온다. ( from xml )
        /// </summary>
        public static bool LoadPointInfo()
        {

            try
            {
                string strPointPath = AppDomain.CurrentDomain.BaseDirectory + @"\Config\MGM_Map.xml";
                FileInfo fi = new FileInfo(strPointPath);

                if (fi.Exists)
                {
                    XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<PointInfo>));

                    using (StreamReader rd = new StreamReader(strPointPath))
                    {
                        pointInf = xs.Deserialize(rd) as ObservableCollection<PointInfo>;
                    }
                }


                fi = null;

            }
            catch (Exception es)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LoadPointInfo = {0}", es.Message));
                return false;
            }
            return true;
        }
        private static bool PointsInitilize()
        {
            try
            {
                // 불러온 xml 파일에서 측정값 정보는 초기화 한다.
                for (int i = 0; i < pointInf.Count; i++)
                {
                    pointInf[i].F_Address = string.Format("{0}.F{1:D3}", ctlNode.CTLNode, pointInf[i].SEQ);
                    pointInf[i].G_Address = string.Format("{0}.G{1:D3}", ctlNode.CTLNode, pointInf[i].SEQ);
                    pointInf[i].CurFaultCode = "2000";
                    pointInf[i].PreFaultCode = "2000";
                    pointInf[i].CurLeakVal = 0.0M;
                    pointInf[i].PreLeakVal = 999.999M;
                    pointInf[i].PreFaultYN = "X";
                    pointInf[i].CurFaultYN = "X";
                    pointInf[i].Alarm1 = false;
                    pointInf[i].Alarm2 = false;

                    //pointInf[i].PointName = "COS_" + pointInf[i].CardNo + "_" + pointInf[i].ComNo + "_" + string.Format("{0:D3}",pointInf[i].GrpSeq);

                }
                COMLogger.SetIMCCLog("Info", string.Format("PointsInitilize = Map files initilized successfully"));
                return true;
            }
            catch(Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("PointsInitilize = {0}", ex.Message));
                return false;
            }
        }

        #endregion

        #region =========node, point 정보 저장===============
        #region node 정보를 저장한다
        public static bool SetControlNodeInfo(ControlNode CtlData)
        {
            string strNodePath = AppDomain.CurrentDomain.BaseDirectory + @"\Config\CTL_NODE.xml";

            StreamWriter wr = new StreamWriter(strNodePath, false, Encoding.UTF8);
            XmlSerializer xs = new XmlSerializer(typeof(ControlNode));
            ctlNode = CtlData;
            xs.Serialize(wr, ctlNode);

            wr.Flush();
            wr.Close();

            return true;
        }
        #endregion node 정보를 저장한다

        #region Point 정보를 저장한다.
        public static bool SetPointsInfo(ObservableCollection<PointInfo> PointsData)
        {
            string strPointPath = AppDomain.CurrentDomain.BaseDirectory + @"\Config\MGM_Map.xml";


            StreamWriter wr = new StreamWriter(strPointPath, false, Encoding.UTF8);
            XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<PointInfo>));
            pointInf = PointsData;
            xs.Serialize(wr, pointInf);


            wr.Flush();
            wr.Close();

            return true;
        }
        #endregion Point 정보를 저장한다.
        #endregion =========node, point 정보 저장===============

        #region 메인쓰레드 시작 함수
        private static void Communication_Start()
        {
            try
            {

                // main 프로세스 쓰레드 생성
                if (comThreadMain != null)
                {
                    comThreadMain.Abort();
                    comThreadMain = null;
                }
                comThreadMain = new Thread(new ThreadStart(Process));
                //comThreadMain.IsBackground = true;
                comThreadMain.Start();

            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("process : {0} ", ex.Message));
                throw ex;
            }


            try
            {

                // main 프로세스 쓰레드 생성
                if (thrLocalLeak != null)
                {
                    thrLocalLeak.Abort();
                    thrLocalLeak = null;
                }
                // leak value 로컬디비 저장용
                thrLocalLeak = new Thread(new ThreadStart(LocalDBSaveManager));
                thrLocalLeak.IsBackground = true;
                thrLocalLeak.Start();
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LocalDBSaveManager : {0} ", ex.Message));
                throw ex;
            }

            try
            {
                // leak value 로컬디비 저장용
                thrLocalDBDelete = new Thread(new ThreadStart(LocalDBDeleteManager));
                thrLocalDBDelete.IsBackground = true;
                thrLocalDBDelete.Start();
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LocalDBDeleteManager : {0} ", ex.Message));
                throw ex;
            }

        }
        #endregion 


        private static void Process()
        {

            // 시작 주소 
            byte[] nAddrByte = new byte[2];
            byte[] nRegCnt = new byte[2];
            byte[] txByte = new byte[12]; // tx 공통

            // 읽어올 워드 개수 (1word =    2byte
            byte[] rxLeak = new byte[9 + iReadAddSize * 5 * 2];

            byte[] wTByte = new byte[14];
            byte[] wRByte = new byte[12];

            int iWhileCnt = pointInf.Count();

            int iLoop = 0;
            int iSign = 1;
            //string sFaultYNCheck = "N";
            string sFaultCodeCheck = "2000";
            // Etos 통신 성공 횟수
            int intSuccessCount = 0;
            /*
            lemp[0] = false;
            lemp[1] = false;
            lemp[2] = false;
            */

            // 감지기 Alarm발생여부 변수 1, 2
            bool bAlarmOut1 = false;
            bool bAlarmOut2 = false;

            // Fault LED 상태
            bool b_F_ONOFF = false;
            // Alarm LED 상태
            bool b_A_ONOFF = false;


            // 통신 정상 시작
            LedLampOnOff("N", true);
            





            SocketError socError = new SocketError();

            // 최초 소켓 open
            Socket_Connection(ctlNode.LIpAddress, ctlNode.ComPort);

            // SMCS upload 1분
            Stopwatch swUpload = new Stopwatch();
            swUpload.Start();


            // 통신 소켓 설정 연결
            while (isCommRun)
            {
                try
                {

                    // 통신 정상 시작
                    LedLampOnOff("N", true);


                    // Looping count 초기화
                    iLoop = 0;

                    if (isCommRun == false)
                    {
                        LedLampInitialize();
                        break;
                    }

                    // Led 램프 출력용 변수 초기화
                    bAlarmOut1 = false;
                    bAlarmOut2 = false;

                    // LED 램프 상태
                    b_F_ONOFF = false;
                    b_A_ONOFF = false;



                    // 1word 씩 (2byte )
                    while (iWhileCnt > iLoop)
                    {

                        if (isCommRun == false)
                        {
                            LedLampInitialize();
                            break;
                        }


                        // SMCS 데이터베이스 연결 상태 확인
                        bSMCS_DB_Status = JUST_DB_Check();

                        //===============Etos Tx 데이터 준비===============
                        #region 데이터 요청 쿼리 부분
                        // 사용 배열 초기화
                        Array.Clear(nAddrByte, 0, nAddrByte.Length);
                        Array.Clear(nRegCnt, 0, nRegCnt.Length);
                        Array.Clear(txByte, 0, txByte.Length);
                        Array.Clear(rxLeak, 0, rxLeak.Length);


                        // LEAK START  ANALAYZER 1
                        txByte[0] = 0x00;
                        txByte[1] = 0x00;

                        txByte[2] = 0x00;
                        txByte[3] = 0x00;

                        txByte[4] = 0x00;
                        txByte[5] = 0x06; // DataLength ( 고정 )

                        txByte[6] = 0x01;
                        txByte[7] = 0x03; // Function Code Holding Register


                        //txByte[8] ~ txByte[9]  Start Address
                        UInt16 uStartAddress = (UInt16)(Convert.ToInt16(pointInf[iLoop].PointAddress));

                        COMLogger.SetIMCCLog("Info", string.Format("TX Send Start Address : {0} ", uStartAddress));

                        nAddrByte = BitConverter.GetBytes(uStartAddress);
                        Array.Reverse(nAddrByte);
                        Array.Copy(nAddrByte, 0, txByte, 8, 2);

                        //txByte[11] = 0x01; // 읽을 Register Count
                        nRegCnt = BitConverter.GetBytes(Convert.ToInt16(iReadAddSize * 5)); // 읽어올 Regist Count
                        Array.Reverse(nRegCnt);
                        Array.Copy(nRegCnt, 0, txByte, 10, 2);

                        #endregion 데이터 요청 쿼리 부분



                        // Socket 이 연결 되지 않았으면 연결한다.
                        #region 소켓 확인
                        if (socketLoop.Connected == false)
                        {
                            // 통신소켓 연결 설정
                            if (Socket_Connection(ctlNode.LIpAddress, ctlNode.ComPort) == false)
                            {
                                // 소켓 연결이 안되었을경우 다시 연결한다.
                                LedLampInitialize();
                                goto SocketError;
                            }
                        }
                        #endregion 소켓 확인

                        #region Etos 데이터 송 수신
                        try
                        {
                            socketLoop.Send(txByte, 0, txByte.Length, SocketFlags.None);
                        }
                        catch (SocketException ex)
                        {
                            COMLogger.SetIMCCLog("Error", string.Format("송신데이터 소켓 오류 : {0}", ex.Message));
                            LedLampInitialize();
                            goto SocketError;

                        }
                        catch (Exception ex)
                        {
                            COMLogger.SetIMCCLog("Error", string.Format("Send Error : {0}", ex.Message));
                            continue;
                        }

                        Thread.Sleep(200);

                        try
                        {
                            socketLoop.Receive(rxLeak, 0, rxLeak.Length, SocketFlags.None, out socError);
                            COMLogger.SetIMCCLog("Info", string.Format("RX Receive Completed..."));

                        }
                        catch (SocketException ex)
                        {
                            COMLogger.SetIMCCLog("Error", string.Format("수신데이터 오류 : {0}", ex.Message));
                            LedLampInitialize();
                            goto SocketError;
                        }
                        catch (Exception ex)
                        {
                            COMLogger.SetIMCCLog("Error", string.Format("Receive Error : {0}", ex.Message));
                            Thread.Sleep(50);
                            continue;
                        }
                        Thread.Sleep(50);
                        #endregion  Etos 데이터 송 수신


                        #region 읽어온 메모리 처리 
                        //================= 메모리 점프 5개씩(5word) = 2 * 5 byte ===================
                        for (int j = 0, m = 0; j < iReadAddSize; j++, m = m + 10)
                        {
                            sFaultCodeCheck = "";

                            if (isCommRun == false)
                            {
                                LedLampInitialize();
                                break;
                            }

                            // Point 갯수와 Looping + 메모리 Read count 와 같으면 중단 
                            if (pointInf.Count == iLoop + j) break;


                            /****************** 사용여부 읽기 처리 시작 *********************/
                            #region USE YN 사용자 처리 확인 하여 Etos 메모리에 쓰기 
                            // 바이트 순번 농도(%): 15  ~  16
                            // 워드 단위로 쓰여 있는것이라 바이트 Reverse 를 한다.
                            string UseYN = Convert.ToString(rxLeak[m + 16], 2).PadLeft(8, '0') + Convert.ToString(rxLeak[m + 15], 2).PadLeft(8, '0');
                            //"0000000000000001" ==> 사용 , "0000000000000000" ==> 미사용
                            bool eTosWrite = UseYN.Substring(15, 1).Equals("1");
                            bool vIewWirte = pointInf[iLoop + j].UseYN.Equals("Y");
                            if (eTosWrite.Equals(vIewWirte) == false)
                            {

                                COMLogger.SetIMCCLog("Info", string.Format(" SEQ{0}, Card : {1}, Com Port : {2}, GroupSeq{3}, Write Vale {4} 사용여부 차이로 데이터 쓰기 >> ", pointInf[iLoop + j].SEQ, pointInf[iLoop + j].CardNo, pointInf[iLoop + j].ComNo, pointInf[iLoop + j].GrpSeq, vIewWirte));

                                /*****************Modbus Write*************************/
                                // 사용 배열 초기화
                                Array.Clear(nAddrByte, 0, nAddrByte.Length);
                                Array.Clear(nRegCnt, 0, nRegCnt.Length);
                                Array.Clear(txByte, 0, txByte.Length);


                                // LEAK START  ANALAYZER 1
                                txByte[0] = 0x00;
                                txByte[1] = 0x00;
                                txByte[2] = 0x00;
                                txByte[3] = 0x00;
                                txByte[4] = 0x00;
                                txByte[5] = 0x06; // DataLength ( 고정 )
                                txByte[6] = 0x01;
                                txByte[7] = 0x06; // Function Code Holding Register Writer

                                // txByte[8] ~ txByte[9]  Start Address
                                nAddrByte = BitConverter.GetBytes(Convert.ToInt16(pointInf[iLoop + j].PointAddress + 3));
                                Array.Reverse(nAddrByte);
                                Array.Copy(nAddrByte, 0, txByte, 8, 2);

                                // 사용 여부가 다를경우 화면상에 보이는 xml 기준 데이터로 쓰기한다.
                                nRegCnt = BitConverter.GetBytes(Convert.ToInt16(vIewWirte ? 1 : 0)); // 사용 여부로 쓸 데이터( 역순이라 서 Reverse 하지 않는다)
                                Array.Copy(nRegCnt, 0, txByte, 10, 2);

                                // Socket 이 연결 되지 않았으면 연결한다.
                                #region 소켓 확인
                                if (socketLoop.Connected == false)
                                {
                                    // 통신소켓 연결 설정
                                    if (Socket_Connection(ctlNode.LIpAddress, ctlNode.ComPort) == false)
                                    {
                                        // 소켓 연결이 안되었을경우 다시 연결한다.
                                        goto SocketError;
                                    }
                                }
                                #endregion 소켓 확인


                                if (socketLoop.Connected)
                                {
                                    #region Etos 데이터 송 수신
                                    try
                                    {
                                        //COMLogger.SetIMCCLog("Info", string.Format("TX Send >> " + this.RAW_HEXA_STRING(txByte)));
                                        socketLoop.Send(txByte, 0, txByte.Length, SocketFlags.None);
                                    }
                                    catch (SocketException ex)
                                    {
                                        COMLogger.SetIMCCLog("Error", string.Format("송신데이터 소켓 오류 : {0}", ex.Message));

                                        // 소켓 문제일 경우 다시 연결한다.
                                        goto SocketError;
                                    }
                                    catch (Exception ex)
                                    {
                                        COMLogger.SetIMCCLog("Error", string.Format("Send Error : {0}", ex.Message));

                                        Thread.Sleep(50);
                                        continue;
                                    }

                                    try
                                    {
                                        socketLoop.Receive(rxLeak, 0, rxLeak.Length, SocketFlags.None, out socError);
                                        //COMLogger.SetIMCCLog("Info", string.Format("RX Receive >> " + this.RAW_HEXA_STRING(rxLeak)));

                                    }
                                    catch (SocketException ex)
                                    {
                                        COMLogger.SetIMCCLog("Error", string.Format("수신데이터 오류 : {0}", ex.Message));

                                        // 소켓 문제일 경우 다시 연결한다.
                                        goto SocketError;
                                    }
                                    catch (Exception ex)
                                    {
                                        COMLogger.SetIMCCLog("Error", string.Format("Receive Error : {0}", ex.Message));
                                        Thread.Sleep(50);
                                        continue;
                                    }
                                    #endregion  Etos 데이터 송 수신
                                }
                            }
                            #endregion USE YN 사용자 처리 확인 하여 Etos 메모리에 쓰기 
                            /****************** 사용여부 읽기 처리 종료 *********************/

                            // point 가 미 사용 체크 되면 건너 뛴다.
                            #region 미사용 포인트 skip 로직
                            if (pointInf[iLoop + j].UseYN.Equals("Y") == false)
                            {
                                pointInf[iLoop + j].PreFaultYN = "N";
                                pointInf[iLoop + j].CurFaultYN = "N";
                                pointInf[iLoop + j].CurFaultCode = "";
                                pointInf[iLoop + j].Alarm1 = false;
                                pointInf[iLoop + j].Alarm2 = false;

                                pointInf[iLoop + j].PreLeakVal = 0;
                                pointInf[iLoop + j].CurLeakVal = 0;

                                continue;
                            }
                            #endregion
                            // point 가 미 사용 체크 되면 건너 뛴다.


                            /****************** fault 처리 시작 *********************/
                            #region Fault 처리 로직

                            // 단선 여부 확인
                            int isDisConnect = rxLeak[m + 11] + rxLeak[m + 12] + rxLeak[m + 13] + rxLeak[m + 14];



                            // 바이트 순번 9 ~ 10
                            string sStatusResult = Convert.ToString(rxLeak[m + 9], 2).PadLeft(8, '0') + Convert.ToString(rxLeak[m + 10], 2).PadLeft(8, '0');

                            // 
                            if (sStatusResult.Length == 16 && isDisConnect > 0) // 2 Byte
                            {

                                iSign = 1;
                                sFaultCodeCheck = "2000";
                                //sFaultYNCheck = "N";

                                if (sStatusResult.Substring(0, 1) == "0")
                                {
                                    iSign = 1;
                                }
                                else
                                {
                                    iSign = -1;
                                }
                                /*
                                    *  정상 code         : 2000
                                    유량 1 error      : 2041
                                    유량 2 error      : 2042
                                    converter 고장    : 2043
                                    sensor 고장       : 2044
                                    */

                                // 유량 1 체크
                                bool bErr = false;
                                if (sStatusResult.Substring(2, 1) == "1")
                                {
                                    bErr = true;
                                    //sFaultYNCheck = "Y";

                                    if (sFaultCodeCheck == "2000")
                                    {
                                        sFaultCodeCheck = "2041";// Error 코드 변경
                                    }
                                }
                                pointInf[iLoop + j].Fault_Flow1 = bErr;

                                // 유량 2 체크
                                bErr = false;
                                if (sStatusResult.Substring(3, 1) == "1")
                                {
                                    bErr = true;
                                    //sFaultYNCheck = "Y";

                                    if (sFaultCodeCheck == "2000")
                                    {
                                        sFaultCodeCheck = "2042";// Error 코드 변경
                                    }
                                }
                                pointInf[iLoop + j].Fault_Flow2 = bErr;

                                // CONVERTER 고장
                                bErr = false;
                                if (sStatusResult.Substring(4, 1) == "1")
                                {
                                    bErr = true;
                                    //sFaultYNCheck = "Y";

                                    if (sFaultCodeCheck == "2000")
                                    {
                                        sFaultCodeCheck = "2043";// Error 코드 변경
                                    }
                                }
                                pointInf[iLoop + j].Fault_Flow2 = bErr;

                                // SENSOR 고장
                                bErr = false;
                                if (sStatusResult.Substring(5, 1) == "1")
                                {
                                    bErr = true;
                                    //sFaultYNCheck = "Y";

                                    if (sFaultCodeCheck == "2000")
                                    {
                                        sFaultCodeCheck = "2044";// Error 코드 변경
                                    }
                                }
                                pointInf[iLoop + j].Fault_Sensor = bErr;

                                // 경보1 
                                bErr = false;
                                if (sStatusResult.Substring(6, 1) == "1")
                                {
                                    bErr = true;
                                    bAlarmOut1 = true;
                                    //sFaultYNCheck = "Y";

                                }
                                pointInf[iLoop + j].Alarm1 = bErr;

                                // 경보2
                                bErr = false;
                                if (sStatusResult.Substring(7, 1) == "1")
                                {
                                    bErr = true;
                                    bAlarmOut2 = true;
                                    //sFaultYNCheck = "Y";
                                }
                                pointInf[iLoop + j].Alarm2 = bErr;
                            }
                            else
                            {
                                // 단선처리
                                sFaultCodeCheck = "5003";

                                // 에러코드 변경
                                pointInf[iLoop + j].PreFaultCode = sFaultCodeCheck;
                                pointInf[iLoop + j].CurFaultCode = sFaultCodeCheck;

                                // 리크값 초기화
                                pointInf[iLoop + j].PreLeakVal = 0;
                                pointInf[iLoop + j].CurLeakVal = 0;

                                // 알람값 초기화
                                pointInf[iLoop + j].Alarm1 = false;
                                pointInf[iLoop + j].Alarm2 = false;

                            }

                            // Fault Led ON
                            if (sFaultCodeCheck.Equals("2000") == false && b_F_ONOFF == false)
                            {
                                b_F_ONOFF = true;
                                LedLampOnOff("F", b_F_ONOFF);
                            }

                            // alarm1, alarm2 Led ON
                            if ((bAlarmOut1 == true || bAlarmOut2 == true) && b_A_ONOFF == false)
                            {
                                b_A_ONOFF = true;
                                LedLampOnOff("A", b_A_ONOFF);
                            }



                            //COMLogger.SetIMCCLog("Info", string.Format("Fault Code is : {0}", sFaultCodeCheck));


                            // 통합 fault 처리
                            if (sFaultCodeCheck != "2000") // 고장이나 통신 단선 둘중에 하나가 있으면 //
                            {
                                // SMCS 로직과 상관없이 현재 Fault 값
                                pointInf[iLoop + j].CurFaultCode = sFaultCodeCheck;

                                if (pointInf[iLoop + j].PreFaultYN != "Y")
                                {
                                    if (FAULT_UPLOAD_DATA(pointInf[iLoop + j], sFaultCodeCheck))
                                    {
                                        pointInf[iLoop + j].PreFaultYN = "Y";
                                        pointInf[iLoop + j].CurFaultYN = "Y";
                                        pointInf[iLoop + j].CurFaultCode = sFaultCodeCheck;
                                        pointInf[iLoop + j].InUp_Gubun = "I";

                                        bSMCS_DB_Status = true;

                                        FAULT_Local_Save(pointInf[iLoop + j]);
                                        COMLogger.SetIMCCLog("Info", string.Format(" SEQ{0}, Card : {1}, Com Port : {2}, GroupSeq{3}, Fault Code is : {4}>>", pointInf[iLoop + j].SEQ, pointInf[iLoop + j].CardNo, pointInf[iLoop + j].ComNo, pointInf[iLoop + j].GrpSeq, sFaultCodeCheck));
                                        //COMLogger.SetIMCCLog("Info", string.Format("==>Fault 발생 upload FaultCode :{0}", sFaultCodeCheck));
                                    }
                                    else
                                    {
                                        bSMCS_DB_Status = false;

                                        COMLogger.SetIMCCLog("Error", string.Format("Fault Address: {0}  Code: {1} Upload Failed", pointInf[iLoop + j].F_Address, sFaultCodeCheck));
                                        continue;
                                        //this.mainApp.WRITE("", Ctl_Node.CTLNode, "ERROR", string.Format("Fault Address: {0}  Code: {1} Upload Failed", sData[iLoop + j].F_ADDRESS, sFaultCode));
                                    }

                                }
                                else
                                {
                                    pointInf[iLoop + j].InUp_Gubun = "";
                                }
                            }
                            else // 고장이 하나도 없으면 
                            {
                                if (pointInf[iLoop + j].PreFaultYN != "N")
                                {
                                    if (FAULT_UPLOAD_DATA(pointInf[iLoop + j], sFaultCodeCheck))
                                    {
                                        pointInf[iLoop + j].PreFaultYN = "N";
                                        pointInf[iLoop + j].CurFaultYN = "N";
                                        pointInf[iLoop + j].CurFaultCode = sFaultCodeCheck;
                                        pointInf[iLoop + j].InUp_Gubun = "U";
                                        FAULT_Local_Save(pointInf[iLoop + j]);

                                        COMLogger.SetIMCCLog("Info", string.Format("==>정상 복귀 upload"));
                                        //this.mainApp.GRID_WRITE_MON_VAL("FAULT", sData[iLoop + j]);
                                    }
                                    else
                                    {
                                        COMLogger.SetIMCCLog("Info", string.Format("Fault Address: {0}  Code: {1} Upload Failed", pointInf[iLoop + j].F_Address, sFaultCodeCheck));
                                        continue;
                                        //this.mainApp.WRITE("", Ctl_Node.CTLNode, "ERROR", string.Format("Fault Address: {0}  Code: {1} Upload Failed", sData[iLoop + j].F_ADDRESS, sFaultCode));
                                    }


                                }
                                else
                                {
                                    pointInf[iLoop + j].InUp_Gubun = "";
                                }

                            }


                            #endregion Fault 처리 로직 종료
                            /****************** fault 처리 종료 *********************/

                            /****************** Leak 처리 시작  *********************/
                            #region Leak 처리 로직


                            // 감지기 상태가 정상일때만 
                            if (sFaultCodeCheck.Equals("2000"))
                            {
                                // 바이트 순번 농도(%): 11 ~ 12, Scale: 13 ~ 14
                                decimal dOutValue = 0.000000000000M;
                                // 농도(%)
                                uint iOutValue0Per = 0;
                                // full scale
                                uint iOutValue0FS = 0;

                                //iOutValue1F = 0;
                                string sOutValue0FS = "";
                                // 농도(%) : 11- 12
                                iOutValue0Per = Convert.ToUInt16(rxLeak[m + 11] * 256 + rxLeak[m + 12]);//   BitConverter.ToInt16.ToInt16(rxByteTemp, 0);
                                dOutValue = Convert.ToDecimal(iSign * iOutValue0Per) / 100.0M; // 농도 (%)

                                // FS : 13 - 14
                                iOutValue0FS = Convert.ToUInt16(rxLeak[m + 13] * 256 + rxLeak[m + 14]);//BitConverter.ToInt16(rxByteTemp, 0);
                                sOutValue0FS = string.Format("{0:D8}", iOutValue0FS); // 10진수로변경(8자리로 채운 )
                                int iSquare = sOutValue0FS.Substring(7, 1).Equals("9") ? -1 : int.Parse(sOutValue0FS.Substring(7, 1));
                                dOutValue = dOutValue * Convert.ToDecimal(sOutValue0FS.Substring(0, 7)) * Convert.ToDecimal(Math.Pow(10, iSquare)); // full scale 계산

                                /*
                                    	Full Scale
                                    Full Scale 의 정보는 3개의 십진수(Decimal)의 숫자로 여겨집니다.
                                    처음 2개의 숫자로 정수값을 정해놓고,
                                    마지막 1개의 숫자로 소수점위치를 정하고 있습니다.
                                    단, 마지막 1개의 숫자로는 일반적인 음수값 표현을 할 수 없으므로
                                    저희는 특수한 계산처리를 하고 있습니다.

                                    예1 [010] -> (01)*10^(0) -> 1.00
                                    예2 [100] -> (10)*10^(0) -> 10.0
                                    예3 [101] -> (10)*10^(1) -> 100
                                    예4 [102] -> (10)*10^(2) -> 1000
                                    예5 [159] -> (15)*10^(-1) -> 1.50
                                    예6 [150] -> (15)*10^(0) -> 15.0
                                    예7 [151] -> (15)*10^(1) -> 150
                                    예8 [152] -> (15)*10^(2) -> 1500

                                    ※여기서 예5가 특수한 처리를 한 예시 입니다.
                                    즉 10^(-1)를 표현하는 방법으로서 「9」를 사용한다는 것입니다.
                                    마지막 1개의 숫자는 실질적으로 「9,0,1,2」만 사용하고 있습니다.

                                */
                                if (dOutValue != pointInf[iLoop + j].PreLeakVal)
                                {
                                    if (LEAK_UPLOAD_DATA(pointInf[iLoop + j], dOutValue))
                                    {
                                        pointInf[iLoop + j].PreLeakVal = dOutValue;
                                        pointInf[iLoop + j].CurLeakVal = dOutValue;


                                        bSMCS_DB_Status = true;

                                        COMLogger.SetIMCCLog("Info", string.Format("==>GAS leak upload : {0}", dOutValue));
                                        //this.mainApp.GRID_WRITE_MON_VAL("LEAK", sData[iLoop + j]);
                                    }
                                    else
                                    {
                                        bSMCS_DB_Status = false;
                                        COMLogger.SetIMCCLog("Error", string.Format("Leak Address: {0}  value: {1} Upload Failed", pointInf[iLoop + j].G_Address, dOutValue));
                                        continue;
                                    }
                                }
                            }
                            #endregion Leak 처리 로직
                            /****************** Leak 처리 종료  *********************/

                            // 통신처리횟수 ++
                            intSuccessCount++;

                        } // for
                        Thread.Sleep(50);
                        #endregion 읽어온 메모리 처리


                        // 메모리 맵 까지 돌기위한 iLoop count 확인용
                        iLoop += iReadAddSize;

                    } // while (iWhileCnt > iLoop)

                    /****************** LED Lamp Alarm Fault( 전체 포인트 램프 상태 확인)   *********************/
                    #region LED 램프 출력
                    if (isCommRun)
                    {
                        LedLampOnOff("F", b_F_ONOFF);
                        LedLampOnOff("A", b_A_ONOFF);
                    }
                    #endregion
                    /****************** LED Lamp Alarm Fault( 전체 포인트 램프 상태 확인)   *********************/


                    /****************** Process 정보 Upload ( Ctl_node 정보 )  *********************/
                    #region Process 정보 Upload
                    bool bUpload = intSuccessCount > 0 && swUpload.ElapsedMilliseconds >= 1000 * 60;
                    //COMLogger.SetIMCCLog("Info", string.Format("SP_CTL_PROCESS ElapsedMilliseconds :{0} , intSuccessCount: {1}, bUpload: {2}", swUpload.ElapsedMilliseconds, intSuccessCount, bUpload));
                    if (bUpload)
                    {
                        if (DB_UPLOAD_CTL(COSMOS_Protocol.ctlNode.CTLNode))
                        {
                            COMLogger.SetIMCCLog("Info", string.Format("Control node data upload Success in SMCS"));

                            //DT_PROCESS_PRE = DateTime.Now;
                            Thread.Sleep(500);
                            swUpload.Reset();
                            swUpload.Start();

                            bSMCS_DB_Status = true;

                            intSuccessCount = 0;
                        }
                        else
                        {
                            bSMCS_DB_Status = false;
                        }
                    }
                    #endregion Process 정보 Upload
                    /****************** Process 정보 Upload ( Ctl_node 정보 )  *********************/

                } // try
                catch (Exception ex)
                {
                    COMLogger.SetIMCCLog("Error", string.Format("Error : {0}", ex.Message));
                    swUpload.Reset();

                }
                finally
                {

                }
                SocketError:
                continue;

            } // while ( 통신 시작 )
            //comThreadMain.Join();
        }// process


        private static void LocalDBSaveManager()
        {
            Stopwatch swLocalLeak = new Stopwatch();
            swLocalLeak.Start();

            while (isCommRun)
            {
                if (swLocalLeak.ElapsedMilliseconds > 1000 * 60 * 9.5)
                {
                    try
                    {
                        LEAK_Local_Save();
                        swLocalLeak.Reset();
                        swLocalLeak.Start();
                    }
                    catch (Exception ex)
                    {
                        swLocalLeak.Reset();
                        swLocalLeak.Start();
                        COMLogger.SetIMCCLog("Error", string.Format("LocalDBSaveManager : {0}", ex.Message));
                    }
                }
                Thread.Sleep(1000);
            }
        }


        private static void LocalDBDeleteManager()
        {
            Stopwatch swDBDelete = new Stopwatch();
            swDBDelete.Start();

            while (isCommRun)
            {
                // 일주일에 한번꼴로 처리
                if (swDBDelete.ElapsedMilliseconds > 7 * 24 * 60 * 60 * 1000)
                //if (swDBDelete.ElapsedMilliseconds > 60 * 1000)
                {
                    try
                    {
                        Local_DB_Delete();
                        swDBDelete.Reset();
                        swDBDelete.Start();
                        COMLogger.SetIMCCLog("Info", string.Format("LocalDBDeleteManager : Local DB Deletion success"));
                    }
                    catch (Exception ex)
                    {
                        swDBDelete.Reset();
                        swDBDelete.Start();
                        COMLogger.SetIMCCLog("Error", string.Format("LocalDBDeleteManager : {0}", ex.Message));
                    }
                }
                Thread.Sleep(1000);
            }
        }



        //private bool[] onLed = new bool[4];

        private static void LedLampOnOff(string sGubun, bool onOff)
        {
            try
            {
                // 정상작동
                if (sGubun.Equals("N") )
                {
                    ModeBusRtu.WriteFunction_B(onOff, 0x08);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Normal Operation {0} ", (onOff ? "On" : "Off")));
                }
                // Fault
                if (sGubun.Equals("F") )
                {
                    ModeBusRtu.WriteFunction_B(onOff, 0x02);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Fault Operation {0} ", (onOff ? "On" : "Off")));
                }
                // Alarm
                if (sGubun.Equals("A") )
                {
                    ModeBusRtu.WriteFunction_B(onOff, 0x04);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Alarm Operation {0} ", (onOff ? "On" : "Off")));
                }
                /*
                // 정상작동
                if (sGubun.Equals("N") && lemp[0] != onOff)
                {
                    lemp[0] = onOff;
                    ModeBusRtu.WriteFunction(onOff, 0x08);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Normal Operation {0} ", (onOff ? "On" : "Off")));
                }
                // Fault
                if (sGubun.Equals("F") && lemp[1] != onOff)
                {
                    lemp[1] = onOff;
                    ModeBusRtu.WriteFunction(onOff, 0x02);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Fault Operation {0} ", (onOff ? "On" : "Off")));
                }
                // Alarm
                if (sGubun.Equals("A") && lemp[2] != onOff)
                {
                    lemp[2] = onOff;
                    ModeBusRtu.WriteFunction(onOff, 0x04);
                    COMLogger.SetIMCCLog("Info", string.Format("LedLampOnOff ==> Alarm Operation {0} ", (onOff ? "On" : "Off")));
                }
                */
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LedLampOnOff : {0}", ex.Message));
            }


        }

        public static void LedLampInitialize()
        {
            try
            {
                // 램프 모두 종료
                if (ModeBusRtu.value != 0x00)
                {
                    ModeBusRtu.GoodByeLed();
                }
                //lemp[0] = false;
                //lemp[1] = false;
                //lemp[2] = false;



                COMLogger.SetIMCCLog("Info", string.Format("LedLampInitialize : {0}", "success"));
            }
            catch (Exception ex)
            {
                //COMLogger.SetIMCCLog("Error", string.Format("LedLampInitialize : lemp[{0}],lemp[{1}],lemp[{2}] ", lemp[0], lemp[1], lemp[2]));
                COMLogger.SetIMCCLog("Error", string.Format("LedLampInitialize : Error Message ==> {0} ", ex.Message));
            }

        }

        private static bool FAULT_Local_Save(PointInfo sensing)
        {
            bool bResult = false;
            try
            {
                dbexe.SaveFaultHistroy(sensing, COSMOS_Protocol.ctlNode);
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("FAULT_Local_Save : {0}", ex.Message));
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL : {0}, Error: {1} ", sbSql.ToString(),ex.Message));
                bResult = false;
            }

            return bResult;
        }

        private static bool LEAK_Local_Save()
        {
            bool bResult = false;
            try
            {
                string sLID = DateTime.Now.ToString("yyyyMMddHHmm");
                int inum = int.Parse(sLID.Substring(sLID.Length - 1, 1));
                if (inum >= 5)
                {
                    sLID = sLID.Substring(0, sLID.Length - 1) + "5";
                }
                else
                {
                    sLID = sLID.Substring(0, sLID.Length - 1) + "0";
                }

                dbexe.SaveLeakHistroyOneTime(pointInf, COSMOS_Protocol.ctlNode, sLID);

                /*
                for (int i = 0; i < COSMOS_Protocol.pointInf.Count; i++)
                {
                    dbexe.SaveLeakHistroy(COSMOS_Protocol.pointInf[i], COSMOS_Protocol.ctlNode, sLID);
                }
                */

                // dbexe.SaveLeakHistroy(sensing, COSMOS_Protocol.ctlNode);
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("LEAK_Local_Save : {0}", ex.Message));
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL : {0}, Error: {1} ", sbSql.ToString(),ex.Message));
                bResult = false;
            }

            return bResult;
        }


        private static bool Local_DB_Delete()
        {
            bool bResult = false;
            try
            {
                dbexe.DeleteHistroy(DateTime.Now.AddDays(-1 * iLocalDB_Saved_Days));
                COMLogger.SetIMCCLog("Info", string.Format("Local_DB_Delete : {0}", "로컬 DB 1년 경과 데이터 삭제 처리 완료"));
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Error", string.Format("Local_DB_Delete : {0}", ex.Message));
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL : {0}, Error: {1} ", sbSql.ToString(),ex.Message));
                bResult = false;
            }

            return bResult;
        }




        private static bool JUST_DB_Check()
        {

            bool bResult = false;
            try
            {
                string sbSql = string.Format(" SELECT GETDATE() ");

                if (dbexe.ExecuteNonQuery(sbSql))
                {
                    bResult = true;
                }
            }
            catch (Exception ex)
            {
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL : {0}, Error: {1} ", sbSql.ToString(),ex.Message));
                bResult = false;
            }


            return bResult;
        }


        private static bool FAULT_UPLOAD_DATA(PointInfo sensing, string sFaultCode)
        {

            bool bResult = false;
            try
            {
                string sbSql = string.Format("EXEC SP_UPLOAD_FAULT_DATA '{0}' , '{0}.{1}' , '{2}' ", COSMOS_Protocol.ctlNode.CTLNode, sensing.F_Address, sFaultCode);

                if (dbexe.ExecuteNonQuery(sbSql))
                {
                    bResult = true;
                }
            }
            catch (Exception ex)
            {
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL : {0}, Error: {1} ", sbSql.ToString(),ex.Message));
                bResult = false;
            }


            return bResult;
        }

        /// <summary>
        /// Fault, Leak 데이터 업로드 로직
        /// </summary>
        private static bool LEAK_UPLOAD_DATA(PointInfo sensing, decimal fLeakVal)
        {

            bool bResult = false;
            try
            {
                string sbSql = string.Format("EXEC SP_UPLOAD_LEAK_DATA '{0}' , '{0}.{1}' , {2:#0.000} , '{3}' ", COSMOS_Protocol.ctlNode.CTLNode, sensing.G_Address, fLeakVal, "D");

                if (dbexe.ExecuteNonQuery(sbSql))
                {
                    bResult = true;
                }
            }
            catch (Exception ex)
            {
                //mainApp.WRITE("F", Ctl_Node.CTLNode, "ERROR", string.Format("SQL: {0},  Error: {1} ", sbSql.ToString(), ex.Message));
                bResult = false;
            }
            return bResult;
        }

        public static bool DB_UPLOAD_CTL(string paramCtlNode)
        {
            bool boRtn = false;

            StringBuilder sbSql = new StringBuilder();

            try
            {
                sbSql.Clear();
                sbSql.AppendFormat("EXEC SP_CTL_PROCESS '{0}' ", paramCtlNode);

                if (dbexe.ExecuteNonQuery(sbSql.ToString()))
                {
                    COMLogger.SetIMCCLog("Info", string.Format("Node:{0} 단선정보(SP_CTL_PROCESS) upload Success", paramCtlNode));
                    //xApp.WRITE("", paramSeq, paramCtlNode, xRow.POINT_NM, "SQL", sbSql.ToString());
                    boRtn = true;
                }
            }
            catch (Exception ex)
            {
                COMLogger.SetIMCCLog("Info", string.Format("Error Node:{0} 단선정보(SP_CTL_PROCESS) upload Failed: {1}", paramCtlNode, ex.Message));
                //xApp.GRID_WRITE_MON(Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, 2); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", sbSql.ToString() + ">>" + ex.Message);
                Thread.Sleep(50);
            }

            sbSql = null;

            return boRtn;
        }



        private static bool Socket_Connection(string sIPAddress, int iPort)
        {
            bool boRtn = false;

            try
            {
                socketLoop = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                socketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                socketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, 3000);
                socketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, 3000);
                socketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.NoDelay, true);

                socketLoop.Connect(sIPAddress, iPort);

                COMLogger.SetIMCCLog("Info", string.Format("IP : {0} ,Port : {1} access Success... ", sIPAddress, iPort));

                boRtn = true;
                bEtos_COMM_Status = true;

            }
            catch (SocketException ess)
            {
                boRtn = false;
                bEtos_COMM_Status = false;
                COMLogger.SetIMCCLog("Error", string.Format("Error Socket IP : {0} ,Port : {1} access failed... ", sIPAddress, iPort));

            }
            catch (Exception es)
            {
                boRtn = false;
                bEtos_COMM_Status = false;
                COMLogger.SetIMCCLog("Error", string.Format(es.Message));
            }

            return boRtn;
        }

        public static string REVERSE_STRING(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        public static string RAW_HEXA_STRING(byte[] bytes)
        {
            string strRtn = "";

            StringBuilder xHEX_Data = new StringBuilder();

            try
            {
                xHEX_Data.Clear();

                foreach (byte drLoop in bytes)
                {
                    xHEX_Data.Append(" " + string.Format("{0:X2}", Convert.ToInt32(drLoop)));
                }

                strRtn = xHEX_Data.ToString();
            }
            catch (Exception es)
            {
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", "RAW Convert Error:" + es.Message);
            }

            xHEX_Data = null;
            return strRtn;
        }

    }
}
