using System;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace IMCC_ALL_OF_POE_GET.Adaptors
{

    public abstract class Adaptor 
    {
        protected Socket xsocketLoop;
        protected SocketError socError;
        

        public Adaptor()
        {
            
        }

        public abstract void Process(Object threadContext);


        private bool CheckPortStatus(string CheckIP, int CheckPort)
        {
            try
            {
#if true
                // 20150918-ydy-기존 소켓 통신이 Sync 모드여서 오래걸리는 부분 Async로 변경
                IAsyncResult result = null;
                bool success = false;
                Socket client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                client.LingerState = new LingerOption(true, 1);

                if (CheckIP != null && CheckIP.Length > 0)
                {
                    result = client.BeginConnect(IPAddress.Parse(CheckIP), CheckPort, null, null);
                }
                else
                {
                    result = client.BeginConnect(IPAddress.Loopback, CheckPort, null, null);
                }
                success = result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(1));
                if (success == true)
                {
                    client.EndConnect(result);
                    if (client.Connected == true)
                    {
                        client.Shutdown(SocketShutdown.Both);
                    }
                    client.Close();
                    return true;
                }
                else
                {
                }

                client.Close();
#else
                TcpClient client = new TcpClient();
                if (CheckIP != null && CheckIP.Length > 0)
                {
                    client.Connect(new IPEndPoint(IPAddress.Parse(CheckIP), CheckPort));
                }
                else
                {
                    client.Connect(new IPEndPoint(IPAddress.Loopback, CheckPort));
                }
                client.Close();
#endif
            }
            catch (Exception ex)
            {
            }
            return false;
        }


        public bool Socket_Connection(string sIPAddress, int iPort)
        {
            bool boRtn = false;

            socError = new SocketError();

            try
            {
                try
                {
                    boRtn = this.CheckPortStatus(sIPAddress, iPort);
                    //boRtn = true;
                }
                catch {
                    boRtn = false;
                    throw new Exception("Network Connection Failed");
                }
                if (boRtn)
                {
                    xsocketLoop = null;
                    xsocketLoop = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    xsocketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                    xsocketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, 2000);
                    xsocketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, 2000);
                    xsocketLoop.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.NoDelay, true);

                    xsocketLoop.Connect(new IPEndPoint(IPAddress.Parse(sIPAddress), iPort));

                    boRtn = true;
                }

                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "INFO", "Socket Connection Success");
            }
            catch (SocketException ess)
            {
                boRtn = false;
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", "Socket Connect Fail.." + ess.Message);
                throw ess;
                
            }
            catch (Exception es)
            {
                boRtn = false;
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", "Socket Connect Fail.." + es.Message);
                throw es;
                
            }

            return boRtn;
        }

        

        protected string REVERSE_STRING(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        protected string RAW_HEXA_STRING(byte[] bytes)
        {
            if (bytes == null) return "";
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
            catch(Exception es)
            {
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", "RAW Convert Error:" + es.Message);
            }

            xHEX_Data = null;
            return strRtn;
        }


        public virtual string DB_UPLOAD_DATA(int idx, string ctlNode, string sNormalCode ="")
        {
            StringBuilder sbSql = new StringBuilder();

            string sMessage = "";

            var tempRows = from tempRow in CommonData.saveDt
                           where tempRow.SEQ == idx && tempRow.CTL_NODE == ctlNode
                           select tempRow;
            try
            {


                if (tempRows.First().FAULT_CODE != tempRows.First().FAULT_CODE_PRE )
                {
                    sbSql.Clear();
                    sbSql.AppendFormat("EXEC SP_UPLOAD_FAULT_DATA '{0}' , '{1}' , '{2}' ", tempRows.First().CTL_NODE, tempRows.First().F_ADDRESS, tempRows.First().FAULT_CODE);

                    if (CommonData.ExecuteNonQuery(sbSql.ToString()))
                    {
                        tempRows.First().FAULT_CODE_PRE = tempRows.First().FAULT_CODE;
                        sMessage = string.Format("DB_UPLOAD_DATA>>UPLOAD_Fault>>Address {0}, FaultCode:{1}", tempRows.First().F_ADDRESS, tempRows.First().FAULT_CODE);
                    }
                }
            }
            catch(Exception ex)
            {
                throw new Exception(string.Format("DB_UPLOAD_DATA>>UPLOAD_Fault>>Address {0}, FaultCode:{1}", tempRows.First().F_ADDRESS, tempRows.First().FAULT_CODE) + "==> " + ex.Message);
            }

            try
            {
                string NormalCode = sNormalCode.Equals("") ? "0000" : sNormalCode;
                if (tempRows.First().FAULT_CODE.Equals(NormalCode) && tempRows.First().LEAK_VAL != tempRows.First().LEAK_VAL_PRE)
                {
                    sbSql.Clear();
                    sbSql.AppendFormat("EXEC SP_UPLOAD_LEAK_DATA '{0}' , '{1}' , {2:#0.000} , '{3}' ", tempRows.First().CTL_NODE, tempRows.First().G_ADDRESS, tempRows.First().LEAK_VAL, "D");

                    if (CommonData.ExecuteNonQuery(sbSql.ToString()))
                    {
                        tempRows.First().LEAK_VAL_PRE = tempRows.First().LEAK_VAL;
                        sMessage += sMessage.Equals("") ? "" : System.Environment.NewLine;
                        sMessage += string.Format("DB_UPLOAD_DATA>>UPLOAD_LEAK>>Address {0}, Value:{1}", tempRows.First().G_ADDRESS, tempRows.First().LEAK_VAL);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception (string.Format("DB_UPLOAD_DATA>>UPLOAD_LEAK>>Address {0}, Value:{1}", tempRows.First().G_ADDRESS, tempRows.First().LEAK_VAL)+"==> " + ex.Message);
            }

            sbSql = null;
            return sMessage;
        }

        public bool DB_UPLOAD_DATA_NETWORKERROR(int idx, string ctlNode)
        {
            bool boRtn = false;

            StringBuilder sbSql = new StringBuilder();

            var tempRows = from tempRow in CommonData.saveDt
                           where tempRow.SEQ == idx && tempRow.CTL_NODE == ctlNode
                           select tempRow;


            try
            {
                if (tempRows.First().FAULT_CODE != tempRows.First().FAULT_CODE_PRE)
                {
                    sbSql.Clear();
                    sbSql.AppendFormat("EXEC SP_UPLOAD_FAULT_DATA '{0}' , '{1}' , '{2}' ", tempRows.First().CTL_NODE, tempRows.First().F_ADDRESS, tempRows.First().FAULT_CODE);

                    if (CommonData.ExecuteNonQuery(sbSql.ToString()))
                    {
                        tempRows.First().FAULT_CODE_PRE = tempRows.First().FAULT_CODE;
                    }
                }

                boRtn = true;
            }
            catch (Exception ex)
            {
                //xApp.GRID_WRITE_MON(Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, 2); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", sbSql.ToString() + ">>" + ex.Message);
                //Thread.Sleep(50);
                boRtn = false;
                sbSql = null;
                throw ex;
            }
            sbSql = null;




            return boRtn;
        }


        public virtual bool DB_UPLOAD_CTL(string sCtl_Node)
        {
            bool boRtn = false;

            //StringBuilder sbSql = new StringBuilder();
            string sSql = "";
            try
            {

                sSql = string.Format("EXEC SP_CTL_PROCESS '{0}' ", sCtl_Node);

                if (CommonData.ExecuteNonQuery(sSql))
                {
                    //xApp.WRITE("", paramSeq, paramCtlNode, xRow.POINT_NM, "SQL", sbSql.ToString());
                    boRtn = true;
                }
            }
            catch (Exception ex)
            {
                //xApp.GRID_WRITE_MON(Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, 2); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER
                //xApp.WRITE("", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "ERROR", sbSql.ToString() + ">>" + ex.Message);
                //Thread.Sleep(50);
                throw ex;
            }
            return boRtn;
        }
        
        private void displaySensingValue()
        {
            while (CommonData.boProgramRun == true)
            {
                //this.xApp.GRID_WRITE_MON(xSaveData, 1);
                Thread.Sleep(1000);
            }
        }


        public virtual void MakeSendMessage(int iAdd, int RegCount, ref byte[] txbyteSendMsg)
        {
            byte[] n2Byte = new byte[2];
            //byte[] txbyteSendMsg = new byte[12];

            Array.Clear(txbyteSendMsg, 0, txbyteSendMsg.Length);

            try
            {

                // 주소값 변환

                // 상태값 요청 //
                txbyteSendMsg[0] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)  
                txbyteSendMsg[1] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)
                txbyteSendMsg[2] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[3] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[4] = 0x00;     // 00 LENGTH (고정)
                txbyteSendMsg[5] = 0x06;     // 06 LENGTH (고정)
                txbyteSendMsg[6] = 0x01;     // 01 UNIT ID (바꿀수도 있음) 
                txbyteSendMsg[7] = 0x03;     // 03 FUNCTION CODE (거의 고정)

                txbyteSendMsg[8] = 0x00;     // 02 START ADDRESS (변동)
                txbyteSendMsg[9] = 0x00;     // 03 START ADDRESS (변동)
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(iAdd));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 8, 2);

                txbyteSendMsg[10] = 0x00;     // 00 QUANTITY (변동)
                txbyteSendMsg[11] = 0x01;     // 00 QUANTITY (변동) 1Register = 2byte
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(RegCount));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 10, 2);

                
            }
            catch (SocketException exSocket)
            {
                throw exSocket;
            }
            catch (Exception exStatus)
            {
                throw exStatus;
            }


        }

        public virtual byte[] ReadHoldingRegister(int iAdd, int RegCount, int iRecByteCount)
        {
            byte[] n2Byte = new byte[2];
            byte[] txbyteSendMsg = new byte[12];
            byte[] rxbyteReceive = new byte[iRecByteCount];

            try
            {

                // 주소값 변환

                // 상태값 요청 //
                txbyteSendMsg[0] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)  
                txbyteSendMsg[1] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)
                txbyteSendMsg[2] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[3] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[4] = 0x00;     // 00 LENGTH (고정)
                txbyteSendMsg[5] = 0x06;     // 06 LENGTH (고정)
                txbyteSendMsg[6] = 0x01;     // 01 UNIT ID (바꿀수도 있음) 
                txbyteSendMsg[7] = 0x03;     // 03 FUNCTION CODE (거의 고정)

                txbyteSendMsg[8] = 0x00;     // 02 START ADDRESS (변동)
                txbyteSendMsg[9] = 0x00;     // 03 START ADDRESS (변동)
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(iAdd));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 8, 2);

                txbyteSendMsg[10] = 0x00;     // 00 QUANTITY (변동)
                txbyteSendMsg[11] = 0x01;     // 00 QUANTITY (변동) 1Register = 2byte
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(RegCount));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 10, 2);


                xsocketLoop.Send(txbyteSendMsg, 0, txbyteSendMsg.Length, SocketFlags.None);
                //xApp.GRID_WRITE_MON(xRow.SEQ, xRow.CTL_NODE, 1); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER

                //xApp.WRITE("F", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "RAW", "TX STATUS >> " + RAW_HEXA_STRING(txbyteSendMsg));
                Thread.Sleep(200);

                Array.Clear(rxbyteReceive, 0, rxbyteReceive.Length);
                xsocketLoop.Receive(rxbyteReceive, 0, rxbyteReceive.Length, SocketFlags.None, out socError);
                //xApp.WRITE("F", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "RAW", "RX STATUS >> " + RAW_HEXA_STRING(rxbyteReceive));
            }
            catch (SocketException exSocket)
            {
                throw exSocket;
            }
            catch (Exception exStatus)
            {
                throw exStatus;
            }

            return rxbyteReceive;
        }


        public virtual void ReadHoldingRegister(ref byte[] txbyteSendMsg, int iAdd, int RegCount, ref byte[] rxbyteReceive)
        {
            byte[] n2Byte = new byte[2];
            Array.Clear(txbyteSendMsg, 0, txbyteSendMsg.Length);
            Array.Clear(rxbyteReceive, 0, rxbyteReceive.Length);


            try
            {

                // 주소값 변환

                // 상태값 요청 //
                txbyteSendMsg[0] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)  
                txbyteSendMsg[1] = 0x00;     // 00 TRANSACTION ID  (바꿀수도 있음)
                txbyteSendMsg[2] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[3] = 0x00;     // 00 PROTOCOL ID  (바꿀수도 있음)
                txbyteSendMsg[4] = 0x00;     // 00 LENGTH (고정)
                txbyteSendMsg[5] = 0x06;     // 06 LENGTH (고정)
                txbyteSendMsg[6] = 0x01;     // 01 UNIT ID (바꿀수도 있음) 
                txbyteSendMsg[7] = 0x03;     // 03 FUNCTION CODE (거의 고정)

                txbyteSendMsg[8] = 0x00;     // 02 START ADDRESS (변동)
                txbyteSendMsg[9] = 0x00;     // 03 START ADDRESS (변동)
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(iAdd));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 8, 2);

                txbyteSendMsg[10] = 0x00;     // 00 QUANTITY (변동)
                txbyteSendMsg[11] = 0x01;     // 00 QUANTITY (변동) 1Register = 2byte
                Array.Clear(n2Byte, 0, n2Byte.Length);
                n2Byte = BitConverter.GetBytes(Convert.ToInt16(RegCount));
                Array.Reverse(n2Byte);
                Array.Copy(n2Byte, 0, txbyteSendMsg, 10, 2);


                xsocketLoop.Send(txbyteSendMsg, 0, txbyteSendMsg.Length, SocketFlags.None);
                //xApp.GRID_WRITE_MON(xRow.SEQ, xRow.CTL_NODE, 1); // 0 - STOP 1 - SUCCESS 2 - FAIL  3-9 OTHER

                //xApp.WRITE("F", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "RAW", "TX STATUS >> " + RAW_HEXA_STRING(txbyteSendMsg));
                Thread.Sleep(200);

                
                xsocketLoop.Receive(rxbyteReceive, 0, rxbyteReceive.Length, SocketFlags.None, out socError);
                //xApp.WRITE("F", Convert.ToInt32(xRow.SEQ), xRow.CTL_NODE, xRow.POINT_NM, "RAW", "RX STATUS >> " + RAW_HEXA_STRING(rxbyteReceive));
            }
            catch (SocketException exSocket)
            {
                throw exSocket;
            }
            catch (Exception exStatus)
            {
                throw exStatus;
            }

        }

        public virtual void TRX_HoldingRegister1(byte[] txbyteSendMsg, ref byte[] rxbyteReceive)
        {

            try
            {
                Thread.Sleep(200);

                xsocketLoop.Send(txbyteSendMsg, 0, txbyteSendMsg.Length, SocketFlags.None);

                Thread.Sleep(200);

                Array.Clear(rxbyteReceive, 0, rxbyteReceive.Length);
                xsocketLoop.Receive(rxbyteReceive, 0, rxbyteReceive.Length, SocketFlags.None, out socError);

                
            }
            catch (SocketException exSocket)
            {
                throw exSocket;
            }
            catch (Exception exStatus)
            {
                throw exStatus;
            }
        }



        public virtual void SaveDataFault(int idx, string ctlNode, string paramFaultYn, string paramFaultCode)
        {
            var tempRows = from tempRow in CommonData.saveDt
                           where tempRow.SEQ == idx && tempRow.CTL_NODE == ctlNode
                           select tempRow;

            tempRows.First().FAULT_YN = paramFaultYn;
            tempRows.First().FAULT_CODE = paramFaultCode;
            tempRows.First().ChangeDt = DateTime.Now;


        }

        public virtual void SaveDataLeak(int idx, string ctlNode, Single paramLeakVal)
        {

            var tempRows = from tempRow in CommonData.saveDt
                           where tempRow.SEQ == idx && tempRow.CTL_NODE == ctlNode
                           select tempRow;

            tempRows.First().LEAK_VAL = paramLeakVal;
            tempRows.First().ChangeDt = DateTime.Now;
        }






    } // Class

} // Namespace
