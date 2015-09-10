using System;
using System.Collections.Generic;
using System.Text;

using MicroTorque.Exceptions;

namespace MicroTorque
{
    namespace Protocol
    {

        public class AsciiConnection : MTCom.Connection
        {
            #region Constants
            const byte STX = 0x02;
            const byte ETX = 0x03;
            const byte ENQ = 0x05;
            const byte ACK = 0x06;
            const byte DLE = 0x10;
            const byte NAK = 0x15;
            #endregion

            #region Properties
            public int ReadTimeout { get; set; }
            public int MaxRetries { get; set; }
            #endregion

            #region Members
            static readonly Encoding mtEncoding = Encoding.GetEncoding(28591);
            object threadLock = new object();
            #endregion

            /// <summary>

            public AsciiConnection()
                : base()
            {
                ReadTimeout = 1000;
                MaxRetries = 2;
            }

            /// <summary>
            /// Opens a connection to a MicroTorque device.
            /// </summary>
            /// <param name="devInfo">Device info structure as received from MTCom.SafeNativeMethods.GetDeviceEntryEx.</param>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public void Open(MTCom.DeviceInfo devInfo)
            {
                this.Open(devInfo.DevId, MTCom.ComProtocol.MT_PROT_ASCII, 0);
            }

            /// <summary>
            /// Opens a connection to a MicroTorque device.
            /// </summary>
            /// <param name="connStr">Connection string of device to open, see documentation of MTCom.ComInterface.</param>
            /// <param name="protocol">Protocol to use.</param>
            /// <param name="param">Protocol parameter (reserved for future use).</param>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public override void Open(string connStr, MTCom.ComProtocol protocol, int param)
            {
                lock (threadLock)
                {
                    base.Open(connStr, protocol, param);
                }
            }

            /// <summary>
            /// Closes a connection to a MicroTorque device.
            /// </summary>
            public override void Close()
            {
                lock (threadLock)
                {
                    if (IsOpen)
                    {
                        base.Close();
                    }
                }
            }

            /// <summary>
            /// Sends a command and waits for an ACK/NAK.
            /// </summary>
            /// <param name="command">Command to send.</param>
            /// <returns>True if an ack was received, false otherwise.</returns>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            /// <exception cref="MicroTorque.Exceptions.ProtocolTimeoutException"></exception>
            public bool SendWaitAck(string command)
            {
                bool result = false;

                lock (threadLock)
                {
                    byte[] txData = mtEncoding.GetBytes(command);
                    Send(txData, txData.Length);

                    for (int retries = 0; ; retries++)
                    {
                        byte[] rxData = Receive(ReadTimeout);

                        if (IsValidMessage(rxData) && rxData.Length == 6)
                        {
                            if (rxData[2] == 'A')
                            {
                                result = true;
                                break;
                            }

                            if (rxData[2] == 'N')
                            {
                                result = false;
                                break;
                            }
                        }

                        if (retries >= MaxRetries)
                            throw new ProtocolTimeoutException("Timeout waiting for response.");
                    }
                }

                return result;
            }

            /// <summary>
            /// Sends a protocol message and waits for a reply.
            /// </summary>
            /// <param name="message">Message to send.</param>
            /// <param name="cmdLength">Length of the command part of the protocol message.</param>
            /// <returns>Received reply.</returns>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            /// <exception cref="MicroTorque.Exceptions.ProtocolTimeoutException"></exception>
            public string SendWaitReply(string message, int cmdLength)
            {
                lock (threadLock)
                {
                    string reply = "";
                    byte[] txData = mtEncoding.GetBytes(message);
                    Send(txData, txData.Length);

                    for (int retries = 0; ; retries++)
                    {
                        byte[] rxData = Receive(ReadTimeout);

                        if (IsValidMessage(rxData))
                        {
                            reply = mtEncoding.GetString(rxData);

                            if (reply.Length > (cmdLength + 2)
                                && message.StartsWith(reply.Substring(2, cmdLength)))
                            {
                                // Response received
                                break;
                            }
                        }

                        if (retries >= MaxRetries)
                            throw new ProtocolTimeoutException("Timeout waiting for response.");
                    }

                    return reply;
                }
            }

            /// <summary>
            /// Checks if an ASCII protocol message is valid by verifying its length and checksum.
            /// </summary>
            /// <param name="msg">Message to verify.</param>
            /// <returns>True if the message is valid.</returns>
            static bool IsValidMessage(byte[] msg)
            {
                bool isValid = false;
                int checksum = 0;

                if (msg.Length >= 5 && msg[0] == STX && msg[msg.Length - 1] == ETX)
                {
                    string chkString;

                    if (msg[1] == DLE)
                    {
                        chkString = mtEncoding.GetString(msg, msg.Length - 4, 2);

                        for (int i = 2; i < msg.Length - 4; i++)
                            checksum = checksum + msg[i];
                    }
                    else
                    {
                        chkString = mtEncoding.GetString(msg, msg.Length - 3, 2);

                        for (int i = 1; i < msg.Length - 3; i++)
                            checksum = checksum + msg[i];
                    }

                    isValid = ((checksum & 0xff) == Int16.Parse(chkString, System.Globalization.NumberStyles.HexNumber));
                }

                return isValid;
            }
        }
    }

    #region Protocol exceptions
    namespace Exceptions
    {
        public class ProtocolException : MTException
        {
            public ProtocolException(string message) : base(message) { }
        }

        public class ProtocolTimeoutException : ProtocolException
        {
            public ProtocolTimeoutException(string message) : base(message) { }
        }

        public class ProtocolNakException : ProtocolException
        {
            public ProtocolNakException(string message) : base(message) { }
        }
    }
    #endregion
}
