using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MicroFocus.RDE.Framework;
using System.Diagnostics;
using System.Drawing;
namespace MainframeAutomation
{
    class EhlApi : EhlApiInterface
    {
        
        static IntPtr ehlPtr = (IntPtr)100;
        static string sessionName = ApiSettings.Default.SessionName;
        public EhlApi()
        {
            SetSessionParams();
        }
         Boolean GetConnection()
        {
            short sessionReturnCode;
        
            try
            {
                sessionReturnCode = EHLLAPI.WD_ConnectPS(ehlPtr, sessionName);
                CheckException(sessionReturnCode);
                return true;
            }
            catch (EhlApiExceptions exception)
            {
                return false;
            }
            catch (EhlApiError error)
            {
                System.Console.WriteLine("Unable to connect");
                System.Environment.Exit(1);
                return false;
            }

     
        }

         Boolean InitialCheck(String expectedMenu)
        {
            if (!expectedMenu.Equals(""))
            {

                if (SearchPS(1, expectedMenu) != 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

         Boolean ReserveSession()
        {
            try
            {
                short reserveSessionReturnCode = EHLLAPI.WD_Reserve(ehlPtr);
                CheckException(reserveSessionReturnCode);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

         Boolean ReleaseSession()
        {
            try
            {
                short releaseSessionReturnCode = EHLLAPI.WD_Release(ehlPtr);
                CheckException(releaseSessionReturnCode);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

         void StartHostNotification()
        {
           
            StringBuilder notificationParms = new StringBuilder(6);
            notificationParms.Append(sessionName);
            notificationParms.Append("   P");
//            notificationParms.Append("P");
            //      notificationParms.Append("   ");

            short hostNotifyReturnCode = EHLLAPI.WD_StartHostNotification(ehlPtr, notificationParms);
            CheckException(hostNotifyReturnCode);
        
        }
         void PauseAndReset()
        {
          
   //         System.Console.WriteLine("Wait");
             StringBuilder queryHostUpdtParms = new StringBuilder(sessionName);
             short pauseReturnCode = EHLLAPI.WD_Pause(ehlPtr, 10);
             short queryReturnCode= EHLLAPI.WD_QueryHostUpdate(ehlPtr, queryHostUpdtParms);
   //          System.Console.WriteLine("End Wait");
             
   
  
        }
         string CopyPSToString(short startPosition, short length)
         {
             try
             {
             //    StartHostNotification();
                 StringBuilder output = new StringBuilder() ;
                 output.Length = length;
                 short CopyPSReturnCode = EHLLAPI.WD_CopyPSToString(ehlPtr, startPosition, output, length);
                //hort copy = EHLLAPI.WD_CopyPSToString(ehlPtr,1,StringBuilder Ad,length);
           //      PauseAndReset();
                 CheckException(CopyPSReturnCode);
                 return output.ToString();
             }
             catch (Exception e)
             {

                 return null;
             }

         }
         public string GetStatus()
         {
             string status = null;

             try
             {
                 DisconnectAndConnect();
                 ReserveSession();
                 status = CopyPSToString(3, 60);

                 ReleaseSession();
                 endConnection();
                 return status;

             }
             catch (EhlApiExceptions exception)
             {
                 return null;
             }

         }


         Boolean SendKey(String keyValue) 
        {
            try
            {
      //          System.Console.WriteLine("SendKey");
                StartHostNotification();
                short sendKeyReturnCode;
                sendKeyReturnCode = EHLLAPI.WD_SendKey(ehlPtr, keyValue);
        //      
                PauseAndReset();
                CheckException(sendKeyReturnCode);
           
                return true;
            }
            catch (Exception e)
            {

                return false;
            }
        }
        Boolean SetSessionParams()
        {
   //         System.Console.WriteLine("SetSessionParams");
            short sessionParmsReturnCode;
            try
            {
                sessionParmsReturnCode = EHLLAPI.WD_SetSessionParameters(ehlPtr, EHLLAPI.SSP_TWAIT);
                CheckException(sessionParmsReturnCode);
                sessionParmsReturnCode = EHLLAPI.WD_SetSessionParameters(ehlPtr, EHLLAPI.SSP_IPAUSE);
                CheckException(sessionParmsReturnCode);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        String QuerySessionStatus()
        {
      //      System.Console.WriteLine("Query Session");
            return null;
        }

        Boolean DisconnectAndConnect()
       {

           endConnection();
           SetSessionParams();

           while (!GetConnection())
           {
       //        System.Console.WriteLine("host busy");
           }
           return true;
       }

         short SearchField(short searchPosition,String searchData)
        {
            try
            {
                short searchReturnCode;
                short returnLocation;

                searchReturnCode = EHLLAPI.WD_SearchField(ehlPtr, out returnLocation, searchPosition, searchData);
                CheckException(searchReturnCode);
                return returnLocation;
            }
            catch (Exception e)
            {
                return 0;
            }
        }

         short SearchPS(short searchPosition, String searchData)
        {
            try
            {
     //           System.Console.WriteLine("Search PS");
                short searchReturnCode;
                short returnLocation;

                searchReturnCode = EHLLAPI.WD_SearchPS(ehlPtr, out returnLocation, searchPosition, searchData);
                CheckException(searchReturnCode);
              //  short returnPosition = returnLocation;
                return returnLocation;
            }
            catch (Exception e)
            {
                return 0;
            }
        }

         short FindFieldPosition(String searchData)
        {
            try
            {
     //           System.Console.WriteLine("Find Field Position");
                short location = SearchPS(1, searchData);
                short fieldLocation;
                short findFieldPositionRetCode = EHLLAPI.WD_FindFieldPosition(ehlPtr, out fieldLocation, location, "NU");
                CheckException(findFieldPositionRetCode);
                return fieldLocation;
            }
            catch (EhlApiExceptions e)
            {
                return 0;
            }


        }

         short FindFieldLength(String searchData)
        {
   //         System.Console.WriteLine("Find Field");
            return 2;
        }

         Boolean CopyStringToField(short position, String data)
        {
            try
            {

      //          System.Console.WriteLine("Copy to Field");
                short copyStringFieldReturnCode;
                copyStringFieldReturnCode = EHLLAPI.WD_CopyStringToField(ehlPtr, position, data);
                CheckException(copyStringFieldReturnCode);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }


         Boolean CopyStringToPS(short position, short length, String data)
        {
            try
            {
 //               System.Console.WriteLine("Copy to PS");
                short copyStringPSReturnCode;
                copyStringPSReturnCode = EHLLAPI.WD_CopyStringToPS(ehlPtr, position, data, length);
           
                CheckException(copyStringPSReturnCode);
          //      Wait("KEY", "");
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

         short SetCursor(short position)
        {
         //   short cursorPosition = 0;
            try
            {
                
                short serCursorReturnCode;
                serCursorReturnCode = EHLLAPI.WD_SetCursor(ehlPtr, position);
                CheckException(serCursorReturnCode);
                return position;
            }
            catch (Exception e)
            {
                return 0;
            }
        }


         void endConnection()
        {
            EHLLAPI.WD_ResetSystem(ehlPtr);
            EHLLAPI.WD_DisconnectPS(ehlPtr);
        }
         short SearchAndPlaceCursor(String searchField, short offset)
        {
            try
            {
                short location = SearchPS(1, searchField);
                if (location != 0)
                {
                    //          StartHostNotification();
                    location += (short)searchField.Length;
                    location += (short)offset;
                    short currentLocation = SetCursor(location);
                    if (currentLocation != 0)
                    {
                        return location;
                    }
                    //           PauseAndReset();
                }
                return 0;
                    
            }
            catch (Exception e)
            {
                return 0;
            }
        }

        //public short SearchforCommandPosition()
        //{
        //    try
        //    {
        //      //  short commandPosition = 0;
        //        string commandField = "==>";
        //        short location = SearchPS(1, commandField);
        //        location += (short)commandField.Length;
        //        return (location);
        //    }
        //    catch (Exception e)
        //    {
        //        return 0;
        //    }
        //}

        public Boolean ExecuteCommand(String command,String expectedMenu)
        {
            try
            {
                DisconnectAndConnect();
                ReserveSession();

                System.Console.WriteLine("Execute Command");
               
                    //             short commandLocation = SearchforCommandPosition();
                    //             short curentLocation = SetCursor(commandLocation);
                    if (InitialCheck(expectedMenu))
                    {
                        if (command.Equals("ENTER"))
                        {
                            if (SendKey("@E"))
                            {
                                ReleaseSession();
                                endConnection();
                                return true;
                            }
                      
                        }
                        if (command.Equals("PF05"))
                        {
                            if (SendKey("@5"))
                            {
                                ReleaseSession();
                                endConnection();
                                return true;
                            }

                        }

                        if (command.Equals("PF03"))
                        {
                            if (SendKey("@3"))
                            {
                                ReleaseSession();
                                endConnection();
                                return true;
                            }

                        }
                        else
                        {
                            short curentLocation = SearchAndPlaceCursor("==>", 2);
                            if (curentLocation != 0)
                            {
                                short commandLength = (short)command.Length;
                                if (CopyStringToPS(curentLocation, commandLength, command))
                                {
                                    if (SendKey("@E"))
                                    {
                                        ReleaseSession();
                                        endConnection();
      //                                  System.Console.WriteLine("end of send key");
                                        return true;
                                    }
                                }
                               
                            }
                        }
                }
                

                ReleaseSession();
                endConnection();

                throw new EhlApiExceptions("Command Execution failed:" + command);
               
            }
            catch (EhlApiExceptions exception)
            {
                return false;
            }
        }

        public Boolean SearchFieldAndPopulate(String fieldName, String fieldValue)
        {
            try
            {
                DisconnectAndConnect();
                ReserveSession();
   //             System.Console.WriteLine("Search and populate");
                short fieldPosition = FindFieldPosition(fieldName);
         //       short curentLocation = SearchAndPlaceCursor(fieldName, 2);

           //     System.Drawing.Point point = new System.Drawing.Point((int)fieldPosition);
           ////         System.Drawing.Size
           //                    EHLLAPI.WD_Convert(ehlPtr, EHLLAPI.CONVERT_POSITION, ref point, "A");
           //    System.Console.WriteLine(point.X);
           //    System.Console.WriteLine(point.Y);
           //     if (fieldPosition != 0)
           //     {
           //         short cursorLocation = SetCursor(fieldPosition);

                if (fieldPosition != 0)
                    {
                        short fieldValueLength = (short)fieldValue.Length;
                        if (CopyStringToPS(fieldPosition, fieldValueLength, fieldValue))
                        {
                            ReleaseSession();
                            endConnection();
                            return true;
                        }
                    }
           //     }
                ReleaseSession();
                endConnection();


                throw new EhlApiExceptions("Field population failed:" + fieldName);

            }
            catch (EhlApiExceptions exception)
            {
                return false;
            }
        }

        void CheckException(short returnCode)
        {

            StackTrace stackTrace = new StackTrace();
            // get calling method name
            String calledModule = stackTrace.GetFrame(1).GetMethod().Name;
            if (returnCode == EHLLAPI.APIOK)
            {
                return;
            }
            else if ((returnCode == EHLLAPI.APINOTCONNECTED) || (returnCode == EHLLAPI.APISYSERROR))
            {
                throw new EhlApiError("Module:" + calledModule + ": ReturnCode: " + returnCode);
            }
            else if (returnCode == EHLLAPI.APIPSBUSY || (returnCode != EHLLAPI.APIINHIBITED))
            {
                // System.Console.WriteLine("Host Busy");
                throw new EhlApiExceptions("Host Busy");
            }
            else
            {
                throw new EhlApiExceptions("Module:" + calledModule + ": ReturnCode: " + returnCode);
            }
           
        }
      
    }
}
