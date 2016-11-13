using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MicroFocus.RDE.Framework;
using System.Data;

namespace MainframeAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            System.Console.WriteLine("Welcome to the Mainframe Automation");
            string excelLocation = ApiSettings.Default.ExcelLocation;
            EhlApi ehlApiObj = new EhlApi();
            ImportExcel midExcel = new ImportExcel(excelLocation, ".xlsx");
            DataTable excelData = midExcel.ReadFile("Addition");
            if(ehlApiObj.ExecuteCommand("LV", "MAIN MENU"))
            {
             ehlApiObj.ExecuteCommand("A", "VEHICLE MAINTENANCE");
            }
            if (excelData != null)
            {
                foreach (DataRow dr in excelData.Rows)
                {
                    String policyNumber = dr[0].ToString();
                    String RegNo = dr[1].ToString();
                    String vehicleType = dr[2].ToString();
                    String make = dr[3].ToString();
                    String model = dr[4].ToString(); 
                    String cc = dr[5].ToString();
                    String weight = dr[6].ToString();
                    String coverType = dr[7].ToString();
                    String clientName = dr[8].ToString();
                    String vehicleAdditionDate = dr[9].ToString();
                    String vehicleOffDate = dr[10].ToString();
                  
                            if (ehlApiObj.SearchFieldAndPopulate("Policy", policyNumber))
                            {
                                if (ehlApiObj.SearchFieldAndPopulate("Type of Vehicle", vehicleType))
                                {
                                    if (ehlApiObj.SearchFieldAndPopulate("Registration Number", RegNo))
                                    {
                                        if (ehlApiObj.ExecuteCommand("ENTER", "ADD"))
                                        {
                                            if (ehlApiObj.ExecuteCommand("ENTER", "ADD"))
                                            {
                                                ehlApiObj.ExecuteCommand("PF05", "DVLA Vehicle Data");
                                                ehlApiObj.SearchFieldAndPopulate("Cover", coverType);
                                                ehlApiObj.SearchFieldAndPopulate("Make", make);
                                                ehlApiObj.SearchFieldAndPopulate("Model", model);
                                                ehlApiObj.SearchFieldAndPopulate("Capacity", cc);
                                                ehlApiObj.SearchFieldAndPopulate("On date", vehicleAdditionDate);
                                                ehlApiObj.SearchFieldAndPopulate("Off date", vehicleOffDate);
                                                ehlApiObj.ExecuteCommand("ENTER", "ADD PRIVATE CAR");
                                                if (!ehlApiObj.ExecuteCommand("PF05", "PRESS UPDATE KEY"))
                                                {
                                                    string status = ehlApiObj.GetStatus();
                                                    System.Console.WriteLine("Couldnt add:" + policyNumber + " Vehicle:" + RegNo);
                                                    dr[11] = "Failed due to: "+ status;
                                                   
                                                    ehlApiObj.ExecuteCommand("PF03", "");
                                                    ehlApiObj.ExecuteCommand("A", "VEHICLE MAINTENANCE");
                                                }
                                                else
                                                {
                                                    dr[11] = "Success";
                                                }

                                            }
                                       

                                        }
                                    }
                                }
                            
                       
                    }
                }

                midExcel.UpdateExcel(excelData, "Result");
                System.Console.WriteLine("FINISHED EXECUTION.");
                System.Console.WriteLine("Press any key to exit......");
            }
            else
            {
                System.Console.WriteLine("Excel cant be opened");
            }

          

              
            System.Console.ReadLine();






        }
    }
}