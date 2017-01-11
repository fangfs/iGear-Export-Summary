using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace iGear_Export_Summary
{
    class Program
    {

        static System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(iGear_Export_Summary.Properties.Settings.Default.sqlConnection);
        static System.Data.SqlClient.SqlConnection sqlConnection1 = new System.Data.SqlClient.SqlConnection(iGear_Export_Summary.Properties.Settings.Default.sqlConnection);
        static string sSource = "iGear";
        static string sLog = "Dana";


        static void Main(string[] args)
        {

            //make the event log
            //ensure event log is there and working...
            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);

            //run exports through a test - return based on results
            //0 = nothing to do
            //1 = an export was done
            //2 = an error occured
            int actTest = exTract();

            if (actTest == 1)
            {
                //ok first part passed - do the next
                EventLog.WriteEntry(sSource, "Done", EventLogEntryType.Information, 301);
            }
            else if (actTest == 0)
            {
                //do nothing the first part failed.
                EventLog.WriteEntry(sSource, "Nothing to be done", EventLogEntryType.Information, 302);
            }

        }

        static int exTract()
        {
            //take the date and time - if I am first of mionth, do the previous month
            //if I am sunday do last week  - both if sunday is first of the month
            DateTime dteStart = DateTime.Now;
            string DayOfWeek = DateTime.Now.DayOfWeek.ToString();
            //set up filename type
            string filename = "";
            //stop is always today
            DateTime dteStop = DateTime.Now.AddDays(-1);

            //open database ready to read
            sqlConnection.Open();
            sqlConnection1.Open();

            SqlDataReader reader;
            SqlDataReader reader1;

            try
            {
                //set default to 0 - as most times this will produce nothing
                int returnvalue = 0;

                if (DateTime.Now.Day == 1)
                {
                    //first of the month
                    dteStart = DateTime.Now.AddMonths(-1);
                    filename = "Month";

                    //R1276932 - MSW - 11th Jan 2017
                    //workorderhistory changed from workorder
                    string CommandText = " select a_ProductionDate, b.supplierpart MAT_NUM, '' OLD_MAT_REF, '3602' PLANT, 'SOEM' SLOC,'A' BATCH_NUM, " +
                         "a.serialnumber HU_NUM, c.PartSerialNumber SERIAL, c.scantimestamp , d.Number 'WO #' " +
                         "from pallet a " +
                         "inner join PartDefinition b on a.PartDefinition_ID = b.ID " +
                         "inner join PalletDetail c on a.id = c.Pallet_ID " +
                         "inner join WorkOrderHistory d on a.WorkOrderHistory_ID = d.ID " +
                         "where Status_ID = 1 " +
                         "and c.scantimestamp > '" + dteStart.ToString("yyyy-MM-dd HH:mm:ss") + "' and c.scantimestamp <= '" + dteStop.ToString("yyyy-MM-dd HH:mm:ss") + "' " +
                         "order by c.ScanTimestamp";

                    //run the command
                    reader = new SqlCommand(CommandText, sqlConnection).ExecuteReader();

                    if (reader.HasRows)
                    {
                        string streamFile = dteStart.ToString("yyyyMMdd");
                        using (StreamWriter sw = new StreamWriter(iGear_Export_Summary.Properties.Settings.Default.streamPath + filename + streamFile + ".txt"))
                        {
                            //write header
                            sw.WriteLine("Production Date\tMAT_NUM\tOLD_MAT_REF\tPLANT\tSLOC\tBATCH_NUM\tHU_NUM\tSERIAL\tscantimestamp\tWO #");
                            while (reader.Read())
                            {
                                //write to data file
                                //streamwrite the output to a timestamped file
                                sw.WriteLine(reader["a_ProductionDate"] + "\t" + reader["MAT_NUM"] + "\t" + reader["OLD_MAT_REF"] + "\t" + reader["PLANT"] + "\t" + reader["SLOC"] + "\t" +
                                reader["BATCH_NUM"] + "\t" + reader["HU_NUM"] + "\t" + reader["SERIAL"] + "\t" + reader["scantimestamp"] + "\t" + reader["WO #"]);
                            }
                        }
                    }
                    //shut the reader - just in case
                    returnvalue = 1;
                }

                //if I get here i have been succesfull - carry on the check the day of the week.
                if (DayOfWeek == "Sunday")
                {
                    //first of the month
                    dteStart = DateTime.Now.AddDays(-8);
                    filename = "Week";

                    //R1276932 - MSW - 11th Jan 2017
                    //workorderhistory changed from workorder
                    string CommandText1 = " select a_ProductionDate, b.supplierpart MAT_NUM, '' OLD_MAT_REF, '3602' PLANT, 'SOEM' SLOC,'A' BATCH_NUM, " +
                        "a.serialnumber HU_NUM, c.PartSerialNumber SERIAL, c.scantimestamp , d.Number 'WO #' " +
                        "from pallet a " +
                        "inner join PartDefinition b on a.PartDefinition_ID = b.ID " +
                        "inner join PalletDetail c on a.id = c.Pallet_ID " +
                        "inner join WorkOrderHistory d on a.WorkOrderHistory_ID = d.ID " +
                        "where Status_ID = 1 " +
                        "and c.scantimestamp > '" + dteStart.ToString("yyyy-MM-dd HH:mm:ss") + "' and c.scantimestamp <= '" + dteStop.ToString("yyyy-MM-dd HH:mm:ss") + "' " +
                        "order by c.ScanTimestamp";

                    //run the command
                    reader1 = new SqlCommand(CommandText1, sqlConnection1).ExecuteReader();

                    if (reader1.HasRows)
                    {
                        string streamFile = dteStart.ToString("yyyyMMdd");
                        using (StreamWriter sw = new StreamWriter(iGear_Export_Summary.Properties.Settings.Default.streamPath + filename + streamFile + ".txt"))
                        {
                            //write header
                            sw.WriteLine("Production Date\tMAT_NUM\tOLD_MAT_REF\tPLANT\tSLOC\tBATCH_NUM\tHU_NUM\tSERIAL\tscantimestamp\tWO #");
                            while (reader1.Read())
                            {
                                //write to data file
                                //streamwrite the output to a timestamped file
                                sw.WriteLine(reader1["a_ProductionDate"] + "\t" + reader1["MAT_NUM"] + "\t" + reader1["OLD_MAT_REF"] + "\t" + reader1["PLANT"] + "\t" + reader1["SLOC"] + "\t" +
                                reader1["BATCH_NUM"] + "\t" + reader1["HU_NUM"] + "\t" + reader1["SERIAL"] + "\t" + reader1["scantimestamp"] + "\t" + reader1["WO #"]);
                            }
                        }
                    }
                    //shut the reader just in case
                    returnvalue = 1;
                }

                //if I get here i have been succesfull - finshed now
                sqlConnection.Close();
                sqlConnection1.Close();
                return returnvalue;
            }
            catch (Exception e)
            {
                //write error log, then return false
                EventLog.WriteEntry(sSource, e.Message, EventLogEntryType.Error, 200);
                sqlConnection.Close();
                sqlConnection1.Close();
                return 2;
            }

        }
    }

}

