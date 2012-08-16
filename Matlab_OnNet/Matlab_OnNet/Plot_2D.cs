using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using MathWorks;
using MathWorks.MATLAB;
using MathWorks.MATLAB.NET.Arrays;
using MathWorks.MATLAB.NET.Utility;
using MLApp;


namespace Matlab_OnNet
{
    class Plot_TwoDim
    {
        private const double Pi = 3.1416;
        private string exlspec = string.Empty;
        MLAppClass matlab;
        string FILE_NAME = "_";
        string SheetName="_";
        string PlotBlock="_";
        string PolarOrt="_";
        string command = "_";
        int Cut = 0;
        public static int figure_add = 99;

        public Plot_TwoDim(string file, string page, string block, string orientation, int arc)
        {
            FILE_NAME = file;
            SheetName = page;
            PlotBlock = block;
            PolarOrt = orientation;
            Cut = arc;
            matlab = new MLAppClass();
        }
        public void Run()
        {
            int Jump2SummationRow = 0;
            int Jump2Vertical = 0;
            int Jump2Horizontal = 0;
            int OneBlockRows = 0;
            string exlspec = "_";
            string temp_data = "_";
            string temp_polarscale = "_";
            double insert_list_data = 0;
            double insert_list_scale = 0;

            List<string> PolarBlockElements = new List<string>();
            List<double> DataElements = new List<double>();
            List<double> PolarScaleList = new List<double>();
            
            exlspec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FILE_NAME + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";

            OleDbConnection con = new OleDbConnection(exlspec);
            con.Open();
            DataTable dss = new DataTable();
            dss = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            OleDbDataAdapter odp = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), con);
            DataSet dt = new DataSet();
            odp.Fill(dt, SheetName);

            for (int i = 0; i < dt.Tables[0].Rows.Count; i++)
            {
                //To get the row position of this string keyin by user.
                Console.WriteLine(dt.Tables[0].Rows[i][0].ToString().Trim());
                if (dt.Tables[0].Rows[i][0].ToString() == PlotBlock)
                    Jump2SummationRow = i; 
                if (dt.Tables[0].Rows[i][0].ToString() == "Horizontal Polarization")
                    Jump2Horizontal = i;
                if (dt.Tables[0].Rows[i][0].ToString() == "Vertical Polarization")
                    Jump2Vertical = i;
            }
            
            OneBlockRows = Math.Abs(Jump2Vertical - Jump2Horizontal);            

            if (PolarOrt == "horizontal")
            {
                for (int j = Jump2SummationRow + 3; j < (OneBlockRows + Jump2SummationRow); j++)
                {
                    temp_data = dt.Tables[0].Rows[j][Cut].ToString();
                    temp_polarscale = dt.Tables[0].Rows[j][0].ToString();
                    
                    if (temp_data == "" || temp_polarscale == "") break;

                    insert_list_data = Convert.ToDouble(temp_data);
                    insert_list_scale = Convert.ToDouble(temp_polarscale);
                    DataElements.Add(insert_list_data);
                    PolarScaleList.Add(insert_list_scale);
                }
            }
            else if (PolarOrt == "vertical")
            {
                for (int i = Jump2SummationRow + 1; i < (OneBlockRows + Jump2SummationRow); i++) //Jump to specified polarization block to read data
                {
                    for (int j = 1; j < dt.Tables[0].Columns.Count; j++)
                    {
                        temp_polarscale = dt.Tables[0].Rows[i][j].ToString();
                        temp_data = dt.Tables[0].Rows[i + Cut][j].ToString();
                
                        if (temp_data == "" || temp_polarscale == "")
                            goto End;

                        insert_list_scale = Convert.ToDouble(temp_polarscale);
                        insert_list_data = Convert.ToDouble(temp_data);
                        PolarScaleList.Add(insert_list_scale);
                        DataElements.Add(insert_list_data);
                    }
                }
            
                End: ; // break 2 loops  

                int temp_interval = 0;
                List<double> PolarScaleCopy2Reverse = new List<double>();
                List<double> DataElementCopy2Reverse = new List<double>(DataElements);

                DataElementCopy2Reverse.Reverse();
                DataElementCopy2Reverse.Remove(DataElementCopy2Reverse[0]);
                DataElements.AddRange(DataElementCopy2Reverse);

                temp_interval = Convert.ToInt32(Math.Abs(PolarScaleList[2] - PolarScaleList[1]));
                int temp_count = PolarScaleList.Count;
                for (int i = temp_count; i < (temp_count + temp_count) - 1; i++)
                    PolarScaleList.Add((PolarScaleList[i - 1] + temp_interval));                
            }

            for (int i = 0; i < PolarScaleList.Count; i++)
            {
                command = "theta(" + (i + 1) + ")=deal(" + PolarScaleList[i] / 180 * Pi + ");";
                matlab.Execute(command);
                command = "data(" + (i + 1) + ")=deal(" + DataElements[i] + ");";
                matlab.Execute(command);
            }            
            matlab.Execute("figure("+ figure_add +")");            
            matlab.Execute("polar(theta,data)");
            matlab.Execute("title('SheetName: " + SheetName + "    Block: " + PlotBlock + "')");
            figure_add++;
        }
    }
}
