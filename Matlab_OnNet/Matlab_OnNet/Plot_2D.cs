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

        string FILE_PATH = "_";
        string SheetName = "_";
        string PlotBlock = "_";
        string PolarOrt = "_";
        string command = "_";
        string CutNumber = "_";
        public static int figure_add = 99;

        MLAppClass matlab;
        OleDbConnection con;
        DataTable dtt;
        OleDbDataAdapter odp;
        DataSet ds;
        Plot_ThreeDim threefun;

        public Plot_TwoDim(string file_path, string page, string block, string orientation, string cut_number)
        {
            FILE_PATH = file_path;
            SheetName = page;
            PlotBlock = block;
            PolarOrt = orientation;
            CutNumber = cut_number;
            matlab = new MLAppClass();
            //threefun = new Plot_ThreeDim(file_path, page, block, 16);
            threefun = new Plot_ThreeDim();
        }
        public void OpenExcel(string path)
        {
            try
            {
                string exlspec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FILE_PATH + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";
                con = new OleDbConnection(exlspec);
                con.Open();
                dtt = new DataTable();
                dtt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                String[] exsh_name = new String[dtt.Rows.Count];

                int i = 0;
                do
                {
                    exsh_name[i] = dtt.Rows[i]["TABLE_NAME"].ToString();
                    i++;
                } while (i < dtt.Rows.Count);
            }
            catch (Exception mes)
            {
                mes = new EvaluateException("" + DateTime.Now + " 2D Error : Can't find the file or can not open file...." + '\t' + "");
                threefun.ErrorLogger(mes.Message);
            }
        }
        public void SelectPage(string page_name)
        {
            try
            {
                odp = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", page_name), con);
                ds = new DataSet();
                odp.Fill(ds, page_name);
            }
            catch (Exception mes)
            {
                mes = new Matlab_OnNet.Plot_ThreeDim.Exception_Handle("" + DateTime.Now + " 2D Error : an error occurs when changed a page or specified a page..." + '\t' + "");
                threefun.ErrorLogger(mes.Message);
            }
        }
        public void Run()
        {
            int Jump2PlotBlock = 0;
            int Jump2Vertical = 0;
            int Jump2Horizontal = 0;
            int OneBlockRows = 0;
            //string exlspec = "_";
            string temp_data = "_";
            string temp_polarscale = "_";
            string temp_reverse_data = "_";
            double insert_list_data = 0;
            double insert_list_scale = 0;
            double insert_list_reverse_data = 0;
            bool Catch_string_useless = false;
            int reverse_cut = 0;
            int cut = 0;

            List<double> DataElements = new List<double>();
            List<double> PolarScaleList = new List<double>();
            List<double> ReverseDataElements = new List<double>();
            List<string> RowsDataList = new List<string>();
            List<string> ColumnDataList = new List<string>();
            /*
            exlspec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FILE_PATH + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";
            OleDbConnection con = new OleDbConnection(exlspec);
            con.Open();
            DataTable dss = new DataTable();
            dss = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            OleDbDataAdapter odp = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), con);
            DataSet dt = new DataSet();
            odp.Fill(dt, SheetName);
            */
            OpenExcel(FILE_PATH);
            SelectPage(SheetName);

            for (int i = 1; i < ds.Tables[0].Columns.Count; i++)
                ColumnDataList.Add(ds.Tables[0].Rows[1][i].ToString());

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                //To get the row position of this string keyin by user.                
                RowsDataList.Add(ds.Tables[0].Rows[i][0].ToString());
                if (ds.Tables[0].Rows[i][0].ToString() == "Phi" && ds.Tables[0].Rows[i + 1][0].ToString() == "Theta")
                    Catch_string_useless = true;
                if (ds.Tables[0].Rows[i][0].ToString() == PlotBlock)
                    Jump2PlotBlock = i;
                if (ds.Tables[0].Rows[i][0].ToString() == "Horizontal Polarization")
                    Jump2Horizontal = i;
                if (ds.Tables[0].Rows[i][0].ToString() == "Vertical Polarization")
                    Jump2Vertical = i;
            }

            Jump2PlotBlock = RowsDataList.IndexOf(PlotBlock);
            OneBlockRows = Math.Abs(Jump2Vertical - Jump2Horizontal); // The row number of one block
            if (Catch_string_useless)
                OneBlockRows = OneBlockRows - 4; //must deduct the strings of "Phi","theta" and polar name and 1 blank             

            if (PolarOrt == "vertical")
                cut = RowsDataList.IndexOf(CutNumber) - 1;
            else if (PolarOrt == "horizontal")
                cut = ColumnDataList.IndexOf(CutNumber) + 1;

            KillList(ColumnDataList);
            KillList(RowsDataList);///kill List capacity

            reverse_cut = cut + ((OneBlockRows - 1) / 2);
            if (reverse_cut >= (OneBlockRows - 1))
                reverse_cut = reverse_cut - (OneBlockRows - 1) + 2;

            if (PolarOrt == "horizontal")
            {
                for (int j = Jump2PlotBlock + 3; j < (OneBlockRows + Jump2PlotBlock) + 4; j++)
                {
                    temp_data = ds.Tables[0].Rows[j][cut].ToString();
                    temp_polarscale = ds.Tables[0].Rows[j][0].ToString();

                    if (temp_data == "" || temp_polarscale == "") break;

                    insert_list_data = Convert.ToDouble(temp_data);
                    insert_list_scale = Convert.ToDouble(temp_polarscale);
                    DataElements.Add(insert_list_data);
                    PolarScaleList.Add(insert_list_scale);
                }
            }
            else if (PolarOrt == "vertical")
            {
                for (int i = Jump2PlotBlock + 1; i < (OneBlockRows + Jump2PlotBlock); i++) //Jump to specified polarization block to read data
                {
                    for (int j = 1; j < ds.Tables[0].Columns.Count; j++)
                    {
                        temp_polarscale = ds.Tables[0].Rows[i][j].ToString();
                        temp_data = ds.Tables[0].Rows[i + cut][j].ToString();
                        temp_reverse_data = ds.Tables[0].Rows[i + reverse_cut][j].ToString();
                        if (temp_data == "" || temp_polarscale == "" || temp_reverse_data == "")
                            goto End;

                        insert_list_scale = Convert.ToDouble(temp_polarscale);
                        insert_list_data = Convert.ToDouble(temp_data);
                        insert_list_reverse_data = Convert.ToDouble(temp_reverse_data);

                        PolarScaleList.Add(insert_list_scale);
                        DataElements.Add(insert_list_data);
                        ReverseDataElements.Add(insert_list_reverse_data);//prepare to reverse data list 
                    }
                }

            End: ; // break 2 loops  

                int temp_interval = 0;
                List<double> DataElementCopy2Reverse = new List<double>(ReverseDataElements);

                DataElementCopy2Reverse.Reverse();
                DataElementCopy2Reverse.Remove(DataElementCopy2Reverse[0]);
                DataElements.AddRange(DataElementCopy2Reverse);

                temp_interval = Convert.ToInt32(Math.Abs(PolarScaleList[2] - PolarScaleList[1]));
                int temp_count = PolarScaleList.Count;
                for (int i = temp_count; i < (temp_count + temp_count) - 1; i++)
                    PolarScaleList.Add((PolarScaleList[i - 1] + temp_interval));

            }

            KillList(ReverseDataElements);//kill list capacity


            for (int i = 0; i < PolarScaleList.Count; i++)
            {
                Console.WriteLine(DataElements[i]);
                Console.WriteLine(PolarScaleList[i]);

                command = "theta(" + (i + 1) + ")=deal(" + PolarScaleList[i] / 180 * Pi + ");";
                matlab.Execute(command);
                command = "data(" + (i + 1) + ")=deal(" + DataElements[i] + ");";
                matlab.Execute(command);
            }
            matlab.Execute("figure(" + figure_add + ")");
            matlab.Execute("polar(theta,data)");
            matlab.Execute("title('SheetName: " + SheetName + "    Block: " + PlotBlock + "')");
            figure_add++;

            KillList(PolarScaleList);
            KillList(DataElements);//kill List capacity

        }//end Run()

        public void KillList<T>(List<T> list_cell)
        {
            list_cell.Clear();
        }
    }
}
