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
    class Plot_ThreeDim
    {
        string FILE_NAME = "_";
        string SheetName = "_";
        string PlotBlock = "_";
        int iframe = 0;
        private const double Pi = 3.1416;
        private string exlspec = string.Empty;
        public static int Figure_acc = 1;
        public static int Sub_figure_num = 2;

        public static int FrameSplitup = 0;

        MLAppClass matlab;

        public Plot_ThreeDim(string file, string page, string block, int fram)
        {
            FILE_NAME = file;
            SheetName = page;
            PlotBlock = block;
            iframe = fram;
            matlab = new MLAppClass();
        }
        public void Run()
        {
            List<string> PolarBlockElements = new List<string>();
            List<string> TestInformation = new List<string>();

            int Jump_2_PlotRow = 0;
            int Jump2Horizontal = 0;
            int Jump2Vertical = 0;
            int temp = 0;
            int column_count = 0;
            string command = "_";

            string exlspec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FILE_NAME + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";
            OleDbConnection con = new OleDbConnection(exlspec);
            con.Open();
            DataTable dss = new DataTable();
            dss = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            OleDbDataAdapter odp = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), con);
            DataSet dt = new DataSet();
            odp.Fill(dt, SheetName);

            for (int i = 0; i < dt.Tables[0].Rows.Count; i++)
            {
                if (dt.Tables[0].Rows[i][0].ToString() == PlotBlock)
                    Jump_2_PlotRow = i; //To get the row position of this string keyin by user.
                if (dt.Tables[0].Rows[i][0].ToString() == "Horizontal Polarization")
                    Jump2Horizontal = i;
                if (dt.Tables[0].Rows[i][0].ToString() == "Vertical Polarization")
                    Jump2Vertical = i;
            }

            for (int i = Jump_2_PlotRow; i < dt.Tables[0].Rows.Count; i++) //Jump to specified polarization block to read data                
            {
                for (int j = 0; j < dt.Tables[0].Columns.Count; j++)
                {
                    if (dt.Tables[0].Rows[i][j].ToString() == "")
                    {
                        if (PolarBlockElements.Contains("Theta") && temp == 0)
                        {
                            temp = 1;
                            column_count = j - 1; // Get column's length under this block                            
                        }

                        continue;
                    }
                    if (column_count > 0 && j > 1 && dt.Tables[0].Rows[i][j].ToString() != "Phi")
                        PolarBlockElements.Add(dt.Tables[0].Rows[i][0].ToString());//fill Row's data into this List container                        
                    PolarBlockElements.Add(dt.Tables[0].Rows[i][j].ToString());//fill test value into List container                    
                }

                //remove these string and search end terminal in the block                
                if (dt.Tables[0].Rows[i][0].ToString() == "" &&
                    PolarBlockElements.Contains(PlotBlock) &&
                    PolarBlockElements.Contains("Phi") &&
                    PolarBlockElements.Contains("Theta"))
                {
                    PolarBlockElements.Remove(PlotBlock);
                    PolarBlockElements.Remove("Phi");
                    PolarBlockElements.Remove("Theta");
                    break;
                }
            }//end for loop

            string[] column_array = new string[column_count];
            string[] row_array = new string[PolarBlockElements.Count];

            column_array = PolarBlockElements.GetRange(0, column_count).ToArray();
            row_array = PolarBlockElements.GetRange(column_count, (PolarBlockElements.Count - column_count)).ToArray();

            List<string> temp_list = new List<string>(column_array);
            //Let temp_list to set automatication                              
            int SerieGeoItem = PolarBlockElements.Count;
            do
            {
                SerieGeoItem = Convert.ToInt32(SerieGeoItem / 2);
                SerieGeoItem--;
                temp_list.AddRange(temp_list); //Copy column value repeat [0 30 60 90 120 ......]
            } while (SerieGeoItem > 0);

            int kg = 0;
            for (int i = 1; i <= (row_array.Length + row_array.Length / 2); i = i + 3)
            {
                temp_list.Insert(i, row_array[kg]);
                temp_list.Insert(i + 1, row_array[kg + 1]);
                kg = kg + 2;
            }

            temp_list.RemoveRange((row_array.Length + row_array.Length / 2), (temp_list.Count - (row_array.Length + row_array.Length / 2)));
            temp_list.Remove("(Unit: dBm)");

            string[] temp_array = new string[temp_list.Count];
            Double[] DataList_2_CoordinateTransformation = new Double[temp_array.Length];

            temp_array = temp_list.GetRange(0, temp_list.Count).ToArray();
            for (int i = 0; i < (temp_list.Count); i++)
            {
                if (temp_array[i] == "")
                    break;
                DataList_2_CoordinateTransformation[i] = Convert.ToDouble(temp_array[i]);
            }
            int interval = Convert.ToInt32(DataList_2_CoordinateTransformation.Length / 3);

            Double[] x = new Double[interval];
            Double[] y = new Double[interval];
            Double[] z = new Double[interval];

            double phi = 0;
            double theta = 0;
            double r = 0;

            int index = 0;
            int ColumnInMatlab = 0;
            int RowInMatlab = 1;


            for (int i = 0; i < (DataList_2_CoordinateTransformation.Length - 3); i = i + 3)
            {

                index = i / 3;

                theta = DataList_2_CoordinateTransformation[i] / 180 * Pi;
                phi = DataList_2_CoordinateTransformation[i + 1] / 180 * Pi;
                r = DataList_2_CoordinateTransformation[i + 2];

                x[index] = r * System.Math.Cos(phi) * System.Math.Sin(theta);
                y[index] = r * System.Math.Sin(phi) * System.Math.Sin(theta);
                z[index] = r * System.Math.Cos(theta);

                //Change sequence from one dimensional(@.NET) to two dimensional(@Matlab) 
                if (index < column_count)
                    ColumnInMatlab = index + 1;
                else
                {
                    RowInMatlab = index / column_count;
                    ColumnInMatlab = index % (column_count * RowInMatlab) + 1;
                    RowInMatlab = RowInMatlab + 1; //row ++ 
                }

                command = "x(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + x[index] + ");";
                matlab.Execute(command);
                command = "y(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + y[index] + ");";
                matlab.Execute(command);
                command = "z(" + RowInMatlab + ", " + ColumnInMatlab + ")=deal(" + z[index] + ");";
                matlab.Execute(command);


            }// end for loop

            //kill array elements
            Array.Clear(x, 0, x.Length);
            Array.Clear(y, 0, y.Length);
            Array.Clear(z, 0, z.Length);
            Array.Clear(DataList_2_CoordinateTransformation, 0, DataList_2_CoordinateTransformation.Length);

            if (FrameSplitup == 0) FrameSplitup = Convert.ToInt32(Math.Ceiling(Math.Sqrt(iframe)));
            Console.WriteLine("Frame: " + FrameSplitup + " ifram: " + iframe + "");

            matlab.Execute("figure('Menubar', 'none');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [20 70 130 20],'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 3][0].ToString() + "');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [140 70 80 20], 'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 3][1].ToString() + "dBm');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [20 45 130 20],'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 2][0].ToString() + "');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [140 45 80 20], 'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 2][1].ToString() + "dBm');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [20 20 130 20],'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 1][0].ToString() + "');");
            matlab.Execute("uicontrol('Style', 'edit', 'Position', [140 20 125 20], 'String', '" + dt.Tables[0].Rows[dt.Tables[0].Rows.Count - 1][1].ToString() + "dBm');");
            matlab.Execute("h   =mesh(x,y,z); xlabel('X-axis');ylabel('Y-axis');zlabel('Z-axis');");
            matlab.Execute("set(h,'edgecolor', [0.2 0.5 0.5], 'FaceColor',[0.99609375 0.99609375 0.55859375]);");
            matlab.Execute("rotate3d on");

            matlab.Execute("title('SheetName: " + SheetName + "    Block: " + PlotBlock + "')");
            matlab.Execute("axis normal;");

            if (Figure_acc == Sub_figure_num) Figure_acc++;

            matlab.Execute("figure(" + Sub_figure_num + ")");
            matlab.Execute("hold on");

            if (Figure_acc > 2) matlab.Execute("subplot(" + FrameSplitup + "," + FrameSplitup + "," + (Figure_acc - 1) + ")");
            else matlab.Execute("subplot(3,3," + Figure_acc + ")");

            //matlab.Execute("subplot("+ FrameSplitup + "," + FrameSplitup + "," + (Figure_acc-1) + ")");
            //else matlab.Execute("subplot(" + FrameSplitup + "," + FrameSplitup + "," + Figure_acc + ")");
            matlab.Execute("s=mesh(x,y,z);rotate3d on");
            matlab.Execute("set(s,'edgecolor', [0.2 0.5 0.5], 'FaceColor',[0.99609375 0.99609375 0.55859375]);");
            matlab.Execute("title('figure(" + Figure_acc + ")')");
            matlab.Execute("hold off");

            Figure_acc++;
            // if (Figure_acc == Sub_figure_num) Figure_acc++;

        }//end Run
    }
}
