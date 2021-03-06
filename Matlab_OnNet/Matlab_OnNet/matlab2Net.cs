﻿/*
 *  LiteON-ModuleTeam RF-Chamber Matlab_OnNet DLL.
 *  
 *  Copyright (c)  NinoLiu\LiteON , Inc 2012
 * 
 *  Description:
 *    Enter the location and name and the specified block of the excel file, the library will open the excel file to
 *    read all the information, and then provided through matlab function to draw 3D graphics. 
 * 
 * ======================================================================================================
 * History
 * ----------------------------------------------------------------------------------------------------
 * 20120607  | NinoLiu  | 1.0.0  | Release first version for user terminal integration.
 * ----------------------------------------------------------------------------------------------------
 * 20120613  | NinoLiu  | 1.1.0  | Add the ability to read statistical information on excel.
 * ----------------------------------------------------------------------------------------------------
 * 20120614  | NinoLiu  | 1.2.0  | Add the subplot for user review and display special visual effects.
 * ----------------------------------------------------------------------------------------------------
 * 20120619  | NinoLiu  | 1.3.0  | Create a GUI and embed 3D-figure and experiment result infor into this GUI.
 * ----------------------------------------------------------------------------------------------------
 * 20120619  | NinoLiu  | 1.4.0  | Fill a single color for graphics surface. 
 * ----------------------------------------------------------------------------------------------------
 * 20120703  | NinoLiu  | 1.5.0  | Due to swap data of phi and theta on the execel, adjusted the order 
 *                                 to scan data.
 * ----------------------------------------------------------------------------------------------------
 * 20120709  | NinoLiu  | 1.6.0  | Add 2D drawing method and refactory program architecture of 2D & 3D.
 * ----------------------------------------------------------------------------------------------------
 * 20120717  | NinoLiu  | 1.7.0  | Add to scan excel bottom's information and then show on the main-UI,
 *                               | and add the method of killing lists and arrays.
 * ----------------------------------------------------------------------------------------------------
 * 20120720  | NinoLiu  | 1.8.0  | Add error log fill function to record error event.
 * ----------------------------------------------------------------------------------------------------
 * 20120731  | NinoLiu  | 1.9.0  | Add encrypt algorithm(dll) to encrypt/decrypt for TRP/TIP and the other information.
 *                               | Add sheetname index access function to correspond sheet to Global page information.
 * ----------------------------------------------------------------------------------------------------
 * 20120815  | NinoLiu  | 1.10.0 | Add Quick_Sort function for sheet page. 
 * ----------------------------------------------------------------------------------------------------
 * 20120815  | NinoLiu  | 1.11.0 | Add data value shift, add a constant value for all r value. 
 * ----------------------------------------------------------------------------------------------------
 * ======================================================================================================
 */
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
    public class matlab_plot
    {
        MLAppClass matlab;

        public matlab_plot()
        {
            matlab = new MLAppClass();
        }

        public void matlab_set()
        {
            matlab.Visible = 1;
            matlab.Execute("clear");
        }
        public void Plot_2D(string FILE_NAME, string SheetName, string PlotBlock, string orientation, string CutNumber)
        {
            Plot_TwoDim plot2d = new Plot_TwoDim(FILE_NAME, SheetName, PlotBlock, orientation, CutNumber);
            plot2d.Run();
        }

        public void Plot_3D(string FILE_NAME, string SheetName, string PlotBlock, int frame)
        {
            Plot_ThreeDim plot3d = new Plot_ThreeDim(FILE_NAME, SheetName, PlotBlock, frame);
            plot3d.Run();

        }//end Plot_3D

    }//end class
}
