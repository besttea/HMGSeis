using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using SAP2000v1;

namespace HMGSeis
{
    class Program
    {
        // private static double x;
        // private static double y;
        // private static double z;
        [STAThread]
        /// <summary>
        /// Sap ApiMain 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            //set the following flag to true to attach to an existing instance of the program
            //otherwise a new instance of the program will be started
            bool AttachToInstance;
            AttachToInstance = false;
            Console.WriteLine("Program start...");
            //set the following flag to true to manually specify the path to SAP2000.exe
            //this allows for a connection to a version of SAP2000 other than the latest installation
            //otherwise the latest installed version of SAP2000 will be launched
            bool SpecifyPath;
            SpecifyPath = true;
            //if the above flag is set to true, specify the path to SAP2000 below
            #region SetSap2000 Path & SDB
            string ProgramPath;

            ProgramPath = "D:\\Program Files\\Computers and Structures\\SAP2000 21\\SAP2000.exe";
            Console.WriteLine("ProgramPath={0}", ProgramPath);


            //full path to the model

            //set it to the desired path of your model

            string ModelDirectory = "G:\\HMGproject\\SapWork201001";
            Console.WriteLine("ModelDirectory={0}", ModelDirectory);
            try
            {

                System.IO.Directory.CreateDirectory(ModelDirectory);

            }

            catch (Exception ex)
            {

                Console.WriteLine("Could not create directory: " + ModelDirectory);

            }

            string ModelName = "shuniu20201001LDSeis.sdb";

            string ModelPath = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName;
            Console.WriteLine("ModelPath={0}", ModelPath);
            #endregion
            //dimension the SapObject as cOAPI type
            cOAPI mySapObject = null;
            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
            int ret = 0;
            #region Prepare for Sap system

            if (AttachToInstance)
            {
                //attach to a running instance of SAP2000
                try
                {   //
                    //Get the Active SapObject
                    //
                    mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("No running instance of the program found or failed to attach.");
                    return;
                }
            }
            else
            {
                //create API helper object
                cHelper myHelper;
                try
                {
                    myHelper = new Helper();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Cannot create an instance of the Helper object");
                    return;
                }
                if (SpecifyPath)
                {
                    //'create an instance of the SapObject from the specified path
                    try
                    {
                        //create SapObject
                        mySapObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);
                        return;
                    }
                }
                else
                {
                    //Create an instance of the SapObject from the latest installed SAP2000
                    try
                    {
                        //create SapObject

                        mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot start a new instance of the program.");
                        return;
                    }
                }
                //Start SAP2000 application
                //eUnits Units = eUnits.kip_in_F; bool visible = true; 
                ret = mySapObject.ApplicationStart();
                Console.WriteLine("Sap2000 Start......");
            }
            ///
            #endregion
            //create SapModel object
            cSapModel mySapModel;
            mySapModel = mySapObject.SapModel;
            //initialize model
            ret = mySapModel.InitializeNewModel((eUnits.kN_m_C));
            //create new blank model
            //ret = mySapModel.File.NewBlank();
            //open a old model
            ret = mySapModel.File.OpenFile(ModelPath);
            #region Get Information from Object of Group
            Console.WriteLine("Get Information from Object of Group....");
            //eItemType itemtype = eItemType.SelectedObjects;
            // ret = mySapModel.SelectObj.PropertyArea("None");
            //
            //get information from  group name "GPtSet_24R" 
            //
            string SelectedPointName = "GPtSet_24R";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_24R = new List<ZkPoints>();
            myPoints_24R = GetPointfromGroup(mySapModel, SelectedPointName);
            //
            //Console.WriteLine("myPoints_24R");
            //PrintList(myPoints_24R);
            //
            //get information from  group name "GPtSet_24L" 
            //
            SelectedPointName = "GPtSet_24L";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_24L = new List<ZkPoints>();
            myPoints_24L = GetPointfromGroup(mySapModel, SelectedPointName);
            //
            //get information from  group name "GPtSet_14U" 
            //
            SelectedPointName = "GPtSet_14U";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_14U = new List<ZkPoints>();
            myPoints_14U = GetPointfromGroup(mySapModel, SelectedPointName);
            //
            //get information from  group name "GPtSet_7D" 
            //
            SelectedPointName = "GPtSet_7D";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_7D = new List<ZkPoints>();
            myPoints_7D = GetPointfromGroup(mySapModel, SelectedPointName);
            Console.WriteLine("Finished to get point information................... ");
            #endregion
            #region Calculating the Points border
            Console.WriteLine("Preper for  Calculating the Points borde...................");
            //
            //prepare for data of points
            List<ZkPoints> myPoints_border_Up = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_Down = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_Left = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_Right = new List<ZkPoints>();
            //
            //
            //
            //bool Left_Right_flag = true,Top_Down_flag=true;
            bool Left_Right_flag = false, Top_Down_flag = false;
            //Define the range of area
            if (Left_Right_flag)
            {
                if (Top_Down_flag)
                {
                    //0,19890.62，0        250738.12,29670.839,2462.44   325282.2.,137255.18，0
                    //Left and Up direction
                    //points set is myPoints_24L and myPoints_7D
                    double XWL_down = 0, XWR_down = Math.Round(250738.12, 0), XHR_top = Math.Round(325282.2, 0) + 1000;
                    myPoints_border_Up = CreateBorderPoints(XWL_down, XWR_down + 1000, myPoints_24R);
                    myPoints_border_Down = CreateBorderPoints(XWL_down, XWR_down + 1000, myPoints_24R);//need modify

                    myPoints_border_Left = CreateBorderPoints(XWR_down - 1000, XHR_top, myPoints_14U);
                    myPoints_border_Right = CreateBorderPoints(XWR_down - 1000, XHR_top, myPoints_14U);//need modify
                }
                else
                {
                    //0,19890.62，0        250738.12,29670.839,2462.44   325282.2.,137255.18，0
                    //Left and down direction
                    //points set is myPoints_24L and myPoints_7D
                    double XWL_down = 0, XWR_down = Math.Round(250738.12, 0), XHR_top = Math.Round(325282.2, 0) + 1000;
                    myPoints_border_Up = CreateBorderPoints(XWL_down, XWR_down + 1000, myPoints_24R);
                    myPoints_border_Down = CreateBorderPoints(XWL_down, XWR_down + 1000, myPoints_24R);//need modify

                    myPoints_border_Left = CreateBorderPoints(XWL_down - 1000, XHR_top, myPoints_14U);
                    myPoints_border_Right = CreateBorderPoints(XWL_down - 1000, XHR_top, myPoints_14U);//need modify
                }
            }
            else
            {
                if (Top_Down_flag)
                {
                    //0,19890.62，0        250738.12,29670.839,2462.44   325282.2.,137255.18，0
                    //Right and Up direction
                    //
                    //points set is myPoints_24L and myPoints_7D
                    //double  XWL_down=0, XWR_down= 251000,  XHR_top = 326000;//Need modify
                    //double XWL_down = 348954, XWR_down = 417000,  XHR_top = 380000;


                    //myPoints_border_Right = CreateBorderPoints(XWR_down, XHR_top, myPoints_14U);
                    //myPoints_border_left = CreateBorderPoints(XWR_down, XHR_top, myPoints_14U);
                    //myPoints_border_Right = CreateBorderPoints(XWL_down, XHR_top, myPoints_7D);//Right and up
                    //myPoints_border_Left = CreateBorderPoints(XWL_down, XHR_top, myPoints_7D);
                }
                else
                {
                    //304577.83,-76564.4        325727.9,-27670.59,   416000.,-18352.27
                    //Right and down direction
                    //points set is myPoints_24L and myPoints_7D
                    double XWL_down = 325727, XWR_down = 417000, XHR_top = 380000;
                    myPoints_border_Up = CreateBorderPoints(XWL_down, XWR_down, myPoints_24L);
                    myPoints_border_Down = CreateBorderPoints(XWL_down, XWR_down, myPoints_24L);//need modify

                    myPoints_border_Left = CreateBorderPoints(0, XWL_down + 1000, myPoints_7D);
                    myPoints_border_Right = CreateBorderPoints(0, XWL_down + 1000, myPoints_7D);//need modify
                }
            }
            //
            //create the points objects of border
            //
            if (Left_Right_flag)
            {

                if (Top_Down_flag)
                {

                    //
                    //the Left Column's of  Y is Divided of Left line.
                    //
                    double left_border = myPoints_border_Down[0].Y;
                    double right_border = myPoints_border_Right[myPoints_border_Right.Count - 1].Y;
                    double DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Right.Count - 1);
                    for (int i = 0; i < myPoints_border_Left.Count; i++)
                    {
                        double x = myPoints_border_Down[0].X;//Need modify
                        double y = left_border + DeltaLengthofborder * i;
                        double z = myPoints_border_Down[0].Z;

                        myPoints_border_Left[i].X = x;
                        myPoints_border_Left[i].Y = y;
                        myPoints_border_Left[i].Z = z;
                        string name = "";
                        //
                        //createing new points of boundary
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Left[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);
                    }

                    //
                    //Right and Top direction.Create myPoints_border_Up
                    //
                    left_border = myPoints_border_Down[0].X;
                    right_border = myPoints_border_Right[myPoints_border_Right.Count - 1].X;//Need modify
                    DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Down.Count - 1);
                    for (int i = 1; i < myPoints_border_Down.Count; i++)
                    {
                        double x = myPoints_border_Down[0].X + DeltaLengthofborder * i;
                        double y = myPoints_border_Right[myPoints_border_Right.Count - 1].Y;//Need modify
                        double z = myPoints_border_Down[myPoints_border_Down.Count - 1].Z;//Need modify
                        string name = "";
                        myPoints_border_Up[i].X = x;
                        myPoints_border_Up[i].Y = y;
                        myPoints_border_Up[i].Z = z;
                        //
                        //createing new points of boundary
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Up[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);

                    }


                }
                else

                {




                }

            }
            else
            {
                //
                //the Right low's y is divided of right line.
                //
                if (Top_Down_flag)
                {
                    double left_border = myPoints_border_Down[myPoints_border_Down.Count - 1].Y;
                    double right_border = myPoints_border_Left[myPoints_border_Left.Count - 1].Y;
                    double DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Right.Count - 1);
                    for (int i = 0; i < myPoints_border_Left.Count; i++)
                    {
                        double x = myPoints_border_Down[myPoints_border_Down.Count - 1].X;//Need modify
                        double y = left_border + DeltaLengthofborder * i;
                        double z = myPoints_border_Down[myPoints_border_Down.Count - 1].Z;

                        myPoints_border_Right[i].X = x;
                        myPoints_border_Right[i].Y = y;
                        myPoints_border_Right[i].Z = z;
                        string name = "";
                        //
                        //createing new points of boundary
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Right[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);
                    }
                }
                else
                {   //
                    //Right and Down direction.creating new points of boundary right line
                    //
                    double left_border = myPoints_border_Left[0].Y;
                    double right_border = myPoints_border_Up[myPoints_border_Up.Count - 1].Y;
                    double DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Left.Count - 1);
                    for (int i = 0; i < myPoints_border_Left.Count; i++)
                    {
                        double x = myPoints_border_Up[myPoints_border_Up.Count - 1].X;//Need modify
                        double y = left_border + DeltaLengthofborder * i;
                        double z = myPoints_border_Up[myPoints_border_Up.Count - 1].Z;

                        myPoints_border_Right[i].X = x;
                        myPoints_border_Right[i].Y = y;
                        myPoints_border_Right[i].Z = z;
                        string name = "";
                        //
                        //createing new points of boundary right line
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Right[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);

                    }

                }


                if (Top_Down_flag)
                {
                    //
                    //Right and Top direction.Create myPoints_border_Up
                    //
                    double left_border = myPoints_border_Left[myPoints_border_Left.Count - 1].X;
                    double right_border = myPoints_border_Down[myPoints_border_Down.Count - 1].X;//Need modify
                    double DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Up.Count - 1);
                    for (int i = 1; i < myPoints_border_Up.Count; i++)
                    {
                        double x = myPoints_border_Left[myPoints_border_Left.Count - 1].X + DeltaLengthofborder * i;
                        double y = myPoints_border_Left[myPoints_border_Left.Count - 1].Y;//Need modify
                        double z = myPoints_border_Down[myPoints_border_Down.Count - 1].Z;//Need modify
                        string name = "";
                        myPoints_border_Up[i].X = x;
                        myPoints_border_Up[i].Y = y;
                        myPoints_border_Up[i].Z = z;
                        //
                        //createing new points of boundary
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Up[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);

                    }
                }
                else
                {
                    //
                    //Right and Down direction.Create myPoints_border_down
                    //
                    double left_border = myPoints_border_Left[0].X;
                    double right_border = myPoints_border_Up[myPoints_border_Up.Count - 1].X;//Need modify
                    double DeltaLengthofborder = (right_border - left_border) / (myPoints_border_Up.Count - 1);
                    for (int i = 0; i < myPoints_border_Up.Count; i++)
                    {
                        double x = myPoints_border_Left[0].X + DeltaLengthofborder * i;
                        double y = myPoints_border_Left[0].Y;
                        double z = myPoints_border_Up[i].Z;
                        string name = "";
                        myPoints_border_Down[i].X = x;
                        myPoints_border_Down[i].Y = y;
                        myPoints_border_Down[i].Z = z;
                        //
                        //Createing New Points Object of boundary
                        // 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_border_Up[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);

                    }
                }
            }
            #endregion
            ///
            ///
            ///
            #region Create Points matrix
            //
            List<ZkPoints> myPoints_Hor = new List<ZkPoints>();
            //
            if (Left_Right_flag)
            {
                if (Top_Down_flag)
                {
                    myPoints_Hor = CreateBorderPoints(0, 250738+1000, myPoints_24R);
                }
                else
                {
                    myPoints_Hor = CreateBorderPoints(0, 210322+1000, myPoints_24L);
                }
                 }
            else
            {
                if (Top_Down_flag)
                {
                    myPoints_Hor = CreateBorderPoints(348954-1000, 416000+1000, myPoints_24R);
                }
                else
                {
                    myPoints_Hor = CreateBorderPoints(325727-1000, 416000+1000, myPoints_24L);
                }
                 }
  
            //Need modify
                //Console.WriteLine("myPoints_Hor:");
                //PrintList(myPoints_Hor);
                //
                //
                //
                List<List<ZkPoints>> PointsListLU = new List<List<ZkPoints>>();
                //
                for (int j = 0; j < myPoints_border_Right.Count; j++)
                {
                    double left_border = myPoints_border_Left[j].X;
                    double right_border = myPoints_border_Right[j].X;
                    double DeltaLength = (right_border - left_border) / (myPoints_border_Up.Count - 1);
                    //
                    //Add the down direction first line of points
                    //
                    List<ZkPoints> layPointsList = new List<ZkPoints>();
                    int lindex; string lname; double lx; double ly; double lz;
                    if (j == 0)
                    {
                        List<ZkPoints> downlayPointsList = new List<ZkPoints>();
                        for (int i = 0; i < myPoints_border_Down.Count; i++)
                        {
                            ZkPoints Point = new ZkPoints(lindex = (myPoints_border_Left.Count - 1) * myPoints_border_Down.Count + i,
                                                 lname = myPoints_border_Down[i].Name,
                                                 lx = myPoints_border_Down[i].X,
                                                 ly = myPoints_border_Down[i].Y,
                                                 lz = myPoints_border_Down[i].Z);
                            downlayPointsList.Add(Point);
                        }
                        PointsListLU.Add(downlayPointsList);
                        continue;
                    }
                //
                //Add Points Set of The Left Column
                //
                lindex = myPoints_border_Left[j].Index;
                lname = myPoints_border_Left[j].Name;
                lx = myPoints_border_Left[j].X;
                ly = myPoints_border_Left[j].Y;
                //lz = myPoints_border_Left[j].Z;
                lz = myPoints_border_Down[0].Z;//change this varibles that would lead to mismatch the vetex of block with the border vetex!!!
                ZkPoints Firstpoints = new ZkPoints(lindex, lname, lx, ly, lz);
                    //
                    //the first and the last data of myPoints_Hor is error
                    //
                    layPointsList.Add(Firstpoints);

                    for (int i = 1; i < myPoints_border_Up.Count; i++)
                    {//create point coordinate of layer 0
                     //
                     //Lv:Xup,Y3,Xd,Yd;
                     //
                        double Xup = myPoints_border_Up[i].X;
                        double Y3 = myPoints_border_Up[i].Y;
                        double Xd = myPoints_border_Down[i].X;
                        double Yd = myPoints_border_Down[i].Y;
                        double Kv = K(Xup, Y3, Xd, Yd), dv = D(Xup, Y3, Xd, Yd);
                        //
                        //Lu:X0,Y4,Xr,Yr
                        //

                    //    double X0 = myPoints_border_Down[0].X;
                        double X0 = myPoints_border_Left[j].X;
                        double Y4 = myPoints_border_Left[j].Y;
                        double Xr = myPoints_border_Right[j].X;
                        double Yr = myPoints_border_Right[j].Y;
                        double Ku = K(X0, Y4, Xr, Yr), du = D(X0, Y4, Xr, Yr);
                    
                        double x = (du - dv) / (Kv - Ku);
                        double y = Kv * (du - dv) / (Kv - Ku) + dv; //interaction of Lu and Lv

                        double z = myPoints_border_Down[i].Z;
                        if (Math.Abs(Kv) < 1e-6) x = Xd; y = Ku * x + du;

                        myPoints_Hor[i].X = x;
                        myPoints_Hor[i].Y = y;
                        myPoints_Hor[i].Z = z;
                        ///
                        ///
                        ///
                        string name = ""; 
                        ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                        myPoints_Hor[i].Name = name;
                        ret = mySapModel.PointObj.AddCartesian(x, y, z - 11651.0, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -19151, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -27251, ref name);
                        ret = mySapModel.PointObj.AddCartesian(x, y, -36531, ref name);
                        ///
                        myPoints_Hor[i].Name = name;
                    //
                    lindex = i + j * myPoints_border_Down.Count;
                    lname = name;
                    lx = x;
                    ly = y;
                    lz = z;
                    ZkPoints Point = new ZkPoints(lindex, lname, lx, ly, lz);
                        layPointsList.Add(Point);
                    }

                    lindex = myPoints_Hor[myPoints_Hor.Count - 1].Index;
                    lname =myPoints_Hor[myPoints_Hor.Count - 1].Name;
                    lx = myPoints_Hor[myPoints_Hor.Count - 1].X;
                    ly = myPoints_Hor[myPoints_Hor.Count - 1].Y;
                    lz= myPoints_Hor[myPoints_Hor.Count - 1].Z;
                    ZkPoints Endpoints = new ZkPoints(lindex ,lname ,lx ,ly,lz );
                    layPointsList.Add(Endpoints);

                    PointsListLU.Add(layPointsList);
                }

                //Up line
                //add Up line of point
                List<ZkPoints> uplayPointsList = new List<ZkPoints>();
                for (int i = 0; i < myPoints_border_Up.Count; i++)
                {
                int lindex= lindex = (myPoints_border_Left.Count - 1) * myPoints_border_Down.Count + i;
                string lname = myPoints_border_Up[i].Name;
                double lx = myPoints_border_Up[i].X;
                double ly = myPoints_border_Up[i].Y;
                double lz= myPoints_border_Up[i].Z;
                ZkPoints Point = new ZkPoints(lindex,lname ,lx,ly,lz) ;
                uplayPointsList.Add(Point);
                }
                PointsListLU.Add(uplayPointsList);
                //ret = mySapModel.View.RefreshView(0, false);
            #endregion
            #region Add point Matrix
            ///
            ///point matrix
            ///
            List<List<Solid>> SolidMatrix = new List<List<Solid>>();

                for (int j = 0; j < myPoints_border_Left.Count - 1; j++)
                {

                    List<Solid> SolidList = new List<Solid>();
                    for (int i = 0; i < myPoints_border_Down.Count - 1; i++)
                    {
                        Solid tempSolid = new Solid();
                        double[] X = new double[8];
                        double[] Y = new double[8];
                        double[] Z = new double[8];
                        string[] Name = new string[8];
                        string BoxName = "";
                        //
                        //Lay Up
                        //
                        for (int k = 0; k < 2; k++)
                        {
                            X[k] = PointsListLU[k + j][i].X;
                            Y[k] = PointsListLU[k + j][i].Y;
                            Z[k] = PointsListLU[k + j][i].Z;
                        }
                        for (int k = 2; k < 4; k++)
                        {
                            X[k] = PointsListLU[k - 2 + j][i + 1].X;
                            Y[k] = PointsListLU[k - 2 + j][i + 1].Y;
                            Z[k] = PointsListLU[k - 2 + j][i + 1].Z;
                        }
                        //
                        //
                        //
                        for (int k = 4; k < 6; k++)
                        {
                            X[k] = PointsListLU[k - 4 + j][i].X;
                            Y[k] = PointsListLU[k - 4 + j][i].Y;
                            Z[k] = PointsListLU[k - 4 + j][i].Z - 11651;
                        }
                        for (int k = 6; k < 8; k++)
                        {
                            X[k] = PointsListLU[k - 6 + j][i + 1].X;
                            Y[k] = PointsListLU[k - 6 + j][i + 1].Y;
                            Z[k] = PointsListLU[k - 6 + j][i + 1].Z - 11651;
                        }

                        ret = mySapModel.SolidObj.AddByCoord(ref X,
                                                             ref Y,
                                                             ref Z,
                                                             ref BoxName);
                        tempSolid.Name = BoxName;
                        tempSolid.X = X;
                        tempSolid.Y = Y;
                        tempSolid.Z = Z;
                        SolidList.Add(tempSolid);
                        ret = mySapModel.View.RefreshView(0, false);
                    ///
                        for (int m = 0; m < 8; m++)
                        {
                            if (m < 4) Z[m] = Z[m] - 11651;
                            else Z[m] = -19151;

                        }
                        ret = mySapModel.SolidObj.AddByCoord(ref X,
                                                             ref Y,
                                                             ref Z,
                                                             ref BoxName);
                        for (int m = 0; m < 8; m++)
                        {
                            if (m < 4) Z[m] = -19151;
                            else Z[m] = -27251;

                        }
                        ret = mySapModel.SolidObj.AddByCoord(ref X,
                                                             ref Y,
                                                             ref Z,
                                                             ref BoxName);
                        for (int m = 0; m < 8; m++)
                        {
                            if (m < 4) Z[m] = -27251;
                            else Z[m] = -36531;

                        }
                        ret = mySapModel.SolidObj.AddByCoord(ref X,
                                                             ref Y,
                                                             ref Z,
                                                             ref BoxName);

                    }
                    SolidMatrix.Add(SolidList);
                }
                #endregion
                ret = mySapModel.View.RefreshView(0, false);
                ///
                ///Write to CAD
                ///
                Console.WriteLine("Write to AutoCAD.....");
                #region Write to AutoCAD
                using (AutoCADConnector connector = new AutoCADConnector())
                {
                    Console.WriteLine(connector.Application.ActiveDocument.Name);
                    AcadApplication CadApp = connector.Application;
                    AcadDocument CadDoc = connector.Application.ActiveDocument;
                    AcadDatabase db = CadDoc.Database;
                    AcadModelSpace CadSpace = CadDoc.ModelSpace;

                    CadApp.Visible = true;
                    //      //    //使AutoCAD可见(d)在按钮的消息处理函数中加入：

                    ////      var face2 = CadSpace.Add3DFace(point1, point2, point3,Point4,point4);



                    CadApp.Application.Update();
                    //     //更新显示



                    Console.WriteLine("Write to AutoCAD.....End");
                    #endregion
                    Console.WriteLine("Points Group:{0},Reading.....End!", SelectedPointName);


                    ret = mySapModel.SelectObj.ClearSelection();

                    //switch to kN-m units

                    ret = mySapModel.SetPresentUnits(eUnits.kN_m_C);

                    /////////////////////////////////////////

                    #region RunAnalysis
                    //
                    ///////////////////////////////////////


                    //save model

                    // ret = mySapModel.File.Save(ModelPath);



                    //run model (this will create the analysis model)

                    //ret = mySapModel.Analyze.RunAnalysis();

                    #endregion

                    #region SAP2000 results
                    //close sap2000
                    ret = mySapObject.ApplicationExit(true);




                    //fill SAP2000 result strings


                    Console.ReadKey();
                    #endregion

                }           
        }
        private static List<ZkPoints> CreateBorderPoints(double border_Low, 
                                                        double border_Up, 
                                                        List<ZkPoints> myPoints)
        {
            
            int  i ;i = 0;
            List<ZkPoints> myPoints_border=new List<ZkPoints>();
            
            foreach (ZkPoints tempPoint in myPoints)
            {
                if ( tempPoint.X >= border_Low && tempPoint.X <= border_Up)
                {
                    ZkPoints myPoint = new ZkPoints(i,
                                                    tempPoint.Name,
                                                    tempPoint.X,
                                                    tempPoint.Y,
                                                    tempPoint.Z);//create a new class instance !
                    myPoints_border.Add(myPoint);
                    i++;
                }

            }

            return myPoints_border;
        }
        private static double interpolationXtoX(double x1,
                                                List<ZkPoints> myPointsList_Up,
                                                List<ZkPoints> myPointsList_Left,
                                                List<ZkPoints> myPointsList_Right,
                                                List<ZkPoints> myPointsList_Low)
        {

            double x = 0;
            ZkPoints UpPoint = new ZkPoints();
            ZkPoints DownPoint = new ZkPoints();

            double up = 0, down = 0;
            if (x1 >= 0 && x1 <= 325282.2)
            {
                //interpolate z1 from x1 and x3 to get z
                for (int j = 0; j < myPointsList_Low.Count; j++)
                {
                    if (x1 > myPointsList_Low[j].X)
                    {
                        down = myPointsList_Low[j].X;
                        DownPoint = myPointsList_Low[j];//get lower limits of point
                    }
                    else
                    {
                        up = myPointsList_Low[j].X;
                        UpPoint = myPointsList_Low[j];//get Uper limits of point
                        break;
                    }

                }
                //point3(2) = z1 - (z1 - z0) * (x1 - x) / (x1 - x0)

                x = DownPoint.Z + (UpPoint.Z - DownPoint.Z) * (x1 - down) / (up - down);

            }

            return x;
        }
        private static double K(double x1,double y1,double x2,double y2)
        {
            if (Math.Abs(x2 - x1) < 1e-6)
            { return 0; }
            else
            { return (y2 - y1) / (x2 - x1); }

        }
        private static double D(double x1,double y1,double x2,double y2)
        {
            return y2-K(x1,y1,x2,y2)*x2;
        }

        private static double interpolationXtoZ(double X,int m,int n,
                                                List<ZkPoints> myPointsList_Up, 
                                                List<ZkPoints> myPointsList_Left, 
                                                List<ZkPoints> myPointsList_Right, 
                                                List<ZkPoints> myPointsList_Low)
        {
            double  Z=0;
            ZkPoints UpPoint = new ZkPoints();
            ZkPoints LowPoint = new ZkPoints();

            double X1 =0,X0=0,Y1=0,Y0=0,Z1=0,Z0=0;
            if (X >= 0 && X <= 325282.2)
            {
                //interpolate z1 from x1 and x3 to get z
                for (int j = 0; j < myPointsList_Low.Count; j++)
                {
                    if (X > myPointsList_Low[j].X)
                    {
                        X0 = myPointsList_Low[j].X;  //X0                      
                        Z0 = myPointsList_Low[j].Z;//Z0
                        LowPoint = myPointsList_Low[j];//get lower limits of point
                    }
                    else
                    {
                        X1 = myPointsList_Low[j].X;   //X1                    
                        Z1 = myPointsList_Low[j].Z;   //Z1
                        UpPoint = myPointsList_Low[j];//get Uper limits of point
                        break;
                    }

                }
                //point3(2) = z1 - (z1 - z0) * (x1 - x) / (x1 - x0)
                Y0 = myPointsList_Left[m].Y;
                Y1 = myPointsList_Right[m].Y;

                Z =Z0+ (Z1-Z0)*(X1 - X) / (X1 - X0);

            }

            return Z;
        }

        /// <summary>
        /// Get Point information from Group
        /// </summary>
        /// <param name="i"></param>
        /// <param name="ret"></param>
        /// <param name="mySapModel"></param>
        /// <param name="SelectedPointName"></param>
        private static List<ZkPoints> GetPointfromGroup( cSapModel mySapModel, string SelectedPointName)
        {
            int NumberSelected_Point = 0;
            int[] ObjectType_Point = new int[60];
            string[] ObjectName_Points = new string[60];
            List<ZkPoints> myPoints  = new List<ZkPoints>();
            
            int ret;

            ret = mySapModel.SelectObj.ClearSelection();
            ret = mySapModel.SelectObj.Group(SelectedPointName);
            ret = mySapModel.SelectObj.GetSelected(ref NumberSelected_Point, ref ObjectType_Point, ref ObjectName_Points);
         
            for (int i = 0; i < NumberSelected_Point; i++)
            {
                double x = 0;double y = 0;double z = 0;
                ret = mySapModel.PointObj.GetCoordCartesian(ObjectName_Points[i], ref x, ref y, ref z);
                ZkPoints tempPt = new ZkPoints(i, ObjectName_Points[i], x, y, z);//create a new instance
                myPoints.Add(tempPt);
               // Console.WriteLine(@"{0},{1},{2},{3},{4}", i, ObjectName_Points[i], x, y, z);
            }
            myPoints.Sort(new ComparePoints_X());
           // PrintList(myPoints);
            return myPoints;
        }

        private static void PrintList(List<ZkPoints> myPointsList)
        {

            foreach (ZkPoints tempPt in myPointsList)
            {
                Console.WriteLine(@"{0},{1},{2},{3},{4}", tempPt.Index, tempPt.Name, tempPt.X, tempPt.Y, tempPt.Z);
            }
        }
    }
}


