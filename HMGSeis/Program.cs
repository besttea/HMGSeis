﻿using System;
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
        private static double x;
        private static double y;
        private static double z;
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

            //get information from  group name "GPtSet_24U" 

            string SelectedPointName = "GPtSet_24R";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_24R = new List<ZkPoints>();
            myPoints_24R = GetPointfromGroup(mySapModel, SelectedPointName);

            SelectedPointName = "GPtSet_14U";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_14U = new List<ZkPoints>();
            myPoints_14U = GetPointfromGroup(mySapModel, SelectedPointName);

            SelectedPointName = "GPtSet_24L";
            Console.WriteLine("Get point information from Group:{0}", SelectedPointName);
            List<ZkPoints> myPoints_24L = new List<ZkPoints>();
            myPoints_24L = GetPointfromGroup(mySapModel, SelectedPointName);

            Console.WriteLine("Finished to get point information ");


            #endregion
            #region Calculating the Points matrix

            //prepare for data of points
            List<ZkPoints> myPoints_border_Up = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_Down = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_left = new List<ZkPoints>();
            List<ZkPoints> myPoints_border_Right = new List<ZkPoints>();

            myPoints_border_Down = CreateBorderPoints(0, 251000, myPoints_24R);

            myPoints_border_left = CreateBorderPoints(250000, 326000, myPoints_14U);

            myPoints_border_Right = CreateBorderPoints(250000, 326000, myPoints_14U);

            for (int i = 0; i < myPoints_border_left.Count; i++)
            {
                myPoints_border_left[i].X = 0;
                myPoints_border_left[i].Z= myPoints_border_Down[0].Z;
            }

            myPoints_border_Up = CreateBorderPoints(0, 260738, myPoints_24R);

            double left_border = 0;double right_border = 325282.2;
            double DeltaLengthofborder = (right_border - left_border)/(myPoints_border_Up.Count-1);
            myPoints_24R.Sort(new ComparePoints_X());

            for (int i = 1; i < myPoints_border_Up.Count; i++)
            {
                myPoints_border_Up[i].X = myPoints_border_Up[i - 1].X+DeltaLengthofborder;
                myPoints_border_Up[i].Y = myPoints_border_Right[myPoints_border_Right.Count - 1].Y;

                myPoints_border_Up[i].Z = interpolationXtoZ(myPoints_border_Up[i].X, myPoints_24R);     
            }
            




            #endregion

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

        private static List<ZkPoints> CreateBorderPoints(double border_Up,double border_Down, List<ZkPoints> myPoints)
        {
            
            int  i = 0;
            List<ZkPoints> myPoints_border=new List<ZkPoints>();
            foreach (ZkPoints tempPoint in myPoints)
            {
                if (tempPoint.X >= border_Up && tempPoint.X < border_Down)
                {
                    tempPoint.Index = i;
                    myPoints_border.Add(tempPoint);
                    i++;
                }

            }

            return myPoints_border;
        }

        private static double interpolationXtoZ(double x1,List<ZkPoints> myPointsList)
        {

            double  z=0;
            ZkPoints UpPoint = new ZkPoints();
            ZkPoints DownPoint = new ZkPoints();

            double up =0,down=0;
            if (x1 >= 0 && x1 <= 325282.2)
            {
                //interpolor z1 from x1 and x3 to get z
                //x0 = vetex1(0)(0): x = point3(0): z0 = vetex1(0)(2): z1 = vetex1(1)(2)
                
                for (int j = 0; j < myPointsList.Count; j++)
                {
                    if (x1 > myPointsList[j].X)
                    {
                        down = myPointsList[j].X;
                        DownPoint = myPointsList[j];//get lower limits of point
                    }
                    else
                    {
                        up = myPointsList[j].X;
                        UpPoint = myPointsList[j];//get Uper limits of point
                        break;
                    }

                }
                //point3(2) = z1 - (z1 - z0) * (x1 - x) / (x1 - x0)

                      z=DownPoint.Z+ (UpPoint.Z-DownPoint.Z)*(x1 - down) / (up - down);

            }

            return z;
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
                double x = y = z = 0;
                ret = mySapModel.PointObj.GetCoordCartesian(ObjectName_Points[i], ref x, ref y, ref z);
                ZkPoints tempPt = new ZkPoints(i, ObjectName_Points[i], x, y, z);
                myPoints.Add(tempPt);
                Console.WriteLine(@"{0},{1},{2},{3},{4}", i, ObjectName_Points[i], x, y, z);
            }
            myPoints.Sort(new ComparePoints_X());
            foreach (ZkPoints tempPt in myPoints)
            {

                Console.WriteLine(@"{0},{1},{2},{3},{4}", tempPt.Index, tempPt.Name, tempPt.X, tempPt.Y, tempPt.Z);
            }
            return myPoints;
        }
    }
}


