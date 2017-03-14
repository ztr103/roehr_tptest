using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using System.Diagnostics;

namespace TPTest
{
    public class CreateGrid : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public CreateGrid()
        {
        }

        protected override void OnClick()
        {
            ArcMap.Application.CurrentTool = null;

            //Do Stuff
            CreateGRID("c:/TestOutput", "InspectionGRID", 0.05, "test_points");
            //AHM
            //CreateGRID("D:/TestOutput", "InspectionGRID", 0.05, "test_points");

        }

        protected override void OnUpdate()
        {
        }

        public void CreateGRID(string pathName, string fcName, double gridSize, string templateLayer)
        {

            //Get Extents of template FC

            IMxDocument pMxDoc;
            pMxDoc = (IMxDocument)ArcMap.Application.Document;

            IMap pMap;
            pMap = pMxDoc.FocusMap;

            IEnumLayer pEnumLayer;
            pEnumLayer = pMap.Layers;

            ILayer pLayer;
            pLayer = pEnumLayer.Next();

            string strName;
            double Xmin = 0;
            double Xmax = 0;
            double Ymin = 0;
            double Ymax = 0;

            while (pLayer != null)
            {
                strName = pLayer.Name;

                if (strName == templateLayer)
                {
                    IEnvelope templateEnv;
                    templateEnv = (IEnvelope)pLayer.AreaOfInterest;

                    //MessageBox.Show(strName, "Template Layer Found", MessageBoxButtons.OK);

                    Xmin = templateEnv.XMin;
                    Xmax = templateEnv.XMax;
                    Ymin = templateEnv.YMin;
                    Ymax = templateEnv.YMax;
                }

                pLayer = pEnumLayer.Next();
            }

            //Populate some values to feed geoprocess parameters
            //string originString = Xmin + " " + Ymax;
            //AHM
            string originString = Xmin + " " + Ymin; //AHM

            //string yaxisString = Xmin + " " + (Ymax - gridSize);
            //AHM
            string yaxisString = Xmin + " " + (Ymax + gridSize); //AHM

            string outputPath = pathName + "\\" + fcName + ".shp";

            //AHM
            string template = Xmin.ToString() + " " + Ymin.ToString() + " " + Xmax.ToString() + " " + Ymax.ToString(); //AHM

            // Create the geoprocessor object
            IGeoProcessor2 geoProcess;
            geoProcess = (IGeoProcessor2)new GeoProcessor();

            //AHM
            geoProcess.OverwriteOutput = true; //AHM

            // Create an IVariantArray to hold the parameter values
            IVariantArray geoParams = new VarArrayClass();

            // Populate parameter values
            geoParams.Add(outputPath); // Output Feature Class
            geoParams.Add(originString); // Output Origin Coordinate
            geoParams.Add(yaxisString); // Output Y Axis Coordinate
            geoParams.Add(gridSize); // Output Cell Width
            geoParams.Add(gridSize); // Output Cell Height
            geoParams.Add("0"); // Number of Rows
            geoParams.Add("0"); // Number of Columns
            geoParams.Add(""); // Output Corner Coord
            geoParams.Add("NO_LABELS"); // Output Lables
            geoParams.Add(template); // Output Template - Xmin, Xmax, Ymin, Ymax, OR template dataset //AHM
            geoParams.Add("POLYGON"); // Output Geometry Type

            try
            {
                // Run geoprocessor
                geoProcess.Execute("CreateFishnet", geoParams, null);

                //AHM
                Debug.WriteLine(geoProcess.GetMessages(0)); //AHM
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                //AHM
                Debug.WriteLine(geoProcess.GetMessages(0)); //AHM
                MessageBox.Show("Error Code: " + ce.ErrorCode.ToString(), "COM Exception", System.Windows.Forms.MessageBoxButtons.OK);
                return;
            }

        }

    }
}
