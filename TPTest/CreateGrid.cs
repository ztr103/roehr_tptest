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
            string originString = Xmin + " " + Ymax;
            string yaxisString = Xmin + " " + (Ymax - gridSize);
            string outputPath = pathName + "\\" + fcName + ".shp";

            // Create the geoprocessor object
            IGeoProcessor2 geoProcess;
            geoProcess = (IGeoProcessor2)new GeoProcessor();

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
            geoParams.Add(""); // Output Template - Xmin, Xmax, Ymin, Ymax, OR template dataset
            geoParams.Add("POLYGON"); // Output Geometry Type

            try
            {
                // Run geoprocessor
                geoProcess.Execute("CreateFishnet", geoParams, null);

            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                MessageBox.Show("Error Code: "+ce.ErrorCode.ToString(), "COM Exception", System.Windows.Forms.MessageBoxButtons.OK);
                return;
            }

}

    }
}
