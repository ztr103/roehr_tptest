using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geometry;

namespace TPTest
{
    public class CreatePointFC : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public CreatePointFC()
        {
        }

        protected override void OnClick()
        {
            ArcMap.Application.CurrentTool = null;

            //Do Stuff
            CreatePointFClass("c:/TestOutput", "InspectionPoints");
            AddLayer("c:/TestOutput", "InspectionPoints");

        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }

        public void CreatePointFClass(string pathName, string fcName)
        {
            IMxDocument pMxDoc;
            pMxDoc = (IMxDocument)ArcMap.Application.Document;

            IMap pMap;
            pMap = pMxDoc.FocusMap;

            IWorkspaceFactory pWSFactory;
            pWSFactory = new ShapefileWorkspaceFactory();

            IFeatureWorkspace pFWorkspace;
            pFWorkspace = (IFeatureWorkspace)pWSFactory.OpenFromFile(pathName, ArcMap.Application.hWnd);

            IFieldsEdit pFieldsEdit;
            pFieldsEdit = (IFieldsEdit)new Fields();

            IFieldEdit pIDField;
            pIDField = (IFieldEdit)new Field();
            pIDField.Name_2 = "OID";
            pIDField.Type_2 = esriFieldType.esriFieldTypeOID;
            pIDField.Length_2 = 5;

            IFieldEdit pCodeField;
            pCodeField = (IFieldEdit)new Field();
            pCodeField.Name_2 = "Feat_Code";
            pCodeField.Type_2 = esriFieldType.esriFieldTypeInteger;

            IFieldEdit pDescField;
            pDescField = (IFieldEdit)new Field();
            pDescField.Name_2 = "Feat_Desc";
            pDescField.Type_2 = esriFieldType.esriFieldTypeString;
            pDescField.Length_2 = 80;

            IFieldEdit pShapeField;
            pShapeField = (IFieldEdit)new Field();
            pShapeField.Name_2 = "Shape";
            pShapeField.Type_2 = esriFieldType.esriFieldTypeGeometry;

            // Create Spatial Reference
            ISpatialReferenceFactory3 pSpatRefGen;
            pSpatRefGen = (ISpatialReferenceFactory3)new SpatialReferenceEnvironment();
            ISpatialReference pSpatialRef;
            pSpatialRef = pSpatRefGen.CreateSpatialReference(4326);

            //Setup Geometry Definition Settings
            IGeometryDef pGeomDef = new GeometryDef();
            IGeometryDefEdit pGDefEdit = (IGeometryDefEdit)pGeomDef;
            pGDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPoint;
            pGDefEdit.GridCount_2 = 1;
            pGDefEdit.set_GridSize(0, 0.1);
            pGDefEdit.SpatialReference_2 = pSpatialRef;

            //Set Shape Field Geometry Definition
            pShapeField.GeometryDef_2 = pGeomDef;

            //Add Fields to Collection
            pFieldsEdit.AddField(pIDField);
            pFieldsEdit.AddField(pCodeField);
            pFieldsEdit.AddField(pDescField);
            pFieldsEdit.AddField(pShapeField);

            pFWorkspace.CreateFeatureClass(fcName + ".shp", pFieldsEdit, null, null, esriFeatureType.esriFTSimple, "Shape", "");

        }

        public void AddLayer(string pathName, string fcName)
        {
            IMxDocument pMxDoc;
            pMxDoc = (IMxDocument)ArcMap.Application.Document;

            IMap pMap;
            pMap = pMxDoc.FocusMap;

            IFeatureLayer pFLayer;
            pFLayer = new FeatureLayer();

            IWorkspaceFactory pWSFactory;
            pWSFactory = new ShapefileWorkspaceFactory();

            IWorkspace pWorkspace;
            pWorkspace = pWSFactory.OpenFromFile(pathName,ArcMap.Application.hWnd);

            IFeatureWorkspace pFWorkspace;
            pFWorkspace = (IFeatureWorkspace)pWorkspace;

            IFeatureClass pFClass;
            pFClass = pFWorkspace.OpenFeatureClass(fcName+".shp");

            pFLayer.FeatureClass = pFClass;
            pFLayer.Name = fcName;

            pMap.AddLayer(pFLayer);

            pMxDoc.UpdateContents();

        }

    }

}
