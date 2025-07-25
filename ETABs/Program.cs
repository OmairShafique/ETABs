using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using ETABSv1;

namespace ETABs
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create helper object
            cHelper helper = new Helper();

            // Start ETABS
            cOAPI etabsObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
            etabsObject.ApplicationStart();

            Console.WriteLine("ETABs started");

            // Get model object
            cSapModel sapModel = etabsObject.SapModel;

            Console.WriteLine("\nModel Initializing...");

            // Initialize model
            sapModel.InitializeNewModel(eUnits.kN_m_C);

            // Making the Grid

            sapModel.File.NewGridOnly(1,3,3,2,2,3,3);

            Console.WriteLine("\nGrid Made");

            // Adding Materials

            sapModel.PropMaterial.SetMaterial("Concrete", eMatType.Concrete, -1, "AISC", "Concrete");

            Console.WriteLine("\nAdded Materials");

            // Defining Frame Sections

            sapModel.PropFrame.SetRectangle("Column", "Concrete", 0.3, 0.3);

            Console.WriteLine("\nColumn Property Added");

            sapModel.PropFrame.SetRectangle("Beam", "Concrete", 0.3, 0.5);

            Console.WriteLine("\nBeam Property Added");

           

            

            // Example: Add a single frame section (column) at coordinates (0, 0, 0) to (0, 0, 3)

            //string name = "Default"; // Or string name = ""; if you want ETABS to assign a name
            //int ret = sapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 3, ref name, "Column");


            // Adding Column Frame Sections to the Model through a lopp
            for (int i = 0; i <= 3; i = i + 3)
            {
                for (int j = 0; j <= 3; j = j + 3)
                {
                    // Example: Add a frame section at each grid intersection
                    string frameName = "Col" + (i * 3 + j + 1);
                    sapModel.FrameObj.AddByCoord(i, j, 0, i, j, 3, ref frameName, "Column", "1", "Global");
                    Console.WriteLine($"Added {frameName} at ({i}, {j})");

                    // Refreshing the View
                    sapModel.View.RefreshView();

                    
                }
            }





            // Adding Beam Frame Sections parallel to the x-axis only to the Model through a loop
            for (int i = 0; i <= 3; i = i + 3)
            {
                for (int j = 0; j <= 3; j = j + 3)
                {

                    if (i < 3)
                    {
                        // Example: Add a frame section at each grid intersection
                        string frameName = "Beam" + (i * 3 + j + 1);
                        sapModel.FrameObj.AddByCoord(i, j, 3, i + 3, j, 3, ref frameName, "Beam", "1", "Global");
                        Console.WriteLine($"Added {frameName} at ({i}, {j})");
                        // Refreshing the View
                        sapModel.View.RefreshView();
                        
                    }

                    if (j < 3)
                    {
                        // Example: Add a frame section at each grid intersection
                        string frameName = "Beam" + (i * 3 + j + 1);
                        sapModel.FrameObj.AddByCoord(i, j, 3, i, j + 3, 3, ref frameName, "Beam", "1", "Global");
                        Console.WriteLine($"Added {frameName} at ({i}, {j})");
                        // Refreshing the View
                        sapModel.View.RefreshView();
                        
                    }


                    
                }
            }



            string slabPropName= "Slab150mm";

            double slabThickness = 0.15; // 150 mm slab thickness

            eSlabType slabType = eSlabType.Slab;

            sapModel.PropArea.SetSlab(slabPropName, slabType, eShellType.ShellThin, "Concrete", slabThickness);


            Console.WriteLine($"\nSlab section '{slabPropName}' ({slabThickness}m, {slabType}) defined successfully.");



            // Fix for CS1501: Adjusting the method call to match the correct signature of AddByCoord.  
            // The method expects arrays for coordinates and not individual values.  

            // Define arrays for the coordinates of the slab  
            double[] xCoords = { 0, 3, 3, 0 };
            double[] yCoords = { 0, 0, 3, 3 };
            double[] zCoords = { 3, 3, 3, 3 };

            // Update the method call to use the correct signature  
            sapModel.AreaObj.AddByCoord(4, ref xCoords, ref yCoords, ref zCoords, ref slabPropName, "Slab150mm", "1", "Global");

            sapModel.View.RefreshView(); // Refresh the view to see the changes

            Console.WriteLine($"\nSlab '{slabPropName}' added successfully at coordinates: {string.Join(", ", xCoords.Select(x => x.ToString()))}.");


            // Applying Load Patterns
            sapModel.LoadPatterns.Add("SDL", eLoadPatternType.Dead, 0, true);

            // Now applying the Loads on the Slab sections in gravity direction

            sapModel.SelectObj.All(true);

            int numberOfAreaObjects = 0;
            string[] areaObjectNames = null;

            sapModel.AreaObj.GetNameList(ref numberOfAreaObjects, ref areaObjectNames);

            foreach (string areaObject in areaObjectNames)
            {
                sapModel.AreaObj.SetSelected(areaObject, true);
                sapModel.View.RefreshView();
                // Apply the load to each slab section
                sapModel.AreaObj.SetLoadUniform(areaObject, "SDL", -25, 6);
            }

            
            Console.WriteLine("\nLoad pattern 'SDL' added and applied to slab sections.");

            eCNameType loadCaseName = new eCNameType();

            sapModel.RespCombo.Add("1.4D",0);
            sapModel.RespCombo.SetCaseList("1.4D", ref loadCaseName, "SDL", 1.4);


            Console.WriteLine("\nLoad Combinations Defined");

            // Making a new folder in DOcuments called ETABS_Testing and checking if it exists if it does no worries otherwise make it
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ETABS_Testing");

            // Now save file to this folder
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                Console.WriteLine($"\nCreated folder: {folderPath}");
            }
            else
            {
                // Now saving model to this folder
                sapModel.File.Save(Path.Combine(folderPath, "ETABS_Model.edb"));
                Console.WriteLine($"\nFolder already exists: {folderPath}");
            }

            

            // Analysing the Model
            sapModel.Analyze.RunAnalysis();





            

            Console.WriteLine("\nETABS model created and saved successfully.");
            Console.ReadLine(); // Keeps console open to view output
        }

        
    }
}
