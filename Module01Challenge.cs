#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace Module01_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Module01Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            // Step 1: Declare a number variable and set it to 250
            int number1 = 250;
            // Step 2: Declare a starting elevation variable and set it to 0
            int startingElevation = 0;
            // Step 3: Declare a floor height variable and set it to 15
            int floorHeight = 15;

            // Create a variable to store the current elevation
            int currentElevation = startingElevation;

            // Create counter-variables to find out how many elements are created to more easily answer the Challenge-Questions
            int numberOfFizz = 0;
            int numberOfBuzz = 0;
            int numberOfFizzBuzz = 0;

            // Step 4: Loop through 1 to the number1-variable
            for (int i = 1; i <= number1; i++)
            {
                // Step 5: Create a level for each number
                Transaction NewLevel = new Transaction(doc);
                NewLevel.Start("New Level");

                Level newLevel = Level.Create(doc, currentElevation);
                newLevel.Name = "Level" + " " + i.ToString();

                NewLevel.Commit();
                NewLevel.Dispose();

                // Step 6: Increment the current elevation by the floor height value
                currentElevation += floorHeight;

                // Step 9: If the number is divisible by 3, create a sheet and name it "FIZZBUZZ_#" (# = current number)
                if (i % 3 == 0 && i % 5 == 0)
                {

                    // get an available title block from document
                    /* Code that did not work
                    FilteredElementCollector collectorTitleBlocks = new FilteredElementCollector(doc);
                    collectorTitleBlocks.OfCategory(BuiltInCategory.OST_TitleBlocks); */

                    // Get an available title block from document
                    FilteredElementCollector collectorTitleBlocks = new FilteredElementCollector(doc);
                    collectorTitleBlocks.OfClass(typeof(FamilySymbol));
                    collectorTitleBlocks.OfCategory(BuiltInCategory.OST_TitleBlocks);

                    Transaction NewSheet = new Transaction(doc);
                    NewSheet.Start("New Sheet and Floor Plan");

                    // create a sheet and name it
                    ViewSheet newSheet = ViewSheet.Create(doc, collectorTitleBlocks.FirstElementId());
                    newSheet.Name = "FIZZBUZZ_" + i.ToString();
                    newSheet.SheetNumber = "A" + i.ToString();

                    // Bonus: Create a floor plan and add it to the sheet by creating a viewport element
                    // find an available floor plan view family type
                    FilteredElementCollector collectorFloorPlanVFT = new FilteredElementCollector(doc);
                    collectorFloorPlanVFT.OfClass(typeof(ViewFamilyType));

                    ViewFamilyType floorPlanVFT = null;
                    foreach (ViewFamilyType curVFT in collectorFloorPlanVFT)
                    {
                        if (curVFT.ViewFamily == ViewFamily.FloorPlan)
                        {
                            floorPlanVFT = curVFT;
                            break;
                        }
                    }

                    // create a floor plan and name it
                    ViewPlan newFloorPlan = ViewPlan.Create(doc, floorPlanVFT.Id, newLevel.Id);
                    newFloorPlan.Name = "FIZZBUZZ_" + i.ToString();

                    // add floor plan to the sheet
                    XYZ insertionPoint = new XYZ(1, 0.5, 0);
                    Viewport newViewport = Viewport.Create(doc, newSheet.Id, newFloorPlan.Id, insertionPoint);

                    NewSheet.Commit();
                    NewSheet.Dispose();
                    numberOfFizzBuzz++;
                    
                }

                // Step 7: If the number is divisible by 3, create a floor plan and name it "FIZZ_#" (# = current number)
                else if (i % 3 == 0)
                {
                    Transaction NewFloor = new Transaction(doc);
                    NewFloor.Start("New Floor Plan");

                    // find an available floor plan view family type
                    FilteredElementCollector collectorFloorPlanVFT = new FilteredElementCollector(doc);
                    collectorFloorPlanVFT.OfClass(typeof(ViewFamilyType));

                    ViewFamilyType floorPlanVFT = null;
                    foreach (ViewFamilyType curVFT in collectorFloorPlanVFT)
                    {
                        if (curVFT.ViewFamily == ViewFamily.FloorPlan)
                        {
                            floorPlanVFT = curVFT;
                            break;
                        }
                    }

                    // create a floor plan and name it
                    ViewPlan newFloorPlan = ViewPlan.Create(doc, floorPlanVFT.Id, newLevel.Id);
                    newFloorPlan.Name = "FIZZ_" + i.ToString();


                    NewFloor.Commit();
                    NewFloor.Dispose();
                    numberOfFizz++;
                }

                // Step 8: If the number is divisible by 5, create a ceiling plan and name it "BUZZ#" (# = current number)
                else if (i % 5 == 0)
                {
                    Transaction NewCeiling = new Transaction(doc);
                    NewCeiling.Start("New Ceiling Plan");

                    // find an available ceiling plan view family type
                    FilteredElementCollector collectorCeilingPlanVFT = new FilteredElementCollector(doc);
                    collectorCeilingPlanVFT.OfClass(typeof(ViewFamilyType));

                    ViewFamilyType ceilingPlanVFT = null;
                    foreach (ViewFamilyType curVFT in collectorCeilingPlanVFT)
                    {
                        if (curVFT.ViewFamily == ViewFamily.CeilingPlan)
                        {
                            ceilingPlanVFT = curVFT;
                            break;
                        }
                    }

                    // create a ceiling plan and name it
                    ViewPlan newCeilingPlan = ViewPlan.Create(doc, ceilingPlanVFT.Id, newLevel.Id);
                    newCeilingPlan.Name = "BUZZ_" + i.ToString();

                    NewCeiling.Commit();
                    NewCeiling.Dispose();
                    numberOfBuzz++;
                }

            }

            // Alert user that the Addin is finished
            TaskDialog.Show("FizzBuzz-Addin", "The Addin was successfully executed!");
            // Give the user an Alert showing how many elements of which type were created
            TaskDialog.Show($"FizzBuzz-Addin-Counter", "created " + numberOfFizzBuzz.ToString() + " FIZZBUZZ-views; " + "created " + numberOfFizz.ToString() + " FIZZ-views; " + "created " + numberOfBuzz.ToString() + " BUZZ-views");

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
