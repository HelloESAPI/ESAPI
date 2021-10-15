using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
  public class Script
  {
    public Script()
    {
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
    {
      // check to make sure a patient is open in eclipse
      if (context.Patient == null)
      {
        MessageBox.Show("Please open a patient");
        return;
      }
      // make sure a plan is pulled in and that its dose is calculated
      if (context.PlanSetup == null || context.PlanSetup.IsDoseValid == false)
      {
        MessageBox.Show("Please open a plan and ensure it has calculated dose");
        return;
      }

      try
      {
        // patient id
        string patientId = context.Patient.Id;
        // current selected plan
        PlanSetup plan = context.PlanSetup;
        //plan.DoseValuePresentation = DoseValuePresentation.Absolute;
        // plan id
        string planId = plan.Id;

        // target volume id
        string targetId = plan.TargetVolumeID;

        // current plan's structure set
        StructureSet structureSet = context.PlanSetup.StructureSet;
        //// check to make sure the plan has a structure set -- since plan has valid dose by this point, there must be a structure set
        //if (structureSet == null)
        //{
        //  MessageBox.Show("Sorry, the selected plan does not have a structure set");
        //  return;
        //}

        

        // target - long way
        //Structure target;
        //foreach (var s in structureSet.Structures)
        //{
        //  if (s.Id == targetId)
        //  {
        //    target = s;
        //  }
        //}


        // target - short way
        Structure target = structureSet.Structures.FirstOrDefault(x => x.Id == targetId);

        // get the body
        Structure body = structureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL");

        if (target == null)
        {
          MessageBox.Show("Please select a target volume");
          return;
        }

        // get the target volume
        double targetVolume = target.Volume;

        // volume at 50% IDL (isodose line) - 50% of the Rx (prescription) dose
        double volumeAt50PctRx = plan.GetVolumeAtDose(body, 
          new DoseValue(50, DoseValue.DoseUnit.Percent), 
          VolumePresentation.AbsoluteCm3);


        // volume at 100% IDL (isodose line) - the Rx (prescription) dose
        double volumeAt100PctRx = plan.GetVolumeAtDose(body,
          new DoseValue(100, DoseValue.DoseUnit.Percent),
          VolumePresentation.AbsoluteCm3);


        // calculate the R100 (basic CI) and R50
        // R100
        double r100 = volumeAt100PctRx / targetVolume;
        // r50
        double r50 = volumeAt50PctRx / targetVolume;


        // display info to user
        MessageBox.Show(
          string.Format("Patient Id:\t\t{0}\n\n" +
                        "Plan Id:\t\t{1}\n\n" +
                        "Max Dose:\t\t{2}\n\n" +
                        "R100 (CI):\t\t{3:F1}\n\n" +
                        "R50 (Gradient):\t{4:F2}",
                        patientId, 
                        planId, 
                        plan.Dose.DoseMax3D.ToString(), 
                        r100,
                        r50));


      }
      catch (Exception ex)
      {
        // string interpolation does not work in single file plugins
        //MessageBox.Show($"Sorry, something went wrong.\n\n{ex.Message}\n\n{ex.StackTrace}");
        MessageBox.Show(string.Format("Sorry, something went wrong.\n\n{0}\n\n{1}", ex.Message, ex.StackTrace));
        throw;
      }

     



    }
  }
}
