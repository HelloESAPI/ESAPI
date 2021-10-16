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
  /// <summary>
  /// A simple script to calculate the R100 and R50 of a plan's target volume
  /// </summary>
  public class Script
  {
    public Script()
    {
    }


    [MethodImpl(MethodImplOptions.NoInlining)]
    public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
    {

      // try to calculate the conformity indices and message user if it fails
      try
      {
        // validate the patient is pulled in and the plan has dose
        ValidatePatientAndPlanDose(context);

        // current selected plan
        PlanSetup plan = context.PlanSetup;

        // target volume id
        string targetId = plan.TargetVolumeID;

        // current plan's structure set
        StructureSet structureSet = GetStructureSet(plan);

        // get the plan target volume
        Structure target = GetPlanTargetVolume(targetId, structureSet);
        // validate a target has been selected in the plan
        ValidateTargetVolume(target);

        // get the body structure
        Structure body = GetBodyStructure(structureSet); // won't validate body since plan dose has already been validated -> plan must have body to have dose

        // R100
        double r100 = CalculateR100(plan, body, target);
        // R50
        double r50 = CalculateR50(plan, body, target);

        // display results to the user
        DisplayConformityStats(context.Patient, plan, r100, r50);

      }
      catch (Exception ex)
      {
        // string interpolation does not work in single file plugins
        //MessageBox.Show($"Sorry, something went wrong.\n\n{ex.Message}\n\n{ex.StackTrace}");
        MessageBox.Show(string.Format("Sorry, something went wrong.\n\n{0}\n\n{1}", ex.Message, ex.StackTrace));
        throw;
      }





      #region original code

      //// check to make sure a patient is open in eclipse
      //if (context.Patient == null)
      //{
      //  MessageBox.Show("Please open a patient");
      //  return;
      //}
      //// make sure a plan is pulled in and that its dose is calculated
      //if (context.PlanSetup == null || context.PlanSetup.IsDoseValid == false)
      //{
      //  MessageBox.Show("Please open a plan and ensure it has calculated dose");
      //  return;
      //}

      //try
      //{
      //  // patient id
      //  string patientId = context.Patient.Id;
      //  // current selected plan
      //  PlanSetup plan = context.PlanSetup;
      //  //plan.DoseValuePresentation = DoseValuePresentation.Absolute;
      //  // plan id
      //  string planId = plan.Id;

      //  // target volume id
      //  string targetId = plan.TargetVolumeID;

      //  // current plan's structure set
      //  StructureSet structureSet = context.PlanSetup.StructureSet;
      //  //// check to make sure the plan has a structure set -- since plan has valid dose by this point, there must be a structure set
      //  //if (structureSet == null)
      //  //{
      //  //  MessageBox.Show("Sorry, the selected plan does not have a structure set");
      //  //  return;
      //  //}



      //  // target - long way
      //  //Structure target;
      //  //foreach (var s in structureSet.Structures)
      //  //{
      //  //  if (s.Id == targetId)
      //  //  {
      //  //    target = s;
      //  //  }
      //  //}


      //  // target - short way
      //  Structure target = structureSet.Structures.FirstOrDefault(x => x.Id == targetId);

      //  // get the body
      //  Structure body = structureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL");

      //  if (target == null)
      //  {
      //    MessageBox.Show("Please select a target volume");
      //    return;
      //  }

      //  // get the target volume
      //  double targetVolume = target.Volume;

      //  // volume at 50% IDL (isodose line) - 50% of the Rx (prescription) dose
      //  double volumeAt50PctRx = plan.GetVolumeAtDose(body,
      //    new DoseValue(50, DoseValue.DoseUnit.Percent),
      //    VolumePresentation.AbsoluteCm3);


      //  // volume at 100% IDL (isodose line) - the Rx (prescription) dose
      //  double volumeAt100PctRx = plan.GetVolumeAtDose(body,
      //    new DoseValue(100, DoseValue.DoseUnit.Percent),
      //    VolumePresentation.AbsoluteCm3);


      //  // calculate the R100 (basic CI) and R50
      //  // R100
      //  double r100 = volumeAt100PctRx / targetVolume;
      //  // r50
      //  double r50 = volumeAt50PctRx / targetVolume;


      //  // display info to user
      //  MessageBox.Show(
      //    string.Format("Patient Id:\t\t{0}\n\n" +
      //                  "Plan Id:\t\t{1}\n\n" +
      //                  "Max Dose:\t\t{2}\n\n" +
      //                  "R100 (CI):\t\t{3:F1}\n\n" +
      //                  "R50 (Gradient):\t{4:F2}",
      //                  patientId,
      //                  planId,
      //                  plan.Dose.DoseMax3D.ToString(),
      //                  r100,
      //                  r50));


      //}
      //catch (Exception ex)
      //{
      //  // string interpolation does not work in single file plugins
      //  //MessageBox.Show($"Sorry, something went wrong.\n\n{ex.Message}\n\n{ex.StackTrace}");
      //  MessageBox.Show(string.Format("Sorry, something went wrong.\n\n{0}\n\n{1}", ex.Message, ex.StackTrace));
      //  throw;
      //}


      #endregion original code



    }
    /// <summary>
    /// Returns the plan's structureset
    /// </summary>
    /// <param name="planSetup"></param>
    /// <returns>StructureSet</returns>
    private static StructureSet GetStructureSet(PlanSetup planSetup)
    {
      return planSetup.StructureSet;
    }

    /// <summary>
    /// Displays the calculate conformity indices to the user for plan evaluation
    /// </summary>
    /// <param name="patientId"></param>
    /// <param name="plan"></param>
    /// <param name="planId"></param>
    /// <param name="r100"></param>
    /// <param name="r50"></param>
    private static void DisplayConformityStats(Patient patient, PlanSetup plan, double r100, double r50)
    {
      // display info to user
      MessageBox.Show(
        string.Format("Patient Id:\t\t{0}\n\n" +
                      "Plan Id:\t\t{1}\n\n" +
                      "Max Dose:\t\t{2}\n\n" +
                      "R100 (CI):\t\t{3:F1}\n\n" +
                      "R50 (Gradient):\t{4:F2}",
                      patient.Id,
                      plan.Id,
                      plan.Dose.DoseMax3D.ToString(),
                      r100,
                      r50));
    }

    /// <summary>
    /// Method used to get the Body structure from a structure set
    /// <para>Determined by DicomType</para>
    /// </summary>
    /// <param name="structureSet"></param>
    /// <returns>Returns the structure whose dicom type is EXTERNAL</returns>
    private static Structure GetBodyStructure(StructureSet structureSet)
    {

      // get the body
      return structureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL");
    }

    private static Structure GetPlanTargetVolume(string targetId, StructureSet structureSet)
    {
      return structureSet.Structures.FirstOrDefault(x => x.Id == targetId);
    }


    /// <summary>
    /// Validates that a patient and plan with valid dose are pulled into the current context
    /// </summary>
    /// <param name="context"></param>
    private void ValidatePatientAndPlanDose(ScriptContext context)
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

     
    }

    /// <summary>
    /// Validates whether the target structure is null
    /// </summary>
    /// <param name="planTargetVolume"></param>
    private void ValidateTargetVolume(Structure planTargetVolume)
    {
      if (planTargetVolume == null)
      {
        MessageBox.Show("Please select a target volume");
        return;
      }
    }

    // validates the structure set - unnecessary
    //// check to make sure the plan has a structure set -- since plan has valid dose by this point, there must be a structure set
    //if (structureSet == null)
    //{
    //  MessageBox.Show("Sorry, the selected plan does not have a structure set");
    //  return;
    //}

    /// <summary>
    /// Calculates the Ratio of the 100% isodose line to the ptv volume - aka R100 or Basic CI
    /// </summary>
    /// <param name="plan"></param>
    /// <param name="volumeOfInterestStructure"></param>
    /// <param name="planTargetVolume"></param>
    /// <returns></returns>
    private double CalculateR100(PlanSetup plan, Structure volumeOfInterestStructure, Structure planTargetVolume)
    {
      // volume at 100% IDL (isodose line) - the Rx (prescription) dose
      return plan.GetVolumeAtDose(volumeOfInterestStructure, new DoseValue(100, DoseValue.DoseUnit.Percent), VolumePresentation.AbsoluteCm3) / planTargetVolume.Volume;
    }

    /// <summary>
    /// Calculates the Ratio of the 50% isodose line to the ptv volume - aka R50
    /// </summary>
    /// <param name="plan"></param>
    /// <param name="volumeOfInterestStructure"></param>
    /// <param name="planTargetVolume"></param>
    /// <returns></returns>
    private double CalculateR50(PlanSetup plan, Structure volumeOfInterestStructure, Structure planTargetVolume)
    {
      // volume at 50% IDL (isodose line) - 50% of the Rx (prescription) dose
      return plan.GetVolumeAtDose(volumeOfInterestStructure, new DoseValue(50, DoseValue.DoseUnit.Percent), VolumePresentation.AbsoluteCm3) / planTargetVolume.Volume;
    }




  }
}
