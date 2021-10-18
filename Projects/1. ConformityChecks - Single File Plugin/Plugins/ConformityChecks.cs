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
      // validate patient in current context
      ValidatePatient(context);

      // validate plan and dose in current context
      ValidatePlanAndPlanDose(context);

      try
      {

        // do stuff here...

        #region Longest Example using Method Extraction - possible example of Abstraction

        //// patient in the current context
        //Patient patient = GetPatient(context);

        //// current plan setup
        //PlanSetup plan = GetPlanSetup(context);

        //// current plan's structure set
        //StructureSet structureSet = GetStructureSet(plan);

        //// the plan target volume
        //Structure target = GetPlanTargetVolume(plan.TargetVolumeID, structureSet);

        //// verify the target is not null
        //ValidateStructureExists(target, "Please select a target volume");

        //// get the body structure
        //Structure body = GetBodyStructure(structureSet);

        //// calculate the R50 - 50% IDL / target volume
        //double r50 = CalculateR50(plan, body, target);

        //// calculate the R100 - 100% IDL / target volume
        //double r100 = CalculateR100(plan, body, target);

        //// display stats to user
        //DisplayConformityStatistics(patient, plan, r50, r100);

        #endregion



        #region Shorter Example using method extraction -- less broken down -- slightly more redundant / slower

        //// if the plan has a target volume selected...
        //if (string.IsNullOrEmpty(context.PlanSetup.TargetVolumeID) == false)
        //{
        //  // calculate and display stats to user
        //  DisplayConformityStatistics(
        //    context.Patient,
        //    context.PlanSetup,
        //    CalculateR50(context.PlanSetup,
        //                  context.StructureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL"),
        //                  context.StructureSet.Structures.FirstOrDefault(x => x.Id == context.PlanSetup.TargetVolumeID)),
        //    CalculateR100(context.PlanSetup,
        //                  context.StructureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL"),
        //                  context.StructureSet.Structures.FirstOrDefault(x => x.Id == context.PlanSetup.TargetVolumeID))
        //    );
        //}
        //else
        //{
        //  MessageBox.Show("Please select a plan target volume");
        //}

        #endregion



        #region Shortest example - abstracts the longer example into a single method

        CalculateAndDisplayConformityStatistics(context);

        #endregion


      }
      catch (Exception ex)
      {
        // string interpolation does not work in single file plugins
        //MessageBox.Show($"Sorry, something went wrong.\n\n{ex.Message}\n\n{ex.StackTrace}");
        MessageBox.Show(string.Format("Sorry, something went wrong.\n\n{0}\n\n{1}", ex.Message, ex.StackTrace));
        throw;
      }


    }

    /// <summary>
    /// Calculates and displays the R100 and R50 to the user
    /// </summary>
    /// <param name="context"></param>
    private void CalculateAndDisplayConformityStatistics(ScriptContext context)
    {
      // patient in the current context
      Patient patient = GetPatient(context);

      // current plan setup
      PlanSetup plan = GetPlanSetup(context);

      // current plan's structure set
      StructureSet structureSet = GetStructureSet(plan);

      // the plan target volume
      Structure target = GetPlanTargetVolume(plan.TargetVolumeID, structureSet);

      // verify the target is not null
      ValidateStructureExists(target, "Please select a target volume");

      // get the body structure
      Structure body = GetBodyStructure(structureSet);

      // calculate the R50 - 50% IDL / target volume
      double r50 = CalculateR50(plan, body, target);

      // calculate the R100 - 100% IDL / target volume
      double r100 = CalculateR100(plan, body, target);

      // display stats to user
      DisplayConformityStatistics(patient, plan, r50, r100);
    }

    /// <summary>
    /// Displays the conformity statistics from the plan to the user in a message box
    /// </summary>
    /// <param name="patient"></param>
    /// <param name="plan"></param>
    /// <param name="r50"></param>
    /// <param name="r100"></param>
    private static void DisplayConformityStatistics(Patient patient, PlanSetup plan, double r50, double r100)
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
    /// Gets the patient from the current script context
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    private static Patient GetPatient(ScriptContext context)
    {
      return context.Patient;
    }

    /// <summary>
    /// Calculates the ratio of the 50% Isodose Line to the Target Volume
    /// </summary>
    /// <param name="plan"></param>
    /// <param name="structureOfInterest"></param>
    /// <param name="targetStructure"></param>
    /// <returns></returns>
    private static double CalculateR50(PlanSetup plan, Structure structureOfInterest, Structure targetStructure)
    {
      // volume of the 50% IDL in the structure of interest / target volume
      return plan.GetVolumeAtDose(structureOfInterest,
              new DoseValue(50, DoseValue.DoseUnit.Percent),
              VolumePresentation.AbsoluteCm3) / targetStructure.Volume;
    }

    /// <summary>
    /// Calculates the ratio of the 100% Isodose Line to the Target Volume
    /// </summary>
    /// <param name="plan"></param>
    /// <param name="structureOfInterest"></param>
    /// <param name="targetStructure"></param>
    /// <returns></returns>
    private static double CalculateR100(PlanSetup plan, Structure structureOfInterest, Structure targetStructure)
    {
      // volume of the 100% IDL in the structure of interest / target volume
      return plan.GetVolumeAtDose(structureOfInterest,
              new DoseValue(100, DoseValue.DoseUnit.Percent),
              VolumePresentation.AbsoluteCm3) / targetStructure.Volume;
    }

    private static Structure GetBodyStructure(StructureSet structureSet)
    {
      // get the body
      return structureSet.Structures.FirstOrDefault(x => x.DicomType == "EXTERNAL");
    }

    /// <summary>
    /// Gets the plan's target volume
    /// </summary>
    /// <param name="targetId"></param>
    /// <param name="structureSet"></param>
    /// <returns></returns>
    private static Structure GetPlanTargetVolume(string targetId, StructureSet structureSet)
    {
      return structureSet.Structures.FirstOrDefault(x => x.Id == targetId);
    }

    /// <summary>
    /// Gets the plan's structure set
    /// </summary>
    /// <param name="planSetup"></param>
    /// <returns></returns>
    private static StructureSet GetStructureSet(PlanSetup planSetup)
    {
      // current plan's structure set
      return planSetup.StructureSet;
    }

    /// <summary>
    /// Gets the plansetup in the current context
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    private static PlanSetup GetPlanSetup(ScriptContext context)
    {
      // current selected plan
      return context.PlanSetup;
    }

    /// <summary>
    /// Validates that the current context has a Patient
    /// <para>Will alert the user and end the script</para>
    /// </summary>
    /// <param name="context"></param>
    private void ValidatePatient(ScriptContext context)
    {
      // check to make sure a patient is open in eclipse
      if (context.Patient == null)
      {
        MessageBox.Show("Please open a patient");
        return;
      }
    }
    /// <summary>
    /// Validates that the current context has a plansetup and that plan has valid dose
    /// <para>Will alert the user and end the script</para>
    /// </summary>
    /// <param name="context"></param>
    private void ValidatePlanAndPlanDose(ScriptContext context)
    {
      // make sure a plan is pulled in and that its dose is calculated
      if (context.PlanSetup == null || context.PlanSetup.IsDoseValid == false)
      {
        MessageBox.Show("Please open a plan and ensure it has calculated dose");
        return;
      }
    }
    /// <summary>
    /// Validates that the current context has a structure set
    /// <para>Will alert the user and end the script</para>
    /// </summary>
    /// <param name="structureSet"></param>
    private void ValidateStructureSet(ScriptContext context)
    {
      // check to make sure the plan has a structure set -- since plan has valid dose by this point, there must be a structure set
      if (context.StructureSet == null)
      {
        MessageBox.Show("Sorry, the selected plan does not have a structure set");
        return;
      }
    }

    /// <summary>
    /// Validates the provided structure is not null
    /// <para>Will alert user with the provided message and end script</para>
    /// </summary>
    /// <param name="structure"></param>
    /// <param name="alertMessage"></param>
    private void ValidateStructureExists(Structure structure, string alertMessage)
    {
      if (structure == null)
      {
        MessageBox.Show(alertMessage);
        return;
      }
    }





    #region Original Code

    //    // check to make sure a patient is open in eclipse
    //      if (context.Patient == null)
    //      {
    //        MessageBox.Show("Please open a patient");
    //        return;
    //      }
    //      // make sure a plan is pulled in and that its dose is calculated
    //      if (context.PlanSetup == null || context.PlanSetup.IsDoseValid == false)
    //      {
    //        MessageBox.Show("Please open a plan and ensure it has calculated dose");
    //        return;
    //      }

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




    #endregion


  }
}