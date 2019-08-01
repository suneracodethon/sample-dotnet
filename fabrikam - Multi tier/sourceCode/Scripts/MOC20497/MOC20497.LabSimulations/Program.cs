using System;
using System.Configuration; // Reference System.Configuration
using MOC20497.VSTS;

namespace Accentient
{
  class Program
  {
    private static VSTSManager vstsManager = new VSTSManager();
    static void SimulateLab2(bool AreasIterationsBacklogOnly)
    {
      Console.WriteLine("");
      Console.Write(" Remove existing iterations and create new ones? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.WriteLine(" Removing iterations");
        vstsManager.RemoveIterations();
        Console.Write(" Creating iterations ");
        vstsManager.CreateIterations();
        Console.WriteLine(" Setting iterations for planning");
        vstsManager.SetPlanningIterations();
      }

      Console.WriteLine("");
      Console.Write(" Remove existing areas and create new ones? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.WriteLine(" Removing areas");
        vstsManager.RemoveAreas();
        Console.Write(" Creating areas ");
        vstsManager.CreateAreas();
        Console.Write(" Setting areas for team ");
        vstsManager.IncludeAllAreas();
      }

      Console.WriteLine("");
      Console.Write(" Destroy all existing work items and create new PBIs? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Destroying work items ");
        vstsManager.DestroyAllWorkItems();
        Console.Write(" Creating PBIs ");
        vstsManager.CreatePBIs();
      }
      if (! AreasIterationsBacklogOnly)
      {
        Console.WriteLine("");
        Console.Write(" Forecast all PBIs into sprint 1? (y/n): ");
        info = Console.ReadKey();
        Console.WriteLine("");
        if (info.KeyChar == 'y' || info.KeyChar == 'Y')
        {
          Console.WriteLine("");
          Console.Write(" Forecasting sprint 1 work ");
          vstsManager.Sprint1Forecast();
        }
        Console.WriteLine("");
        Console.Write(" Destroy existing tasks and create sprint 1 tasks? (y/n): ");
        info = Console.ReadKey();
        Console.WriteLine("");
        if (info.KeyChar == 'y' || info.KeyChar == 'Y')
        {
          Console.WriteLine("");
          Console.Write(" Destroying tasks ");
          vstsManager.DestroyTasks();
          Console.Write(" Creating sprint 1 tasks ");
          vstsManager.Sprint1Plan();
        }
      }
    }
    static void SimulateLab3()
    {
      // SimulateLab2() or equivalent manual steps must be executed before running this

      Console.WriteLine("");
      Console.Write(" Create sprint 1 test plan if necessary? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Destroying all test cases ");
        vstsManager.DestroyTestCases();
        Console.Write(" Creating sprint 1 test cases ");
        vstsManager.CreateSprint1TestCases();
        Console.Write(" Creating sprint 1 test plan ");
        vstsManager.CreateSprint1TestPlan();
      }
      Console.WriteLine("");
      Console.Write(" Mark sprint 1 testing tasks as done? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Marking sprint 1 tasks as done ");
        vstsManager.Mark1SprintTasksDone();
      }
    }
    static void SimulateLab4()
    {
      // SimulateLab3() or equivalent manual steps must be executed before running this

      // Delete and re-create various test results

      Console.WriteLine("");
      Console.Write(" Delete and create test settings? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Deleting test settings ");
        vstsManager.DeleteTestSettings();
        Console.Write(" Creating test settings ");
        vstsManager.CreateTestSettings();
      }
      Console.WriteLine("");
      Console.Write(" Delete existing test results? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Deleting all test results ");
        vstsManager.DeleteSprint1TestResults();
      }
    }
    static void ScriptLab2()
    {
      // Script Lab 2?

      Console.WriteLine("");
      Console.Write("Run Lab 2 scripts? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        SimulateLab2(true);
      }
    }
    static void ScriptAllLabs()
    {
      // Script Lab 2?

      Console.WriteLine("");
      Console.Write("Run scripts to simulate LAB 2? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        SimulateLab2(false);
      }

      // Script Lab 3?

      Console.WriteLine("");
      Console.Write("Run scripts to simulate LAB 3? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        SimulateLab3();
      }

      // Script Lab 4?

      Console.WriteLine("");
      Console.Write("Run scripts to simulate LAB 4? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        SimulateLab4();
      }
    }
    static void Main(string[] args)
    {
      try
      {
        Console.WriteLine("MOC 20497 Setup Utility v2.0");
        Console.WriteLine("");

        // Read config settings

        Console.WriteLine("Configuring environment for MOC20497");
        Console.WriteLine("");
        Console.WriteLine(" Account URL ......... " + vstsManager.AccountURL);
        Console.WriteLine(" Project ............. " + vstsManager.TeamProject);
        Console.WriteLine(" Microsoft Account ... " + vstsManager.MicrosoftAccountAlias);
        Console.WriteLine(" Password ............ " + vstsManager.MicrosoftAccountPassword);
        Console.WriteLine(" XML Input file ...... " + vstsManager.InputFile);
        Console.WriteLine(" Product Owner ....... " + vstsManager.ProductOwner);
        Console.WriteLine(" Script Lab 2 Only? .. " + vstsManager.ScriptLab2Only);
        Console.WriteLine("");

        // Confirm input

        Console.Write("Is the above information correct? (y/n): ");
        ConsoleKeyInfo info = Console.ReadKey();
        Console.WriteLine("");
        if (info.KeyChar != 'y' && info.KeyChar != 'Y')
        {
          Console.WriteLine("");
          Console.WriteLine("Please edit the MOC20497.config file and make any necessary changes.");
          return;
        }

        if (vstsManager.ScriptLab2Only)
        {
          ScriptLab2();
        }
        else
        {
          ScriptAllLabs();
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
      }
      finally
      {
        Console.WriteLine("");
        Console.WriteLine("Press Enter to continue.");
        Console.ReadLine();
      }
    }
  }
}