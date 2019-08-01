using System;
using MOC20497.VSTS;

namespace MOC20497.Lab2Setup
{
  class Program
  {
    private static VSTSManager vstsManager = new VSTSManager();
    static void Lab2Setup()
    {
      Console.WriteLine("");
      Console.Write(" Remove existing iterations and create new ones? (y/n): ");
      ConsoleKeyInfo info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Removing iterations ");
        vstsManager.RemoveIterations();
        Console.Write(" Creating iterations ");
        vstsManager.CreateIterations();
        Console.Write(" Setting iterations for planning ");
        vstsManager.SetPlanningIterations();
      }

      Console.WriteLine("");
      Console.Write(" Remove existing areas and create new ones? (y/n): ");
      info = Console.ReadKey();
      Console.WriteLine("");
      if (info.KeyChar == 'y' || info.KeyChar == 'Y')
      {
        Console.WriteLine("");
        Console.Write(" Removing areas");
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
    }
    static void Main(string[] args)
    {
      try
      {
        Console.WriteLine("MOC 20497 Setup Utility v2.0");
        Console.WriteLine("");
        Console.WriteLine(" Account URL ........ " + vstsManager.AccountURL);
        Console.WriteLine(" Project ............ " + vstsManager.TeamProject);
        Console.WriteLine(" Microsoft Account .. " + vstsManager.MicrosoftAccountAlias);
        Console.WriteLine(" Password ........... " + vstsManager.MicrosoftAccountPassword);
        Console.WriteLine(" XML Input file ..... " + vstsManager.InputFile);
        Console.WriteLine(" Product Owner ...... " + vstsManager.ProductOwner);
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
        Console.WriteLine("");
        Console.Write("Run Lab 2 setup? (y/n): ");
        info = Console.ReadKey();
        Console.WriteLine("");
        if (info.KeyChar == 'y' || info.KeyChar == 'Y')
        {
          Lab2Setup();
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