using Microsoft.TeamFoundation.Client; // Microsoft.TeamFoundation.Client (Add Reference > Extensions)
using Microsoft.TeamFoundation.ProcessConfiguration.Client; // Microsoft.TeamFoundation.ProjectManagement (Add Reference > Extensions)
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client; // Microsoft.TeamFoundation.TestManagement.Client (Add Reference > Extensions)
using Microsoft.TeamFoundation.WorkItemTracking.Client; // Microsoft.TeamFoundation.WorkItemTracking.Client (Add Reference > Extensions)
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration; // Reference System.Configuration
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace MOC20497.VSTS
{
  public class VSTSManager
  {
    private Hashtable TestCases = new Hashtable();

    #region Properties
    public string AccountURL
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["AccountUrl"].Value;
      }
    }
    public string MicrosoftAccountAlias
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["MicrosoftAccountAlias"].Value;
      }
    }
    public string MicrosoftAccountPassword
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["MicrosoftAccountPassword"].Value;
      }
    }
    public string TeamProject
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["TeamProject"].Value;
      }
    }
    public string InputFile
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["InputFile"].Value;
      }
    }
    public int TotalReleases
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return int.Parse(appSettings.Settings["TotalReleases"].Value);
      }
    }
    public int SprintsPerRelease
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return int.Parse(appSettings.Settings["SprintsPerRelease"].Value);
      }
    }
    public int WeeksPerSprint
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return int.Parse(appSettings.Settings["WeeksPerSprint"].Value);
      }
    }
    public bool ResetSprintsEachRelease
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return bool.Parse(appSettings.Settings["ResetSprintsEachRelease"].Value);
      }
    }
    public List<string> PlanningSprints
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return new List<string>(appSettings.Settings["PlanningIterations"].Value.Split(new char[] { ';' }));
      }
    }
    public string ProductOwner
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return appSettings.Settings["ProductOwner"].Value;
      }
    }
    public bool ScriptLab2Only
    {
      get
      {
        ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        fileMap.ExeConfigFilename = "MOC20497.config";
        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None, false);
        AppSettingsSection appSettings = config.AppSettings;
        return bool.Parse(appSettings.Settings["ScriptLab2Only"].Value);
      }
    }

    #endregion

    private TfsTeamProjectCollection TfsConnect()
    {
      try
      {
        // Determine if local TFS or VSTS

        if ((!AccountURL.ToLower().Contains("https")) || AccountURL.ToLower().Contains("8080"))
        {
          // Authenticate to on-premises TFS

          TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(AccountURL));
          tpc.EnsureAuthenticated();
          return tpc;
        }
        else
        {
          // Authenticate to VSTS

          NetworkCredential netCred = new NetworkCredential(MicrosoftAccountAlias, MicrosoftAccountPassword);
          BasicAuthCredential basicCred = new BasicAuthCredential(netCred);
          TfsClientCredentials tfsCred = new TfsClientCredentials(basicCred);
          tfsCred.AllowInteractive = false;
          TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(AccountURL), tfsCred);
          tpc.Authenticate();
          return tpc;
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
        throw;
      }
    }

    #region Iterations
    public void RemoveIterations()
    {
      try
      {
        TfsTeamProjectCollection tpc = TfsConnect();
        ICommonStructureService css = tpc.GetService<ICommonStructureService>();
        ProjectInfo project = css.GetProjectFromName(TeamProject);
        NodeInfo[] hierarchy = css.ListStructures(project.Uri);
        XmlElement tree;
        if (hierarchy[0].Name.ToLower() == "iteration")
        {
          tree = css.GetNodesXml(new string[] { hierarchy[0].Uri }, true);
        }
        else
        {
          tree = css.GetNodesXml(new string[] { hierarchy[1].Uri }, true);
        }
        string parentUri = tree.FirstChild.Attributes["NodeID"].Value;

        // Enumerate nodes

        if (tree.HasChildNodes)
        {
          XmlNode childrenNode = tree.FirstChild;
          if (childrenNode.HasChildNodes && childrenNode.FirstChild.HasChildNodes)
          {
            string[] nodes = new string[childrenNode.ChildNodes[0].ChildNodes.Count];
            for (int i = 0; i < childrenNode.ChildNodes[0].ChildNodes.Count; i++)
            {
              XmlNode node = childrenNode.ChildNodes[0].ChildNodes[i];
              nodes[i] = node.Attributes["NodeID"].Value;
            }
            css.DeleteBranches(nodes, parentUri);
          }
        }
        Console.WriteLine("");
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
        throw;
      }
    }
    public void CreateIterations()
    {
      try
      {
        TfsTeamProjectCollection tpc = TfsConnect();
        DateTime startDate = DateTime.Today;
        int count = 0;

        // Get service and hierarchy

        ICommonStructureService4 css = tpc.GetService<ICommonStructureService4>();
        ProjectInfo project = css.GetProjectFromName(TeamProject);
        NodeInfo[] hierarchy = css.ListStructures(project.Uri);
        XmlElement tree;
        if (hierarchy[0].Name.ToLower() == "iteration")
        {
          tree = css.GetNodesXml(new string[] { hierarchy[0].Uri }, true);
        }
        else
        {
          tree = css.GetNodesXml(new string[] { hierarchy[1].Uri }, true);
        }
        string parentUri = tree.FirstChild.Attributes["NodeID"].Value;
        int lastSprint = 1;

        for (int release = 1; release <= TotalReleases; release++)
        {
          string releaseUri = css.CreateNode("Release " + release.ToString(), parentUri);
          css.SetIterationDates(releaseUri, startDate, startDate.AddDays(SprintsPerRelease * WeeksPerSprint * 7 - 1));
          for (int sprint = 1; sprint <= SprintsPerRelease; sprint++)
          {
            string sprintUri;
            count++;
            Console.Write(".");
            if (ResetSprintsEachRelease)
              sprintUri = css.CreateNode("Sprint " + sprint.ToString(), releaseUri);
            else
              sprintUri = css.CreateNode("Sprint " + lastSprint, releaseUri);
            css.SetIterationDates(sprintUri, startDate, startDate.AddDays(WeeksPerSprint * 7 - 1));
            startDate = startDate.AddDays(WeeksPerSprint * 7);
            lastSprint++;
          }
        }

        //  RefreshCache and SyncToCache

        WorkItemStore store = new WorkItemStore(tpc);
        store.RefreshCache();
        store.SyncToCache();

        Console.WriteLine(string.Format(" ({0} releases, {1} sprints created)", TotalReleases, count));
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
        throw;
      }
    }
    public void SetPlanningIterations()
    {
      TfsTeamProjectCollection tpc = TfsConnect();
      ICommonStructureService4 css = tpc.GetService<ICommonStructureService4>();
      ProjectInfo project = css.GetProjectFromName(TeamProject);
      TeamSettingsConfigurationService teamConfigService = tpc.GetService<TeamSettingsConfigurationService>();
      var teamConfigs = teamConfigService.GetTeamConfigurationsForUser(new string[] { project.Uri });
      foreach (TeamConfiguration teamConfig in teamConfigs)
      {
        if (teamConfig.IsDefaultTeam)
        {
          TeamSettings settings = teamConfig.TeamSettings;
          settings.IterationPaths = PlanningSprints.ToArray();
          teamConfigService.SetTeamSettings(teamConfig.TeamId, settings);
          Console.Write(new string('.', PlanningSprints.Count));
        }
      }
      Console.WriteLine(string.Format(" ({0} iterations)", PlanningSprints.Count));
    }
    #endregion

    #region Areas
    public void RemoveAreas()
    {
      try
      {
        TfsTeamProjectCollection tpc = TfsConnect();
        ICommonStructureService css = tpc.GetService<ICommonStructureService>();
        ProjectInfo project = css.GetProjectFromName(TeamProject);
        NodeInfo[] hierarchy = css.ListStructures(project.Uri);
        XmlElement tree;
        if (hierarchy[0].Name.ToLower() == "area")
        {
          tree = css.GetNodesXml(new string[] { hierarchy[0].Uri }, true);
        }
        else
        {
          tree = css.GetNodesXml(new string[] { hierarchy[1].Uri }, true);
        }
        string parentUri = tree.FirstChild.Attributes["NodeID"].Value;

        // Enumerate nodes

        if (tree.HasChildNodes)
        {
          XmlNode childrenNode = tree.FirstChild;
          if (childrenNode.HasChildNodes && childrenNode.FirstChild.HasChildNodes)
          {
            string[] nodes = new string[childrenNode.ChildNodes[0].ChildNodes.Count];
            for (int i = 0; i < childrenNode.ChildNodes[0].ChildNodes.Count; i++)
            {
              XmlNode node = childrenNode.ChildNodes[0].ChildNodes[i];
              nodes[i] = node.Attributes["NodeID"].Value;
            }
            css.DeleteBranches(nodes, parentUri);
          }
        }
        Console.WriteLine();
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
        throw;
      }
    }
    private void AddArea(TfsTeamProjectCollection tpc, string teamProject, string nodeValue, string parentValue = "")
    {
      try
      {
        // Get service and hierarchy

        ICommonStructureService css = tpc.GetService<ICommonStructureService>();
        ProjectInfo project = css.GetProjectFromName(teamProject);
        NodeInfo[] hierarchy = css.ListStructures(project.Uri);
        XmlElement tree;
        if (hierarchy[0].Name.ToLower() == "area")
        {
          tree = css.GetNodesXml(new string[] { hierarchy[0].Uri }, true);
        }
        else
        {
          tree = css.GetNodesXml(new string[] { hierarchy[1].Uri }, true);
        }

        string parentUri = "";
        if (parentValue == "")
        {
          parentUri = tree.FirstChild.Attributes["NodeID"].Value;
        }
        else
        {
          parentUri = tree.SelectSingleNode("//Children/Node[@Name='" + parentValue + "']").Attributes["NodeID"].Value;
        }
        css.CreateNode(nodeValue, parentUri);
      }
      catch
      {
      }
    }
    public void CreateAreas()
    {
      TfsTeamProjectCollection tpc = TfsConnect();
      XmlDocument xmlInput = new XmlDocument();
      xmlInput.Load(InputFile);
      int count = 0;
      XmlNodeList areas = xmlInput.SelectNodes("//areas/area");
      foreach (XmlNode area in areas)
      {
        count++;
        Console.Write(".");
        if (area.Attributes["parent"] != null)
          AddArea(tpc, TeamProject, area.InnerText, area.Attributes["parent"].Value);
        else
          AddArea(tpc, TeamProject, area.InnerText);
      }

      //  RefreshCache and SyncToCache

      WorkItemStore store = new WorkItemStore(tpc);
      store.RefreshCache();
      store.SyncToCache();

      Console.WriteLine(string.Format(" ({0} areas created)", count));
    }
    public void IncludeAllAreas()
    {
      try
      {
        TfsTeamProjectCollection tpc = TfsConnect();

        ICommonStructureService4 css = tpc.GetService<ICommonStructureService4>();
        ProjectInfo project = css.GetProjectFromName(TeamProject);
        TeamSettingsConfigurationService teamConfigService = tpc.GetService<TeamSettingsConfigurationService>();
        var teamConfigs = teamConfigService.GetTeamConfigurationsForUser(new string[] { project.Uri });
        foreach (TeamConfiguration teamConfig in teamConfigs)
        {
          if (teamConfig.IsDefaultTeam)
          {
            TeamSettings settings = teamConfig.TeamSettings;
            foreach (TeamFieldValue tfv in settings.TeamFieldValues)
            {
              tfv.IncludeChildren = true;
              teamConfigService.SetTeamSettings(teamConfig.TeamId, settings);
            }
          }
        }
        Console.WriteLine();
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: " + ex.Message);
        throw;
      }
    }

    #endregion

    #region Work Items
    public int DestroyWorkItems(string workItemType)
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Get work items

      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];
      string wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "'";
      if (workItemType != "" && workItemType != "*")
      {
        wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] = '" + workItemType + "'";
      }
      WorkItemCollection collection = store.Query(wiql);

      // Get list of work items to destroy

      ArrayList destroyIds = new ArrayList();
      for (int i = 0; i < collection.Count; i++)
      {
        WorkItem wi = collection[i];
        destroyIds.Add(wi.Id);
        Console.Write(".");
      }

      if (destroyIds.Count < 1)
      {
        return 0;
      }

      // Convert to int array

      int[] ids = new int[destroyIds.Count];
      for (int i = 0; i < destroyIds.Count; i++)
      {
        ids[i] = Convert.ToInt32(destroyIds[i]);
      }

      // Destroy work items

      IEnumerable<WorkItemOperationError> errors = store.DestroyWorkItems(ids);
      List<WorkItemOperationError> errorList = new List<WorkItemOperationError>(errors);
      if (errorList.Count > 0)
      {
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < errorList.Count; i++)
        {
          builder.AppendLine(string.Format("> Error destroying work item {0}, {1}", errorList[i].Id, errorList[i].Exception.Message));
        }
        Console.WriteLine(builder.ToString());
      }
      return destroyIds.Count;
    }
    public int DestroyWorkItems()
    {
      return DestroyWorkItems("*");
    }
    #endregion

    #region PBIs
    public void DestroyAllWorkItems()
    {
      int count = DestroyWorkItems();
      Console.WriteLine(string.Format(" ({0} work items destroyed)", count));
    }

    public void CreatePBIs()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      int count = 0;
      Hashtable workItems = new Hashtable();

      // Validate that it's the right kind of team project

      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];
      if (!project.WorkItemTypes.Contains("Product Backlog Item"))
      {
        Console.WriteLine("This team project was not created using the Visual Studio Scrum process template.");
        return;
      }

      // Load work items

      XmlDocument xmlInput = new XmlDocument();
      xmlInput.Load(InputFile);
      XmlNodeList items = xmlInput.SelectNodes("//backlog/item");
      foreach (XmlNode item in items)
      {
        Console.Write(".");
        string WIT = item["type"].InnerText.Trim().ToLower();
        string title = item["title"].InnerText;
        string state;
        if (item["state"] == null)
        {
          state = "approved";
        }
        else
        {
          state = item["state"].InnerText.ToLower();
        }
        WorkItem wi = new WorkItem(project.WorkItemTypes[WIT]);
        wi.Title = title;
        if (item["description"] != null)
        {
          wi.Description = item["description"].InnerText;
        }
        else
        {
          wi.Description = title;
        }
        if (item["reprosteps"] != null)
        {
          wi.Fields["Microsoft.VSTS.TCM.ReproSteps"].Value = item["reprosteps"].InnerText;
        }
        if (item["acceptancecriteria"] != null)
        {
          wi.Fields["Microsoft.VSTS.Common.AcceptanceCriteria"].Value = item["acceptancecriteria"].InnerText;
        }
        wi.AreaPath = TeamProject + item["area"].InnerText;
        if (item["iteration"] != null)
        {
          wi.IterationPath = TeamProject + item["iteration"].InnerText;
        }
        wi.Fields["System.AssignedTo"].Value = ProductOwner;
        if (state == "approved" || state == "active")
        {
          if (item["businessvalue"] != null && item["businessvalue"].InnerText != "")
          {
            wi.Fields["Microsoft.VSTS.Common.BusinessValue"].Value = item["businessvalue"].InnerText;
          }
          if (item["effort"] != null && item["effort"].InnerText != "")
          {
            wi.Fields["Microsoft.VSTS.Scheduling.Effort"].Value = item["effort"].InnerText;
          }
          if (item["order"] != null && item["order"].InnerText != "")
          {
            wi.Fields["Microsoft.VSTS.Common.BacklogPriority"].Value = item["order"].InnerText;
          }
        }

        ArrayList ValidationResult = wi.Validate();
        if (ValidationResult.Count > 0)
        {
          Microsoft.TeamFoundation.WorkItemTracking.Client.Field badField = (Microsoft.TeamFoundation.WorkItemTracking.Client.Field)ValidationResult[0];
          Console.WriteLine();
          Console.WriteLine("  Invalid \"{0}\" value \"{1}\" ({2})", badField.Name, badField.Value, badField.Status);
          return;
        }
        wi.Save();
        count++;
        workItems.Add(wi.Id, wi.Title.Trim().ToLower());

        // linked items?

        if (item["parent"] != null)
        {
          string parent = item["parent"].InnerText.Trim().ToLower();
          int parentId = 0;
          foreach (DictionaryEntry candidate in workItems)
          {
            if (candidate.Value.ToString() == parent)
            {
              parentId = Convert.ToInt32(candidate.Key);
            }
          }
          if (parentId > 0)
          {
            // lookup child link type

            WorkItemLinkTypeEnd childLink = store.WorkItemLinkTypes.LinkTypeEnds["Parent"];
            wi.WorkItemLinks.Add(new WorkItemLink(childLink, parentId));
            wi.Save();
          }
        }

        if (state == "approved")
        {
          wi.State = "Approved";
          wi.Save();
        }
        else if (state == "active")
        {
          wi.State = "Active";
          wi.Save();
        }
        else if (state == "committed")
        {
          wi.State = "Committed";
          wi.Save();
        }
        else if (state == "done")
        {
          wi.State = "Done";
          wi.Save();
        }
      }

      // Done

      Console.WriteLine(string.Format(" ({0} PBIs created)", count));
    }
    #endregion

    #region Sprint 1 Backlog
    public void DestroyTasks()
    {
      int count = DestroyWorkItems("Task");
      Console.WriteLine(string.Format(" ({0} tasks destroyed)", count));
    }
    public void Sprint1Forecast()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Get work items

      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];
      string wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] ='Product Backlog Item'";
      WorkItemCollection collection = store.Query(wiql);

      // Get list of work items to put into Sprint 1

      int count = 0;
      for (int i = 0; i < collection.Count; i++)
      {
        WorkItem wi = collection[i];

        // Set Iteration Path and State

        wi.Open();
        wi.IterationPath = @"Fabrikam\Release 1\Sprint 1";
        wi.State = "Committed";
        wi.Save();
        Console.Write(".");
        count++;
      }

      // Done

      Console.WriteLine(string.Format(" ({0} PBIs added)", count));
    }
    public void Sprint1Plan()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Load tasks from config file

      XmlDocument xmlInput = new XmlDocument();
      xmlInput.Load(InputFile);
      XmlNodeList tasks = xmlInput.SelectNodes("//plan/task");

      // Get PBI work items

      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];
      string wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] ='Product Backlog Item'";
      WorkItemCollection collection = store.Query(wiql);

      // Loop through each PBI

      int count = 0;
      for (int i = 0; i < collection.Count; i++)
      {
        WorkItem PBI = collection[i];

        // Loop through each Task

        foreach (XmlNode task in tasks)
        {
          WorkItem wi = new WorkItem(project.WorkItemTypes["task"]);
          wi.Title = task["title"].InnerText;
          wi.IterationPath = @"Fabrikam\Release 1\Sprint 1";
          wi.Fields["Microsoft.VSTS.Scheduling.RemainingWork"].Value = Convert.ToInt32(task["remainingwork"].InnerText);

          ArrayList ValidationResult = wi.Validate();
          if (ValidationResult.Count > 0)
          {
            Microsoft.TeamFoundation.WorkItemTracking.Client.Field badField = (Microsoft.TeamFoundation.WorkItemTracking.Client.Field)ValidationResult[0];
            Console.WriteLine();
            Console.WriteLine("  Invalid \"{0}\" value \"{1}\" ({2})", badField.Name, badField.Value, badField.Status);
            return;
          }
          Console.Write(".");
          wi.Save();

          // Save link to parent PBI

          WorkItemLinkTypeEnd childLink = store.WorkItemLinkTypes.LinkTypeEnds["Parent"];
          wi.WorkItemLinks.Add(new WorkItemLink(childLink, PBI.Id));
          wi.Save();
          count++;
        }
      }

      // Done

      Console.WriteLine(string.Format(" ({0} tasks added)", count));
    }
    #endregion

    #region Sprint 1 Test Plan

    public void DestroyTestCases()
    {
      int count = DestroyWorkItems("Test Case");
      Console.WriteLine(string.Format(" ({0} test cases destroyed)", count));
    }

    public void CreateSprint1TestCases()
    {
      TfsTeamProjectCollection tpc = TfsConnect();
      ITestManagementTeamProject testProject = tpc.GetService<ITestManagementService>().GetTeamProject(TeamProject);

      // Get PBI work items

      int count = 0;
      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];

      string wiql = "SELECT [System.Id], [System.Title], [Microsoft.VSTS.Common.AcceptanceCriteria] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] = 'Product Backlog Item'";
      WorkItemCollection Backlog = store.Query(wiql);
      for (int i = 0; i < Backlog.Count; i++)
      {
        WorkItem pbi = Backlog[i];

        // Create the basic test case

        WorkItem wi = new WorkItem(project.WorkItemTypes["test case"]);
        wi.Title = "Verify " + pbi.Title + " behavior";
        wi.IterationPath = @"Fabrikam\Release 1\Sprint 1";
        wi.Save();
        int id = wi.Id;
        TestCases.Add(pbi.Title.ToLower(), id);

        // Access test case through different service

        ITestCase testCase = testProject.TestCases.Find(id);

        // Parse test steps

        string criteria = pbi.Fields["Microsoft.VSTS.Common.AcceptanceCriteria"].Value.ToString();
        criteria = criteria.Remove(0, criteria.IndexOf(@"<ol"));

        Regex regex = new Regex(@"<li>(.+?)<\/li>", RegexOptions.IgnoreCase);
        MatchCollection items = regex.Matches(criteria);
        foreach (Match item in items)
        {
          string title = item.ToString();
          title = title.Replace(@"<li>", "");
          title = title.Replace(@"</li>", "");
          title = title.Trim();
          string expected = "";

          // Click Employees {Employee list displays}

          if (title.IndexOf("{") > 0)
          {
            expected = title.Substring(title.IndexOf("{"));
            expected = expected.Replace(@"{", "");
            expected = expected.Replace(@"}", "");
            expected = expected.Trim();
            title = title.Substring(0, title.IndexOf("{"));
            title = title.Trim();
          }
          ITestStep step = testCase.CreateTestStep();
          step.Title = title;
          step.ExpectedResult = expected;
          testCase.Actions.Add(step);
        }
        testCase.Save();
        Console.Write(".");
        count++;
      }
      Console.WriteLine(string.Format(" ({0} test cases created)", count));
    }

    public void CreateSprint1TestPlan()
    {
      TfsTeamProjectCollection tpc = TfsConnect();
      ITestManagementTeamProject project = tpc.GetService<ITestManagementService>().GetTeamProject(TeamProject);

      // Create test plan if none exist
      //
      // See http://bit.ly/2dup2XY for why we can't delete Test Plans or Suites at this point in time
      //
      // If this routine isn't creating the test plan and/or test suites for you, you'll need to manually
      // delete the Test Plan and Test Suites using witadmin

      WorkItemStore store = new WorkItemStore(tpc);
      string wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] ='Test Plan' AND [System.Title] = 'Sprint 1'";
      WorkItemCollection workItems = store.Query(wiql);
      int testPlanCount = workItems.Count;
      wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] ='Test Suite'";
      int testSuiteCount = store.Query(wiql).Count;
      ITestPlan testPlan;
      if (testPlanCount == 0)
      {
        testPlan = project.TestPlans.Create();
        testPlan.Name = "Sprint 1";
        testPlan.Iteration = @"Fabrikam\Release 1\Sprint 1";
        testPlan.Save();
        Console.WriteLine(" . (1 plan created)");
      }
      else
      {
        testPlan = project.TestPlans.Find(workItems[0].Id);
        Console.WriteLine(" . (plan exists)");
      }

      // Create Test Suites if non exist

      if (testSuiteCount <= 10) // May create duplicate test suites
      { 
        Console.Write(" Creating sprint 1 test suites ");

        // suites

        int count = 0;
        IStaticTestSuite staticSuite = project.TestSuites.CreateStatic();
        staticSuite.Title = "Automated";
        testPlan.RootSuite.Entries.Add(staticSuite);
        testPlan.Save();
        Console.Write(".");
        count++;

        staticSuite = project.TestSuites.CreateStatic();
        staticSuite.Title = "Regression";
        testPlan.RootSuite.Entries.Add(staticSuite);
        testPlan.Save();
        Console.Write(".");
        count++;

        // Requirement-based suites

        // Get PBI work items

        wiql = "SELECT [System.Id] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] ='Product Backlog Item'";
        workItems = store.Query(wiql);
        for (int i = 0; i < workItems.Count; i++)
        {
          WorkItem pbi = workItems[i];

          // Link Test Case to PBI

          int testCaseID = (int)TestCases[pbi.Title.ToLower()];
          WorkItemLinkTypeEnd testedByLink = store.WorkItemLinkTypes.LinkTypeEnds["Tested By"];
          pbi.WorkItemLinks.Add(new WorkItemLink(testedByLink, testCaseID));
          pbi.Save();

          // Create Requirement-based test suite

          IRequirementTestSuite reqSuite = project.TestSuites.CreateRequirement(pbi);
          reqSuite.Title = pbi.Title;
          testPlan.RootSuite.Entries.Add(reqSuite);
          testPlan.Save();
          Console.Write(".");
          count++;
        }

        // Query-based suites

        IDynamicTestSuite querySuite = project.TestSuites.CreateDynamic();
        querySuite.Title = "UI Tests";
        querySuite.Query = project.CreateTestQuery(@"SELECT [System.Id],[System.WorkItemType],[System.Title],[Microsoft.VSTS.Common.Priority],[System.AssignedTo],[System.AreaPath] FROM WorkItems WHERE [System.TeamProject] = @project AND [System.WorkItemType] IN GROUP 'Microsoft.TestCaseCategory' AND [System.AreaPath] UNDER 'Fabrikam' AND [System.IterationPath] UNDER 'Fabrikam\Release 1\Sprint 1' AND [System.Title] CONTAINS 'ui'");
        testPlan.RootSuite.Entries.Add(querySuite);
        testPlan.Save();
        Console.Write(".");
        count++;

        querySuite = project.TestSuites.CreateDynamic();
        querySuite.Title = "Bug Existence Tests";
        querySuite.Query = project.CreateTestQuery(@"SELECT [System.Id],[System.WorkItemType],[System.Title],[Microsoft.VSTS.Common.Priority],[System.AssignedTo],[System.AreaPath] FROM WorkItems WHERE [System.TeamProject] = @project AND [System.WorkItemType] IN GROUP 'Microsoft.TestCaseCategory' AND [System.AreaPath] UNDER 'Fabrikam' AND [System.IterationPath] UNDER 'Fabrikam\Release 1\Sprint 1' AND [System.Title] CONTAINS 'bug'");
        testPlan.RootSuite.Entries.Add(querySuite);
        testPlan.Save();
        Console.Write(".");
        count++;

        Console.WriteLine(string.Format(" ({0} suites created)", count));
      }
    }

    public void Mark1SprintTasksDone()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Get Task work items

      int count = 0;
      WorkItemStore store = new WorkItemStore(tpc);
      Project project = store.Projects[TeamProject];

      string wiql = "SELECT [System.Id], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork] FROM WorkItems " + "WHERE [System.TeamProject] = '" + TeamProject + "' AND [System.WorkItemType] = 'Task'";
      WorkItemCollection tasks = store.Query(wiql);
      for (int i = 0; i < tasks.Count; i++)
      {
        WorkItem task = tasks[i];

        if (task.Title.ToLower().Contains("test plan") ||
            task.Title.ToLower().Contains("test suite") ||
            task.Title.ToLower().Contains("test case"))
        {
          task.Open();
          task.State = "Done";
          task.Save();
          Console.Write(".");
          count++;
        }
      }
      Console.WriteLine(string.Format(" ({0} tasks marked done)", count));
    }
    #endregion

    #region Sprint 1 testing
    public void DeleteTestSettings()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Query and Delete Test Settings

      int count = 0;
      ITestManagementTeamProject project = tpc.GetService<ITestManagementService>().GetTeamProject(TeamProject);
      foreach (ITestSettings testSettings in project.TestSettings.Query("SELECT * FROM TestSettings WHERE Name <> 'Local Test Run'"))
      {
        Console.Write(".");
        testSettings.Delete();
        count++;
      }
      Console.WriteLine(string.Format(" ({0} test settings deleted)", count));
    }
    public void CreateTestSettings()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      int count = 0;
      ITestManagementTeamProject project = tpc.GetService<ITestManagementService>().GetTeamProject(TeamProject);

      count++;
      Console.Write(".");
      var testSettings = project.TestSettings.Create();
      testSettings.Name = "Video only";
      testSettings.Description = "Desktop video Only";
      var testSettingDoc = new XmlDocument();
      testSettingDoc.Load(@"TestSettings\VideoOnly.testsettings.xml");
      testSettings.Settings = testSettingDoc.DocumentElement;
      testSettings.Save();

      count++;
      Console.Write(".");
      testSettings = project.TestSettings.Create();
      testSettings.Name = "System information only";
      testSettings.Description = "System information only";
      testSettingDoc = new XmlDocument();
      testSettingDoc.Load(@"TestSettings\SystemInformationOnly.testsettings.xml");
      testSettings.Settings = testSettingDoc.DocumentElement;
      testSettings.Save();

      count++;
      Console.Write(".");
      testSettings = project.TestSettings.Create();
      testSettings.Name = "Full diagnostic";
      testSettings.Description = "Full diagnostic";
      testSettingDoc = new XmlDocument();
      testSettingDoc.Load(@"TestSettings\FullDiagnostic.testsettings.xml");
      testSettings.Settings = testSettingDoc.DocumentElement;
      testSettings.Save();

      Console.WriteLine(string.Format(" ({0} test settings created)", count));
    }

    public void DeleteSprint1TestResults()
    {
      TfsTeamProjectCollection tpc = TfsConnect();

      // Query and Delete Test Runs

      int count = 0;
      ITestManagementTeamProject project = tpc.GetService<ITestManagementService>().GetTeamProject(TeamProject);
      foreach (ITestRun testRun in project.TestRuns.Query("SELECT * FROM TestRun"))
      {
        Console.Write(".");
        testRun.Delete();
        count++;
      }
      Console.WriteLine(string.Format(" ({0} test runs deleted)", count));
    }

    #endregion

  }
}
