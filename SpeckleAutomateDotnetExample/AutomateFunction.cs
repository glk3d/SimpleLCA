using System.Globalization;

using Objects;
using Speckle.Automate.Sdk;
using Speckle.Core.Models;
using Speckle.Core.Models.GraphTraversal;
using SpeckleStructuralModel = Objects.Structural.Analysis.Model;
using Speckle.Core.Api;
using Speckle.Core.Transports;
using Objects.Structural.Geometry;

public static class AutomateFunction
{
    private struct LCAValue
    {
        public string MaterialFamily { get; set; }

        public string MaterialGrade { get; set; }

        public double StageABC { get; set; }

        public double StageD { get; set; }

        public string Unit { get; set; }
    }

    public static async Task Run(AutomationContext automationContext, FunctionInputs functionInputs)
    {
        _ = typeof(ObjectsKit).Assembly; // INFO: Force objects kit to initialize

        // get newly pushed model
        var commitData = await automationContext.ReceiveVersion();
        var currentStream = await automationContext.SpeckleClient
            .ModelGet(
            automationContext.AutomationRunData.ProjectId,
            automationContext.AutomationRunData.Triggers.First().Payload.ModelId);
        var modelName = currentStream.name;

        // receive lca data from another stream
        var lcaData = await GetLCAValues(automationContext, functionInputs);

        // iterate all SpeckleStructuralModels in commit, and attach lca values to elements
        int elemCount = 0;
        int materialCount = 0;
        foreach (var model in FilterStructuralModel(automationContext, commitData))
        {
            if (automationContext.RunStatus == "FAILED") return;
            if (model is null || model.elements is null) continue;
            if (!model.elements.OfType<Element1D>().Any() &&
                !model.elements.OfType<Element2D>().Any())
            {
                automationContext.MarkRunFailed($"No elements of type {typeof(Element1D)} or {typeof(Element2D)} were found.");
                return;
            }

            // get elements grouped by material
            var materialFamily = model.elements
                .GroupBy(e =>
                e is Element1D e1d ? e1d.property.material.materialType.ToString() : // .name :
                e is Element2D e2d ? e2d.property.material.materialType.ToString() : // .name :
                null)
                .Where(group => group.Key != null);

            if (materialFamily is null || materialFamily.Count() == 0)
            {
                automationContext.MarkRunFailed("Could not group elements by material family.");
                return;
            }

            // calculate lca and attach to elements
            foreach (var family in materialFamily)
            {
                var materialGrades = family
                    .GroupBy(e =>
                    e is Element1D e1d ? e1d.property.material.name :
                    e is Element2D e2d ? e2d.property.material.name :
                    null)
                    .Where(g => g.Key != null);

                // select lca based on material family name (Concrete, Timber, Steel...) - if any provided
                var lcaValue = lcaData.FirstOrDefault(l => l.MaterialFamily == family.Key); // default
                foreach (var grade in materialGrades)
                {
                    materialCount++;
                    // select lca based on material grade - if any provided
                    if (lcaData.Any(l => l.MaterialGrade == grade.Key))
                        lcaValue = lcaData.FirstOrDefault(i => i.MaterialGrade == grade.Key);

                    if (string.IsNullOrEmpty(lcaValue.MaterialFamily) ||
                        string.IsNullOrEmpty(lcaValue.MaterialGrade))
                        automationContext.MarkRunException("Some LCA values could not be found. Maybe they are missing in the database.");

                    foreach (var element in family)
                    {
                        // some material get calculated by weight others by volume
                        if (lcaValue.Unit == "kg")
                        {
                            if (element["@Weight"] != null)
                                CreateAndAttachLCA(
                                    element,
                                    Convert.ToDouble(element["@Weight"]) * lcaValue.StageABC,
                                    Convert.ToDouble(element["@Weight"]) * lcaValue.StageD);
                            else if (element["Weight"] != null)
                                CreateAndAttachLCA(
                                    element,
                                    Convert.ToDouble(element["Weight"]) * lcaValue.StageABC,
                                    Convert.ToDouble(element["Weight"]) * lcaValue.StageD);
                            else
                                automationContext.MarkRunException("Some elements did not have their weights attached.");

                        }

                        if (lcaValue.Unit == "m3")
                        {
                            if (element["@Volume"] != null)
                                CreateAndAttachLCA(
                                    element,
                                    Convert.ToDouble(element["@Volume"]) * lcaValue.StageABC,
                                    Convert.ToDouble(element["@Volume"]) * lcaValue.StageD);
                            else if (element["Volume"] != null)
                                CreateAndAttachLCA(
                                    element,
                                    Convert.ToDouble(element["Volume"]) * lcaValue.StageABC,
                                    Convert.ToDouble(element["Volume"]) * lcaValue.StageD);
                            else
                                automationContext.MarkRunException("Some elements did not have their weights attached.");
                        }

                        elemCount++;
                    }
                }
            }
        }

        Console.WriteLine("-------------------------------------------------------------");
        if (automationContext.RunStatus != "FAILED")
        {
            // add another commit
            await automationContext.CreateNewVersionInProject(commitData, $"{modelName} LCA Results");
            // speckle does not allow (yet?) to commit to the same branch
            automationContext.MarkRunSuccess($"Different materials: {materialCount} for {elemCount} Elements.");
        }
    }

    private static void CreateAndAttachLCA(Base element, double abc, double d)
    {
        var LCA = new Base();
        LCA["StageABC"] = abc;
        LCA["StageD"] = d;
        element["LCA"] = LCA;
    }

    private static List<SpeckleStructuralModel> FilterStructuralModel(AutomationContext automationContext, Base @base)
    {
        var traverse = DefaultTraversal.CreateTraversalFunc();
        var structuralModels = traverse.Traverse(@base)
            .Select(o => o.Current)
            .OfType<SpeckleStructuralModel>()
            .ToList();

        if (structuralModels is null || structuralModels.Count() == 0)
        {
            automationContext.MarkRunFailed($"No object of type {typeof(SpeckleStructuralModel)} was found.");
            return null;
        }

        return structuralModels;
    }

    private static async Task<List<LCAValue>> GetLCAValues(AutomationContext automationContext, FunctionInputs functionInputs)
    {
        /*
        if (string.IsNullOrEmpty(functionInputs.LCADataProjectID))
            functionInputs.LCADataProjectID = automationContext.AutomationRunData.ProjectId;
        */

        // find latest excel lca
        var streams = await automationContext.SpeckleClient
            .StreamGetBranches(automationContext.AutomationRunData.ProjectId /*functionInputs.LCADataProjectID */);

        var latestObjectId = streams
            .Where(s => s.id == functionInputs.LCADataModelID)
            .Select(s => s.commits)
            .ToList()[0].items[0].referencedObject;

        // does not work -> commits is null
        // var directExcel = await automationContext.SpeckleClient.ModelGet(functionInputs.LCADataProjectID, functionInputs.LCADataModelID);

        // receive excel lca
        var transport = new ServerTransport(
            automationContext.SpeckleClient.Account,
            automationContext.AutomationRunData.ProjectId/*functionInputs.LCADataProjectID*/);
        var commitData = await Operations.Receive(latestObjectId, transport);

        var data = new List<object>();
        // check if pushed from excel        
        if (commitData["data"] is not null)
        {
            // from excel
            data = commitData["data"] as List<object>;
        }
        else if (commitData.GetMembers(DynamicBaseMemberType.Dynamic).TryGetValue("LCA", out var members))
        {
            // from grasshopper
            foreach (var member in ((Base)members).GetMembers(DynamicBaseMemberType.Dynamic))
                data.Add(member.Value);
        }
        else
        {
            automationContext.MarkRunFailed($"Could not collect LCA data from base.");
        }

        data.RemoveAt(0); // Remove excel header

        var lcaValuesList = new List<LCAValue>();
        foreach (List<object> item in data)
        {
            // otherwise it wont interpret comma correctly
            double.TryParse(
                (item[2] as string),
                System.Globalization.NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double stageABC);

            double.TryParse(
                (item[3] as string),
                System.Globalization.NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double stageD);

            lcaValuesList.Add(
                new LCAValue()
                {
                    MaterialFamily = Convert.ToString(item[0]),
                    MaterialGrade = item[1] as string,
                    StageABC = stageABC,
                    StageD = stageD,
                    Unit = item[4] as string,
                });
        }

        if (lcaValuesList is null || lcaValuesList.Count == 0)
            automationContext.MarkRunFailed($"Could not collect LCA data from base.");

        return lcaValuesList;
    }
}
