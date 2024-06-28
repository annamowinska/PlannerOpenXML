using System.Diagnostics.CodeAnalysis;

namespace PlannerOpenXML.Model.Xlsx;

public class Sheets : List<Sheet>
{
    #region methods
    public new void Add(Sheet sheet)
    {
        base.Add(sheet);
    }

    public bool TryGetByName(string name, [NotNullWhen(true)] out Sheet? sheet)
    {
        sheet = this.FirstOrDefault(x => x.Name == name);
        return sheet is not null;
    }
    #endregion methods
}
