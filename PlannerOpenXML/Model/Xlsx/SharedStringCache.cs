using DocumentFormat.OpenXml.Spreadsheet;
namespace PlannerOpenXML.Model.Xlsx;

internal class SharedStringCache
{
    #region fields
    private readonly Dictionary<string, int> m_Cache = [];
    private readonly SharedStringTable m_StringTable;
    #endregion fields

    #region constructors
    public SharedStringCache(SharedStringTable stringTable)
    {
        m_StringTable = stringTable;
        int i = 0;
        m_Cache.Clear();
        foreach (SharedStringItem item in stringTable.Elements<SharedStringItem>())
        {
            m_Cache.Add(item.InnerText, i);
            i++;
        }
    }

    internal SharedStringCache()
    {
        m_StringTable = new SharedStringTable();
    }
    #endregion constructors

    #region methods
    public int GetIndex(string text)
    {
        if (m_Cache.TryGetValue(text, out int index))
            return index;
        SharedStringItem sharedStringItem = new SharedStringItem(new Text(text));
        m_StringTable.AppendChild(sharedStringItem);
        int result = m_Cache.Count;
        m_Cache.Add(text, result);
        return result;
    }
    #endregion methods
}
