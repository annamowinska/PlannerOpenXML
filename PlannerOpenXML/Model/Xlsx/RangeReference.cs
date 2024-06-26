using CommunityToolkit.Mvvm.ComponentModel;

namespace PlannerOpenXML.Model.Xlsx;

public partial class RangeReference : ObservableObject
{
    #region fields
    private static readonly char[] m_Separator = [':'];
    #endregion fields

    #region properties
    [ObservableProperty]
    private CellReference m_From;

    [ObservableProperty]
    private CellReference m_To;
    #endregion properties

    #region events
    public event EventHandler? RangeReferenceChanged;
    #endregion events

    #region constructors
    public RangeReference(uint columnFrom, uint rowFrom, uint columnTo, uint rowTo)
    {
        From = new CellReference(columnFrom, rowFrom);
        To = new CellReference(columnTo, rowTo);
    }

    public RangeReference(string columnFrom, uint rowFrom, string columnTo, uint rowTo)
    {
        From = new CellReference(columnFrom, rowFrom);
        To = new CellReference(columnTo, rowTo);
    }

    public RangeReference(CellReference from, CellReference to)
    {
        From = from;
        To = to;
    }

    public RangeReference(string addressFrom, string addressTo)
    {
        From = new CellReference(addressFrom);
        To = new CellReference(addressTo);
    }

    public RangeReference(string addressRange)
    {
        var splitted = addressRange.Split(m_Separator, StringSplitOptions.RemoveEmptyEntries);
        if (splitted.Length != 2)
        {
            throw new ArgumentOutOfRangeException(nameof(addressRange), "Not a valid range reference");
        }
        From = new CellReference(splitted[0]);
        To = new CellReference(splitted[1]);
    }

    public RangeReference(RangeReference other)
    {
        From = new CellReference(other.m_From);
        To = new CellReference(other.m_To);
    }
    #endregion constructors

    #region methods
    public override string ToString()
    {
        return From.ToString() + ":" + To.ToString();
    }
    #endregion methods

    #region private methods
    partial void OnFromChanged(CellReference? oldValue, CellReference newValue)
    {
        if (oldValue is not null)
            oldValue.CellReferenceChanged -= CellReferenceChanged;

        newValue.CellReferenceChanged += CellReferenceChanged;
        RangeReferenceChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnToChanged(CellReference? oldValue, CellReference newValue)
    {
        if (oldValue is not null)
            oldValue.CellReferenceChanged -= CellReferenceChanged;

        newValue.CellReferenceChanged += CellReferenceChanged;
        RangeReferenceChanged?.Invoke(this, EventArgs.Empty);
    }

    private void CellReferenceChanged(object? sender, EventArgs e)
    {
        RangeReferenceChanged?.Invoke(this, EventArgs.Empty);
    }
    #endregion private methods
}
