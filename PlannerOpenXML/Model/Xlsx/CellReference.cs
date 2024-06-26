using System.Text.RegularExpressions;

namespace PlannerOpenXML.Model.Xlsx;

public partial class CellReference
{
    #region properties
    private static readonly Regex m_ColumnsRows = ColumnsAndRows();

    public uint Column
    {
        get
        {
            return m_Column;
        }
        set
        {
            if (m_Column != value)
            {
                m_Column = value;
                CellReferenceChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }
    private uint m_Column;

    public uint Row
    {
        get
        {
            return m_Row;
        }
        set
        {
            if (m_Row != value)
            {
                m_Row = value;
                CellReferenceChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }
    private uint m_Row;
    #endregion properties

    #region events
    public event EventHandler? CellReferenceChanged;
    #endregion events

    #region constructors
    public CellReference(uint column, uint row)
    {
        m_Column = column;
        m_Row = row;
    }

    public CellReference(string column, uint row)
    {
        m_Column = ConvertColumnNameToInt(column);
        m_Row = row;
    }

    public CellReference(string address)
    {
        var result = ConvertAddressToColumnRow(address);
        m_Column = result.Item1;
        m_Row = result.Item2;
    }

    public CellReference(CellReference other)
    {
        m_Column = other.m_Column;
        m_Row = other.m_Row;
    }
    #endregion constructors

    #region methods
    public override string ToString() =>
        ConvertColumnNumberToName(m_Column) + m_Row;

    public static uint ConvertColumnNameToInt(string column)
    {
        if (string.IsNullOrEmpty(column)) throw new ArgumentNullException(nameof(column));
        column = column.ToUpperInvariant();
        uint sum = 0;

        for (int i = 0; i < column.Length; i++)
        {
            sum *= 26;
            sum += (uint)(column[i] - 'A' + 1);
        }

        return sum;
    }

    public static string ConvertColumnNumberToName(uint column)
    {
        if (column < 1) return "A";
        string result = string.Empty;

        while (column > 0)
        {
            column -= 1;
            uint digit = column % 26;
            result = (char)('A' + digit) + result;

            column /= 26;
        }

        return result;
    }

    public static Tuple<uint, uint> ConvertAddressToColumnRow(string address)
    {
        Match match = m_ColumnsRows.Match(address);
        if (!match.Success) throw new ArgumentOutOfRangeException(nameof(address), "Not a valid cell reference");
        string column = match.Groups["column"].Value;
        string row = match.Groups["row"].Value;
        return new Tuple<uint, uint>(ConvertColumnNameToInt(column), uint.Parse(row));
    }
    #endregion methods

    #region private methods
    [GeneratedRegex("(?<column>[A-Z]+)(?<row>[0-9]+)")]
    private static partial Regex ColumnsAndRows();
    #endregion private methods
}
