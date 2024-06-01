using Microsoft.Win32;

namespace PlannerOpenXML.Services;

public class DialogService
{
    #region fields
    public static readonly IReadOnlyDictionary<string, string> XLSX_FILES = new Dictionary<string, string> { { "*.xlsx", "Excel files" } };
    public static readonly IReadOnlyDictionary<string, string> JSON_FILES = new Dictionary<string, string> { { "*.json", "Json  files" } };
    #endregion fields

    #region methods
    public string? OpenSingleFileWithExtensionList(string title, IReadOnlyDictionary<string, string> extensions, bool addAllFilesToo = true, string? defaultPath = null)
        => OpenDialog(title, defaultPath, GetFilter(extensions, addAllFilesToo));

    public string? SaveFileWithExtensionList(string title, IReadOnlyDictionary<string, string> extensions, bool addAllFilesToo = true, string? defaultFileName = null)
    {
        var defaultDir = !string.IsNullOrWhiteSpace(defaultFileName) ? System.IO.Path.GetDirectoryName(defaultFileName) : string.Empty;
        var dialog = new SaveFileDialog
        {
            Title = title,
            FileName = System.IO.Path.GetFileName(defaultFileName),
            InitialDirectory = defaultDir,
            OverwritePrompt = true,
            ValidateNames = true,
            Filter = GetFilter(extensions, addAllFilesToo)
        };
        var result = dialog.ShowDialog();
        if (result.HasValue && result.Value) return dialog.FileName;
        return null;
    }
    #endregion methods

    #region private methods
    private static string GetFilter(IReadOnlyDictionary<string, string> extensions, bool addAllFilesToo)
    {
        var items = new List<string>();

        if (extensions != null)
        {
            foreach (var item in extensions)
            {
                items.Add($"{item.Value} ({item.Key})|{item.Key}");
            }
        }
        if (addAllFilesToo)
            items.Add("All files (*.*)|*.*");

        return string.Join("|", items);
    }

    private static string[] OpenDialogMultiple(string title, string defaultPath, string filter)
    {
        var dialog = new OpenFileDialog
        {
            InitialDirectory = defaultPath,
            CheckPathExists = true,
            Filter = filter,
            Multiselect = true
        };
        var result = dialog.ShowDialog();
        if (result.HasValue && result.Value) return dialog.FileNames;
        return [];
    }

    private static string? OpenDialog(string title, string? defaultPath, string filter)
    {
        var dialog = new OpenFileDialog {
            InitialDirectory = defaultPath,
            CheckPathExists = true,
            Filter = filter, 
            Multiselect = false,
            Title = title
        };
        var result = dialog.ShowDialog();
        return result.HasValue && result.Value ? dialog.FileName : null;
    }
    #endregion private methods
}