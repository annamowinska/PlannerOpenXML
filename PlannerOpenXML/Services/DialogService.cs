using Microsoft.Win32;

namespace PlannerOpenXML.Services;

public class DialogService
{
    public string? SaveFileDialog(string title, string filter, string defaultFileName)
    {
        var dialog = new SaveFileDialog
        {
            Title = title,
            Filter = filter,
            FileName = defaultFileName,
            OverwritePrompt = true,
        };

        var result = dialog.ShowDialog();

        if (result.HasValue && result.Value)
            return dialog.FileName;

        return null;
    }
}
