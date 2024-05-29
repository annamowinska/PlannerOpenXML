using Microsoft.Win32;

namespace PlannerOpenXML.Services;

public class DialogService
{
    #region methods
    /// <summary>
    /// Opens a file save dialog for the user.
    /// </summary>
    /// <param name="title">Gets or sets the text shown in the title bar of the file dialog.</param>
    /// <param name="filter">
    ///   Gets or sets the filter string that determines what types of files are displayed
    ///   from either the <seealso cref="OpenFileDialog"/> or <seealso cref="Microsoft.Win32.SaveFileDialog"/>.
    /// </param>
    /// <param name="defaultFileName">
    ///   Gets or sets a string containing the full path of the file selected in a file
    ///   dialog.
    /// </param>
    /// <returns>
    ///   A string containing the full path of the file entered in the dialog if user acceped the dialog.
    ///   Otherwise null.
    /// </returns>
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
    #endregion methods
}