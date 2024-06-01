using PlannerOpenXML.Model;
using System.Windows;

namespace Planner;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    #region methods
    static App()
    {
        EditableObservableCollection<Milestone>.OnCreateNew = () => new Milestone { Description = "New milestone", Date = DateOnly.FromDateTime(DateTime.Now) };
        EditableObservableCollection<Milestone>.OnCloneItem = (s) => new Milestone { Description = $"Clone of {s.Description}", Date = s.Date };
    }
    #endregion methods
}
