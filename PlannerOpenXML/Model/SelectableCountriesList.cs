using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;

namespace PlannerOpenXML.Model;
public class SelectableCountiesList : ObservableObject
{
    public ObservableCollection<SelectableCountry> Countries { get; } = new ObservableCollection<SelectableCountry>
    {
        new SelectableCountry("Germany", "DE"),
        new SelectableCountry("Hungary", "HU"),
        new SelectableCountry("Poland", "PL"),
        new SelectableCountry("USA", "US")
    };
}