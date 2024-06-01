using System.ComponentModel;
using System.Windows.Input;
using System.Windows;

namespace PlannerOpenXML.UserControls;

/// <summary>
/// Interaction logic for EditableListBox.xaml
/// </summary>
public partial class EditableListBox
{
    #region dependency properties
    public static readonly DependencyProperty ItemsProperty =
        DependencyProperty.Register(nameof(Items), typeof(ICollectionView),
            typeof(EditableListBox), new UIPropertyMetadata(null));

    public ICollectionView Items
    {
        get { return (ICollectionView)GetValue(ItemsProperty); }
        set { SetValue(ItemsProperty, value); }
    }

    public static readonly DependencyProperty ItemsDisplayMemberPathProperty =
        DependencyProperty.Register(nameof(ItemsDisplayMemberPath), typeof(string),
            typeof(EditableListBox), new UIPropertyMetadata(string.Empty));

    public string ItemsDisplayMemberPath
    {
        get { return (string)GetValue(ItemsDisplayMemberPathProperty); }
        set { SetValue(ItemsDisplayMemberPathProperty, value); }
    }

    public static readonly DependencyProperty ItemsSelectedValuePathProperty =
        DependencyProperty.Register(nameof(ItemsSelectedValuePath), typeof(string),
            typeof(EditableListBox), new UIPropertyMetadata(string.Empty));

    public string ItemsSelectedValuePath
    {
        get { return (string)GetValue(ItemsSelectedValuePathProperty); }
        set { SetValue(ItemsSelectedValuePathProperty, value); }
    }

    public static readonly DependencyProperty SelectedProperty =
        DependencyProperty.Register(nameof(Selected), typeof(object),
            typeof(EditableListBox), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, (d, e) => ((EditableListBox)d).OnSelectedChanged(e)));

    public object? Selected
    {
        get { return GetValue(SelectedProperty); }
        set { SetValue(SelectedProperty, value); }
    }

    public static readonly DependencyPropertyKey IsSelectedProperty =
        DependencyProperty.RegisterReadOnly(nameof(IsSelected), typeof(bool),
        typeof(EditableListBox), new UIPropertyMetadata(false));

    public bool IsSelected
    {
        get { return (bool)GetValue(IsSelectedProperty.DependencyProperty); }
        set { SetValue(IsSelectedProperty, value); }
    }

    public static readonly DependencyProperty AddProperty =
        DependencyProperty.Register(nameof(Add), typeof(ICommand),
            typeof(EditableListBox), new UIPropertyMetadata(null));

    public ICommand Add
    {
        get { return (ICommand)GetValue(AddProperty); }
        set { SetValue(AddProperty, value); }
    }

    public static readonly DependencyProperty RemoveProperty =
        DependencyProperty.Register(nameof(Remove), typeof(ICommand),
            typeof(EditableListBox), new UIPropertyMetadata(null));

    public ICommand Remove
    {
        get { return (ICommand)GetValue(RemoveProperty); }
        set { SetValue(RemoveProperty, value); }
    }

    public static readonly DependencyProperty CloneProperty =
        DependencyProperty.Register(nameof(Clone), typeof(ICommand),
            typeof(EditableListBox), new UIPropertyMetadata(null));

    public ICommand Clone
    {
        get { return (ICommand)GetValue(CloneProperty); }
        set { SetValue(CloneProperty, value); }
    }
    #endregion dependency properties

    #region constructors
    public EditableListBox()
    {
        InitializeComponent();
    }
    #endregion constructors

    #region private methods
    private void OnSelectedChanged(DependencyPropertyChangedEventArgs e)
    {
        SetValue(IsSelectedProperty, e.NewValue is not null);
    }
    #endregion private methods
}
