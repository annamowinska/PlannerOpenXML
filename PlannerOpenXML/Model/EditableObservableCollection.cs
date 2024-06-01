using CommunityToolkit.Mvvm.Input;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;

namespace PlannerOpenXML.Model
{
    public class EditableObservableCollection<T> : ObservableCollection<T> where T: INotifyPropertyChanged
    {
        #region properties
        public static Func<T>? OnCreateNew { get; set; }

        public static Func<T, T>? OnCloneItem { get; set; }

        public static Func<List<SortDescription>>? Sorting { get; set; }

        [JsonIgnore]
        public string? LastPath => m_LastPath;
        private string? m_LastPath;

        [JsonIgnore]
        public bool Changed { get { return m_Changed; } set { Set(ref m_Changed, value); } }
        private bool m_Changed;

        [JsonIgnore]
        public bool CanSave { get { return m_CanSave; } set { Set(ref m_CanSave, value); } }
        private bool m_CanSave;

        [JsonIgnore]
        public T? Selected { get { return m_Selected; } set { Set(ref m_Selected, value); } }
        private T? m_Selected;

        [JsonIgnore]
        public ICollectionView View { get; }
        #endregion properties

        #region commands
        [JsonIgnore]
        public IRelayCommand AddCommand => m_AddCommand ??= new RelayCommand(OnAdd);
        private IRelayCommand? m_AddCommand;

        private void OnAdd()
        {
            if (OnCreateNew is null)
                return;

            var result = OnCreateNew();
            Add(result);
            View.Refresh();
            Selected = result;
        }

        [JsonIgnore]
        public IRelayCommand CloneCommand => m_CloneCommand ??= new RelayCommand(OnClone);
        private IRelayCommand? m_CloneCommand;

        private void OnClone()
        {
            if (OnCloneItem is null)
                return;

            if (m_Selected is null)
                return;

            var result = OnCloneItem(m_Selected);
            Add(result);
            Selected = result;
        }

        [JsonIgnore]
        public IRelayCommand RemoveCommand => m_RemoveCommand ??= new RelayCommand(OnRemove);
        private IRelayCommand? m_RemoveCommand;

        private void OnRemove()
        {
            if (m_Selected is null)
                return;

            var position = Items.IndexOf(m_Selected) - 1;
            Remove(m_Selected);

            if (Items.Count == 0)
            {
                Selected = default;
                return;
            }
            Selected = Items[position < 0 ? 0 : position];
        }
        #endregion commands

        #region constructors
        public EditableObservableCollection()
        {
            View = CreateSorted(this, Sorting);
        }

        public EditableObservableCollection(IEnumerable<T> collection) : base(collection)
        {
            foreach (var item in collection)
            {
                item.PropertyChanged += Item_PropertyChanged;
            }
            View = CreateSorted(this, Sorting);
        }

        public EditableObservableCollection(List<T> list) : base(list)
        {
            foreach (var item in list)
            {
                item.PropertyChanged += Item_PropertyChanged;
            }
            View = CreateSorted(this, Sorting);
        }
        #endregion constructors

        #region methods
        public static EditableObservableCollection<T> Load(string path)
        {
            if (!File.Exists(path))
            {
                return [];
            }

            var json = File.ReadAllText(path);
            var item = JsonConvert.DeserializeObject<EditableObservableCollection<T>>(json) ?? [];
            item.Selected = item.FirstOrDefault();
            item.m_LastPath = path;
            return item;
        }

        public void Save()
        {
            if (string.IsNullOrWhiteSpace(m_LastPath))
                return;

            File.WriteAllText(m_LastPath, JsonConvert.SerializeObject(this));
            Changed = false;
            CanSave = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            File.WriteAllText(path, JsonConvert.SerializeObject(this));
            Changed = false;
            CanSave = false;
            m_LastPath = path;
        }
        #endregion methods

        #region private methods
        protected override void MoveItem(int oldIndex, int newIndex)
        {
            base.MoveItem(oldIndex, newIndex);
            OnChanged();
        }

        protected override void ClearItems()
        {
            foreach (var item in Items)
            {
                item.PropertyChanged -= Item_PropertyChanged;
            }

            base.ClearItems();
            OnChanged();
        }

        protected override void InsertItem(int index, T item)
        {
            item.PropertyChanged += Item_PropertyChanged;

            base.InsertItem(index, item);
            OnChanged();
        }

        protected override void RemoveItem(int index)
        {
            var item = Items[index];
            item.PropertyChanged -= Item_PropertyChanged;

            base.RemoveItem(index);
            OnChanged();
        }

        protected override void SetItem(int index, T item)
        {
            var oldItem = Items[index];
            oldItem.PropertyChanged -= Item_PropertyChanged;
            item.PropertyChanged += Item_PropertyChanged;

            base.SetItem(index, item);
            OnChanged();
        }

        protected bool Set<U>(ref U field, U value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<U>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
            return true;
        }

        private void Item_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            OnChanged();
            View.Refresh();
        }

        private static ICollectionView CreateSorted<U>(U source, Func<List<SortDescription>>? sortDescriptions) where U : System.Collections.IEnumerable, INotifyCollectionChanged
        {
            return Application.Current.Dispatcher.Invoke(() =>
            {
                var view = CollectionViewSource.GetDefaultView(source);
                if (sortDescriptions != null)
                {
                    foreach (var item in sortDescriptions())
                    {
                        view.SortDescriptions.Add(item);
                    }
                }
                return view;
            });
        }

        private void OnChanged()
        {
            Changed = true;
            if (!string.IsNullOrWhiteSpace(m_LastPath))
                CanSave = true;
        }
        #endregion private methods
    }
}
