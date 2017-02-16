using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Design;
using System.Windows.Forms;
using Aga.Controls.Tree.NodeControls;

namespace Aga.Controls.Tree
{
    public partial class TreeViewAdv
    {
        private Cursor _innerCursor;

        public override Cursor Cursor
        {
            get
            {
                if (_innerCursor != null)
                    return _innerCursor;
                else
                    return base.Cursor;
            }
            set { base.Cursor = value; }
        }

        #region Internal Properties

        private readonly List<TreeNodeAdv> _rowMap;
        private readonly List<TreeNodeAdv> _selection;
        private int _contentWidth;
        private bool _dragMode;
        private int _firstVisibleRow;
        private InputState _input;
        private int _offsetX;
        private IRowLayout _rowLayout;
        private TreeNodeAdv _selectionStart;
        private bool _suspendSelectionEvent;

        private bool DragMode
        {
            get { return _dragMode; }
            set
            {
                _dragMode = value;
                if (!value)
                {
                    StopDragTimer();
                    if (_dragBitmap != null)
                        _dragBitmap.Dispose();
                    _dragBitmap = null;
                }
                else
                    StartDragTimer();
            }
        }

        internal int ColumnHeaderHeight
        {
            get
            {
                if (UseColumns)
                    return _columnHeaderHeight;
                else
                    return 0;
            }
        }

        /// <summary>
        ///     returns all nodes, which parent is expanded
        /// </summary>
        private IEnumerable<TreeNodeAdv> VisibleNodes
        {
            get
            {
                TreeNodeAdv node = Root;
                while (node != null)
                {
                    node = node.NextVisibleNode;
                    if (node != null)
                        yield return node;
                }
            }
        }

        internal bool SuspendSelectionEvent
        {
            get { return _suspendSelectionEvent; }
            set
            {
                if (value != _suspendSelectionEvent)
                {
                    _suspendSelectionEvent = value;
                    if (!_suspendSelectionEvent && _fireSelectionEvent)
                        OnSelectionChanged();
                }
            }
        }

        internal List<TreeNodeAdv> RowMap
        {
            get { return _rowMap; }
        }

        internal TreeNodeAdv SelectionStart
        {
            get { return _selectionStart; }
            set { _selectionStart = value; }
        }

        internal InputState Input
        {
            get { return _input; }
            set { _input = value; }
        }

        internal bool ItemDragMode { get; set; }

        internal Point ItemDragStart { get; set; }


        /// <summary>
        ///     Number of rows fits to the current page
        /// </summary>
        internal int CurrentPageSize
        {
            get { return _rowLayout.CurrentPageSize; }
        }

        /// <summary>
        ///     Number of all visible nodes (which parent is expanded)
        /// </summary>
        internal int RowCount
        {
            get { return RowMap.Count; }
        }

        private int ContentWidth
        {
            get { return _contentWidth; }
        }

        internal int FirstVisibleRow
        {
            get { return _firstVisibleRow; }
            set
            {
                HideEditor();
                _firstVisibleRow = value;
                UpdateView();
            }
        }

        public int OffsetX
        {
            get { return _offsetX; }
            private set
            {
                HideEditor();
                _offsetX = value;
                UpdateView();
            }
        }

        public override Rectangle DisplayRectangle
        {
            get
            {
                Rectangle r = ClientRectangle;
                //r.Y += ColumnHeaderHeight;
                //r.Height -= ColumnHeaderHeight;
                int w = _vScrollBar.Visible ? _vScrollBar.Width : 0;
                int h = _hScrollBar.Visible ? _hScrollBar.Height : 0;
                return new Rectangle(r.X, r.Y, r.Width - w, r.Height - h);
            }
        }

        internal List<TreeNodeAdv> Selection
        {
            get { return _selection; }
        }

        #endregion

        #region Public Properties

        #region DesignTime

        private static readonly Font _font = new Font("Tahoma", 8.25F, FontStyle.Regular, GraphicsUnit.Point, ((0)), false);
        private readonly TreeColumnCollection _columns;
        private readonly NodeControlsCollection _controls;
        private bool _autoRowHeight;
        private BorderStyle _borderStyle = BorderStyle.Fixed3D;
        private float _bottomEdgeSensivity = 0.3f;
        private Color _dragDropMarkColor = Color.Black;
        private float _dragDropMarkWidth = 3.0f;
        private bool _fullRowSelect;
        private GridLineStyle _gridLineStyle = GridLineStyle.None;
        private bool _hideSelection;
        private bool _highlightDropPosition = true;
        private int _indent = 19;
        private Color _lineColor = SystemColors.ControlDark;
        private ITreeModel _model;
        private int _rowHeight = 16;
        private TreeSelectionMode _selectionMode = TreeSelectionMode.Single;
        private bool _showLines = true;
        private bool _showPlusMinus = true;
        private float _topEdgeSensivity = 0.3f;
        private bool _useColumns;

        [DefaultValue(false), Category("Behavior")]
        public bool ShiftFirstNode { get; set; }

        [DefaultValue(false), Category("Behavior")]
        public bool DisplayDraggingNodes { get; set; }

        [DefaultValue(false), Category("Behavior")]
        public bool FullRowSelect
        {
            get { return _fullRowSelect; }
            set
            {
                _fullRowSelect = value;
                UpdateView();
            }
        }

        [DefaultValue(false), Category("Behavior")]
        public bool UseColumns
        {
            get { return _useColumns; }
            set
            {
                _useColumns = value;
                FullUpdate();
            }
        }

        [DefaultValue(false), Category("Behavior")]
        public bool AllowColumnReorder { get; set; }

        [DefaultValue(true), Category("Behavior")]
        public bool ShowLines
        {
            get { return _showLines; }
            set
            {
                _showLines = value;
                UpdateView();
            }
        }

        [DefaultValue(true), Category("Behavior")]
        public bool ShowPlusMinus
        {
            get { return _showPlusMinus; }
            set
            {
                _showPlusMinus = value;
                FullUpdate();
            }
        }

        [DefaultValue(false), Category("Behavior")]
        public bool ShowNodeToolTips { get; set; }

        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "value"), SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic"), DefaultValue(true), Category("Behavior"), Obsolete("No longer used")]
        public bool KeepNodesExpanded
        {
            get { return true; }
            set { }
        }

        /// <Summary>
        ///     The model associated with this <see cref="TreeViewAdv" />.
        /// </Summary>
        /// <seealso cref="ITreeModel" />
        /// <seealso cref="TreeModel" />
        [Browsable(false)]
        public ITreeModel Model
        {
            get { return _model; }
            set
            {
                if (_model != value)
                {
                    AbortBackgroundExpandingThreads();
                    if (_model != null)
                        UnbindModelEvents();
                    _model = value;
                    CreateNodes();
                    FullUpdate();
                    if (_model != null)
                        BindModelEvents();
                }
            }
        }

        // Tahoma is the default font

        /// <summary>
        ///     The font to render <see cref="TreeViewAdv" /> content in.
        /// </summary>
        [Category("Appearance"), Description("The font to render TreeViewAdv content in.")]
        public override Font Font
        {
            get { return (base.Font); }
            set
            {
                if (value == null)
                    base.Font = _font;
                else
                {
                    if (value == DefaultFont)
                        base.Font = _font;
                    else
                        base.Font = value;
                }
            }
        }

        [DefaultValue(BorderStyle.Fixed3D), Category("Appearance")]
        public BorderStyle BorderStyle
        {
            get { return _borderStyle; }
            set
            {
                if (_borderStyle != value)
                {
                    _borderStyle = value;
                    base.UpdateStyles();
                }
            }
        }

        /// <summary>
        ///     Set to true to expand each row's height to fit the text of it's largest column.
        /// </summary>
        [DefaultValue(false), Category("Appearance"), Description("Expand each row's height to fit the text of it's largest column.")]
        public bool AutoRowHeight
        {
            get { return _autoRowHeight; }
            set
            {
                _autoRowHeight = value;
                if (value)
                    _rowLayout = new AutoRowHeightLayout(this, RowHeight);
                else
                    _rowLayout = new FixedRowHeightLayout(this, RowHeight);
                FullUpdate();
            }
        }

        [DefaultValue(GridLineStyle.None), Category("Appearance")]
        public GridLineStyle GridLineStyle
        {
            get { return _gridLineStyle; }
            set
            {
                if (value != _gridLineStyle)
                {
                    _gridLineStyle = value;
                    UpdateView();
                    OnGridLineStyleChanged();
                }
            }
        }

        [DefaultValue(16), Category("Appearance")]
        public int RowHeight
        {
            get { return _rowHeight; }
            set
            {
                if (value <= 0)
                    throw new ArgumentOutOfRangeException("value");

                _rowHeight = value;
                _rowLayout.PreferredRowHeight = value;
                FullUpdate();
            }
        }

        [DefaultValue(TreeSelectionMode.Single), Category("Behavior")]
        public TreeSelectionMode SelectionMode
        {
            get { return _selectionMode; }
            set { _selectionMode = value; }
        }

        [DefaultValue(false), Category("Behavior")]
        public bool HideSelection
        {
            get { return _hideSelection; }
            set
            {
                _hideSelection = value;
                UpdateView();
            }
        }

        [DefaultValue(0.3f), Category("Behavior")]
        public float TopEdgeSensivity
        {
            get { return _topEdgeSensivity; }
            set
            {
                if (value < 0 || value > 1)
                    throw new ArgumentOutOfRangeException();
                _topEdgeSensivity = value;
            }
        }

        [DefaultValue(0.3f), Category("Behavior")]
        public float BottomEdgeSensivity
        {
            get { return _bottomEdgeSensivity; }
            set
            {
                if (value < 0 || value > 1)
                    throw new ArgumentOutOfRangeException("value should be from 0 to 1");
                _bottomEdgeSensivity = value;
            }
        }

        [DefaultValue(false), Category("Behavior")]
        public bool LoadOnDemand { get; set; }

        [DefaultValue(false), Category("Behavior")]
        public bool UnloadCollapsedOnReload { get; set; }

        [DefaultValue(19), Category("Behavior")]
        public int Indent
        {
            get { return _indent; }
            set
            {
                _indent = value;
                UpdateView();
            }
        }

        [Category("Behavior")]
        public Color LineColor
        {
            get { return _lineColor; }
            set
            {
                _lineColor = value;
                CreateLinePen();
                UpdateView();
            }
        }

        [Category("Behavior")]
        public Color DragDropMarkColor
        {
            get { return _dragDropMarkColor; }
            set
            {
                _dragDropMarkColor = value;
                CreateMarkPen();
            }
        }

        [DefaultValue(3.0f), Category("Behavior")]
        public float DragDropMarkWidth
        {
            get { return _dragDropMarkWidth; }
            set
            {
                _dragDropMarkWidth = value;
                CreateMarkPen();
            }
        }

        [DefaultValue(true), Category("Behavior")]
        public bool HighlightDropPosition
        {
            get { return _highlightDropPosition; }
            set { _highlightDropPosition = value; }
        }

        [Category("Behavior"), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public Collection<TreeColumn> Columns
        {
            get { return _columns; }
        }

        [Category("Behavior"), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Editor(typeof (NodeControlCollectionEditor), typeof (UITypeEditor))]
        public Collection<NodeControl> NodeControls
        {
            get { return _controls; }
        }

        /// <summary>
        ///     When set to true, node contents will be read in background thread.
        /// </summary>
        [Category("Behavior"), DefaultValue(false), Description("Read children in a background thread when expanding.")]
        public bool AsyncExpanding { get; set; }

        public override void ResetFont()
        {
            Font = null;
        }

        private bool ShouldSerializeFont()
        {
            return (!Font.Equals(_font));
        }

        #endregion

        #region RunTime

        private readonly ReadOnlyCollection<TreeNodeAdv> _readonlySelection;
        private DropPosition _dropPosition;
        private TreeNodeAdv _root;

        [Browsable(false)]
        public IToolTipProvider DefaultToolTipProvider { get; set; }

        [Browsable(false)]
        public IEnumerable<TreeNodeAdv> AllNodes
        {
            get
            {
                if (_root.Nodes.Count > 0)
                {
                    TreeNodeAdv node = _root.Nodes[0];
                    while (node != null)
                    {
                        yield return node;
                        if (node.Nodes.Count > 0)
                            node = node.Nodes[0];
                        else if (node.NextNode != null)
                            node = node.NextNode;
                        else
                            node = node.BottomNode;
                    }
                }
            }
        }

        [Browsable(false)]
        public DropPosition DropPosition
        {
            get { return _dropPosition; }
            set { _dropPosition = value; }
        }

        [Browsable(false)]
        public TreeNodeAdv Root
        {
            get { return _root; }
        }

        [Browsable(false)]
        public ReadOnlyCollection<TreeNodeAdv> SelectedNodes
        {
            get { return _readonlySelection; }
        }

        [Browsable(false)]
        public TreeNodeAdv SelectedNode
        {
            get
            {
                if (Selection.Count > 0)
                {
                    if (CurrentNode != null && CurrentNode.IsSelected)
                        return CurrentNode;
                    else
                        return Selection[0];
                }
                else
                    return null;
            }
            set
            {
                if (SelectedNode == value)
                    return;

                BeginUpdate();
                try
                {
                    if (value == null)
                    {
                        ClearSelectionInternal();
                    }
                    else
                    {
                        if (!IsMyNode(value))
                            throw new ArgumentException();

                        ClearSelectionInternal();
                        value.IsSelected = true;
                        CurrentNode = value;
                        EnsureVisible(value);
                    }
                }
                finally
                {
                    EndUpdate();
                }
            }
        }

        [Browsable(false)]
        public TreeNodeAdv CurrentNode { get; internal set; }

        [Browsable(false)]
        public int ItemCount
        {
            get { return RowMap.Count; }
        }

        /// <summary>
        ///     Indicates the distance the content is scrolled to the left
        /// </summary>
        [Browsable(false)]
        public int HorizontalScrollPosition
        {
            get
            {
                if (_hScrollBar.Visible)
                    return _hScrollBar.Value;
                else
                    return 0;
            }
        }

        #endregion

        #endregion
    }
}