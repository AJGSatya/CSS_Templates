using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    public partial class Insurers : BaseForm
    {
        private List<IInsurer> _currentInsurers;
        private List<IInsurer> _reccommendedInsurers;

        //public List<string> CurrentInsurers;        
        //public List<string> RecommendedInsurers;    
    
        public List<IInsurer> SelectedCurrentInsurers;
        public List<IInsurer> SelectedReccomendedInsurers;
        private bool validationErrors = false;

        public Insurers()
        {
            InitializeComponent();

           
            //CurrentInsurers = new List<string>();
            //RecommendedInsurers = new List<string>();
            SelectedCurrentInsurers = new List<IInsurer>();
            SelectedReccomendedInsurers = new List<IInsurer>();

          
        }

        private void Insurers_Load(object sender, EventArgs e)
        {
            dgvCurrent.AutoGenerateColumns = false;
            dgvRecommended.AutoGenerateColumns = false;

            _currentInsurers = new List<IInsurer>(); _currentInsurers = GetInsurers(BusinessLogic.Helpers.Constants.CacheNames.CurrentInsuerData);
           // _reccommendedInsurers = new List<IInsurer>(); _reccommendedInsurers.AddRange(_currentInsurers);
            _reccommendedInsurers = new List<IInsurer>(); _reccommendedInsurers = GetInsurers(BusinessLogic.Helpers.Constants.CacheNames.ReccomendedInsuerData);

            BindInsurers();

            dgvRecommended.CellMouseUp += new DataGridViewCellMouseEventHandler(dgvRecommended_CellMouseUp);
            dgvCurrent.CellMouseUp += new DataGridViewCellMouseEventHandler(dgvCurrent_CellMouseUp);
        }

        void dgvCurrent_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            PromptForInsurer(sender, e);
        }

        void dgvRecommended_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            PromptForInsurer(sender, e);
        }

        private void PromptForInsurer(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            dgv.BeginEdit(false);
            var formattedValue = false;
            var category = string.Empty;
            var value  = false;
            var percentIndex = -1;

            if (e.RowIndex < 0) return;

            foreach (DataGridViewCell cell in dgv.Rows[e.RowIndex].Cells)
            {
                switch (cell.OwningColumn.HeaderText)
                {
                    case "R":
                        formattedValue = cell.FormattedValue != null && (bool) cell.EditedFormattedValue;
                        if (cell.Value != null) bool.TryParse(cell.Value.ToString(), out value);
                        break;

                    case "C":
                        formattedValue = cell.FormattedValue != null && (bool)cell.EditedFormattedValue;
                         if(cell.Value != null) bool.TryParse(cell.Value.ToString(), out value);
                        break;

                    case "Category":
                        if (cell.Value != null) category = cell.Value.ToString();
                        break;

                    case "%":
                        percentIndex = cell.ColumnIndex;
                        break;
                }
            }

            if (!formattedValue)
            {
                //clear out the %.
                dgv.Rows[e.RowIndex].Cells[percentIndex].Value = "";
               
                return;
            } //formatted value is the becomming value of the checkbox (either getting unticked or getting ticked)

// ReSharper disable ConditionIsAlwaysTrueOrFalse
            if (value == formattedValue) return;
// ReSharper restore ConditionIsAlwaysTrueOrFalse

            if (String.Equals(category, "Third Party Broker", StringComparison.InvariantCultureIgnoreCase) ||
                String.Equals(category, "Underwriting Agency", StringComparison.InvariantCultureIgnoreCase) ||
                String.Equals(category, "Wholesale Broker", StringComparison.InvariantCultureIgnoreCase))
            {
                var enterInsurer = new EnterInsurer();
                enterInsurer.ShowDialog();
                var enteredInsurer = enterInsurer.EnteredInsurer;

                if (!String.IsNullOrEmpty(enteredInsurer))
                {
                    foreach (DataGridViewCell cell in dgv.Rows[e.RowIndex].Cells)
                    {
                        switch (cell.OwningColumn.HeaderText)
                        {
                            case "Insurer":
                                cell.Value += " / " + enteredInsurer;
                                break;
                        }
                    }
                }
            }

            dgv.Rows[e.RowIndex].Cells[percentIndex].Value = "100";


            dgv.EndEdit();
        }
    
        private List<IInsurer> GetInsurers(string type)
        {
            List<IInsurer> returnItems;
            if (Cache.Contains(type))
            {
                returnItems = (List<IInsurer>)Cache.Get(type);
            }
            else
            {
                var list = new SharePointList(Settings.Default.SharePointContextUrl, Settings.Default.ApprovedInsurersListName,BusinessLogic.Helpers.Constants.SharePointQueries.AllItemsSortByTitle );
                var presenter = new SharePointListPresenter(list, this);
                returnItems = presenter.GetInsurers();
                Cache.Add(type, returnItems, new CacheItemPolicy());
            }

            return returnItems;
        }

        private void BindInsurers()
        {
            if (_currentInsurers == null) return;
            dgvCurrent.DataSource = _currentInsurers;

            if (_reccommendedInsurers == null) return;
            dgvRecommended.DataSource = _reccommendedInsurers;

            dgvCurrent.ClearSelection();
            dgvRecommended.ClearSelection();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            validationErrors = false;
            GetSelectedInsurers(dgvRecommended,SelectedReccomendedInsurers);
            if (validationErrors) return;
            GetSelectedInsurers(dgvCurrent, SelectedCurrentInsurers);
            if (validationErrors) return;
            Close();
        }

        private void GetSelectedInsurers(DataGridView dataGridView, List<IInsurer> storageCollection)
        {
            storageCollection.Clear(); //clear any storage incase rerun.  ie- % validation error


            decimal percentTotal = 0;

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                var selected = false;
                var insurer = String.Empty;
                var category = string.Empty;
                var id = string.Empty;
                decimal percent = 0;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    switch (cell.OwningColumn.HeaderText)
                    {
                        case "C":
                            selected = cell.FormattedValue != null && (bool) cell.FormattedValue;
                            break;

                        case "R":
                            selected = cell.FormattedValue != null && (bool) cell.FormattedValue;
                            break;

                        case "Insurer":
                            if (cell.Value != null) insurer = cell.Value.ToString();
                            break;

                        case "Category":
                            if (cell.Value != null) category = cell.Value.ToString();
                            break;

                        case "Id":
                            if (cell.Value != null) id = cell.Value.ToString();
                            break;

                        case "%":
                            if (cell.EditedFormattedValue != null) decimal.TryParse(cell.EditedFormattedValue.ToString(), out percent);
                            percentTotal += percent;
                            break;
                    }
                }

                if (!selected) continue;
                var i = new Insurer
                    {
                        Category = category,
                        Id = id,
                        Percent = percent,
                        Title = insurer
                    };
                storageCollection.Add(i);
            }

            if (percentTotal != 100)
            {
                MessageBox.Show(@"Your selected insurers do not total 100%, Please correct the % for the seleted insurers.");
                validationErrors = true;
            }

            if (storageCollection.Count < 1)
            {
                MessageBox.Show(@"Please ensure you have selected at least 1 insurer.");
                validationErrors = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            SelectedCurrentInsurers.Clear();
            SelectedReccomendedInsurers.Clear();
            Close();
        }

        private void txtFindCurrent_TextChanged(object sender, EventArgs e)
        {
            FindResult(txtFindCurrent.Text, dgvCurrent, _currentInsurers, BusinessLogic.Helpers.Constants.CacheNames.CurrentInsuerData);
        }

        private void txtFindRecommended_TextChanged(object sender, EventArgs e)
        {
            FindResult(txtFindRecommended.Text, dgvRecommended, _reccommendedInsurers, BusinessLogic.Helpers.Constants.CacheNames.ReccomendedInsuerData);
        }

        private void FindResult(string find, DataGridView datagrid, List<IInsurer> insurers, string type)
        {
            var displayRow = false;
            datagrid.CurrentCell = null;
            
            if (String.IsNullOrEmpty(find))
            {
                displayRow = true;
            }
            
            foreach (DataGridViewRow r in datagrid.Rows)
            {
                if (displayRow == true)
                {
                    r.Visible = true; continue;
                }
                r.Visible = r.Cells[1].Value.ToString().ToUpperInvariant().Contains(find.ToUpperInvariant());
            }
            //this is faster but we cannot rebind datagrid as we lose entered values
            //if (String.IsNullOrEmpty(find))
            //{
            //    datagrid.DataSource = GetInsurers(type);
            //}
            //else
            //{
            //    var found = insurers.FindAll(i => i.Title.ToLowerInvariant().Contains(find.ToLowerInvariant()));
            //    datagrid.DataSource = found;
            //}
           // datagrid.ClearSelection();
        }
    }
}
