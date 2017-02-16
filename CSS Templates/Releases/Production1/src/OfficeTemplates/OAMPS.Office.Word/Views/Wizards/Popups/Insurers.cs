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
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;
using OAMPS.Office.Word.Models.SharePoint;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    public partial class Insurers : BaseForm
    {
        private List<IInsurer> _currentInsurers;
        private List<IInsurer> _reccommendedInsurers;
        public string CurrentInsurer;
        public string RecommendedInsurer;
        public string CurrentInsurerId;
        public string RecommendedInsurerId;

        bool dontRunHandler;



        public Insurers()
        {
            InitializeComponent();
        }

        void dgvRecommended_SelectionChanged(object sender, EventArgs e)
        {
            if (dontRunHandler) return;

            if (dgvRecommended.SelectedCells.Count == 0)
                return;

            var selectedInsurer = dgvRecommended.Rows[dgvRecommended.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            var insurerCategory =
                dgvRecommended.Rows[dgvRecommended.SelectedCells[0].RowIndex].Cells[2].Value.ToString();

            if (String.Equals(insurerCategory, "Third Party Broker", StringComparison.InvariantCultureIgnoreCase) ||
              String.Equals(insurerCategory, "Underwriting Agency", StringComparison.InvariantCultureIgnoreCase) ||
              String.Equals(insurerCategory, "Wholesale Broker", StringComparison.InvariantCultureIgnoreCase))
            {
                var enterInsurer = new EnterInsurer();
                enterInsurer.ShowDialog();
                var enteredInsurer = enterInsurer.EnteredInsurer;

                if (!String.IsNullOrEmpty(enteredInsurer))
                {
                    selectedInsurer = selectedInsurer + " / " + enteredInsurer;
                }
            }

            RecommendedInsurer = selectedInsurer;
            RecommendedInsurerId = dgvRecommended.Rows[dgvRecommended.SelectedCells[0].RowIndex].Cells[1].Value.ToString();

        }

        void dgvCurrent_SelectionChanged(object sender, EventArgs e)
        {
            if (dontRunHandler) return;

            if (dgvCurrent.SelectedCells.Count == 0)
                return;

            var selectedInsurer = dgvCurrent.Rows[dgvCurrent.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            var insurerCategory = dgvCurrent.Rows[dgvCurrent.SelectedCells[0].RowIndex].Cells[2].Value.ToString();

            if (String.Equals(insurerCategory, "Third Party Broker", StringComparison.InvariantCultureIgnoreCase) ||
                String.Equals(insurerCategory, "Underwriting Agency", StringComparison.InvariantCultureIgnoreCase)||
                String.Equals(insurerCategory, "Wholesale Broker", StringComparison.InvariantCultureIgnoreCase))
            {
                var enterInsurer = new EnterInsurer();
                enterInsurer.ShowDialog();
                var enteredInsurer = enterInsurer.EnteredInsurer;

                if (!String.IsNullOrEmpty(enteredInsurer))
                {
                    selectedInsurer = selectedInsurer + " / " + enteredInsurer;
                }
            }

            CurrentInsurer = selectedInsurer;
            CurrentInsurerId = dgvCurrent.Rows[dgvCurrent.SelectedCells[0].RowIndex].Cells[1].Value.ToString();
        }

        private void Insurers_Load(object sender, EventArgs e)
        {
            _currentInsurers = new List<IInsurer>(); _currentInsurers = GetInsurers();
            _reccommendedInsurers = new List<IInsurer>(); _reccommendedInsurers.AddRange(_currentInsurers);
            BindInsurers();

            dgvCurrent.SelectionChanged += new EventHandler(dgvCurrent_SelectionChanged);
            dgvRecommended.SelectionChanged += new EventHandler(dgvRecommended_SelectionChanged);
        }

        private List<IInsurer> GetInsurers()
        {
            List<IInsurer> returnItems;
            if (Cache.Contains("InsurersData"))
            {
                returnItems = (List<IInsurer>)Cache.Get("InsurersData");
            }
            else
            {
                var list = new SharePointList(Settings.Default.SharePointContextUrl,
                                              Settings.Default.ApprovedInsurersListName);
                var presenter = new SharePointListPresenter(list, this);
                returnItems = presenter.GetInsurers();
                Cache.Add("InsurersData", returnItems, new CacheItemPolicy());
            }

            return returnItems;
        }

        private void BindInsurers()
        {
            if (_currentInsurers == null) return;

            dgvCurrent.DataSource = _currentInsurers;
            dgvRecommended.DataSource = _reccommendedInsurers;

            dgvCurrent.ClearSelection();
            dgvRecommended.ClearSelection();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
            RecommendedInsurer = null;
            CurrentInsurer = null;
        }

        private void txtFindCurrent_TextChanged(object sender, EventArgs e)
        {
            dontRunHandler = true;
            if (String.IsNullOrEmpty(txtFindCurrent.Text))
            {
                dgvCurrent.DataSource = GetInsurers();
            }
            else
            {
                var found = _currentInsurers.FindAll(i => i.Title.StartsWith(txtFindCurrent.Text,StringComparison.OrdinalIgnoreCase));
                dgvCurrent.DataSource = found;
            }
           // dgvCurrent.Refresh();
            dgvCurrent.ClearSelection();

            dontRunHandler = false;
        }

        private void txtFindRecommended_TextChanged(object sender, EventArgs e)
        {
            dontRunHandler = true;

            dgvRecommended.ClearSelection();
            if (String.IsNullOrEmpty(txtFindRecommended.Text))
            {
                dgvRecommended.DataSource = GetInsurers();
            }
            else
            {
                var found = _reccommendedInsurers.FindAll(i => i.Title.StartsWith(txtFindRecommended.Text,StringComparison.OrdinalIgnoreCase));
                dgvRecommended.DataSource = found;
            }
         //   dgvRecommended.Refresh();
            dgvRecommended.ClearSelection();

            dontRunHandler = false;

        }

    }
}
