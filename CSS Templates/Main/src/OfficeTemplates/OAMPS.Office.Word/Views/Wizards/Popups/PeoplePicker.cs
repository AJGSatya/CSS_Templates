using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Models.ActiveDirectory;
using OAMPS.Office.Word.Models.Local;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    public partial class PeoplePicker : BaseWizardForm
    {
        private static PrincipalSearcher _searcher;
        private List<Principal> _users;
        public UserPrincipalEx SelectedUser;

        public PeoplePicker(string name, Form owner)
        {
            InitializeComponent();
            txtFind.KeyDown += txtFind_KeyDown;
            lstUsers.DoubleClick += lstUsers_DoubleClick;
            Owner = owner;

            Setup(name);
        }

        public static List<Principal> CachedUsers { get; private set; }

        private void Setup(string name)
        {
            txtFind.Text = name;
            if (string.IsNullOrEmpty(name))
                return;
            Find();
        }

        private void lstUsers_DoubleClick(object sender, EventArgs e)
        {
            ReturnSelectedUser();
        }

        private void ReturnSelectedUser()
        {
            SelectedUser = (UserPrincipalEx)lstUsers.SelectedItem;
            Close();
        }

        private void txtFind_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Find();
            }
        }


        private void Find()
        {
            var searchString = txtFind.Text;
            if (string.IsNullOrEmpty(searchString))
                return;

            //var sw = System.Diagnostics.Stopwatch.StartNew();

            _users = LocalCache.Get<List<Principal>>("PeoplePicker:" + searchString, () => getResults(searchString));

            lstUsers.DisplayMember = "displayName";
            lstUsers.ValueMember = "displayName";
            lstUsers.DataSource = _users;
            //if only 1 row returned=> auto select it
            if (_users.Count == 1)
                SelectedUser = (UserPrincipalEx)_users.FirstOrDefault();

            //sw.Stop();
            //ErrorLog.TraceLog("PeoplePicker.FindOld time taken {0}ms, searched for {1}", sw.ElapsedMilliseconds, txtFind.Text);
            //ErrorLog.TraceLog("PeoplePicker.FindOld user retrieved: {0}", string.Join(";", _users.Select(x => x.DisplayName).ToList()));
        }

        private List<Principal> getResults(string searchString)
        {
            var users = new List<Principal>();
            try
            {
                Cursor = Cursors.WaitCursor;
                if (Owner != null)
                    Owner.Cursor = Cursors.WaitCursor;
                
                using (var context = new PrincipalContext(ContextType.Domain, Settings.Default.PeoplePickerSearchDomain, Settings.Default.PeoplePickerSearchOU))
                {
                    using (var userPrincipal = new UserPrincipalEx(context) { Enabled = true })
                    {
                        userPrincipal.DisplayName = "*" + searchString + "*";
                        using (var searcher = new PrincipalSearcher(userPrincipal))
                        {
                            searcher.QueryFilter = userPrincipal;
                            users = searcher.FindAll().ToList();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OnError(ex);
            }
            finally
            {
                Cursor = Cursors.Default;
                if (Owner != null)
                    Owner.Cursor = Cursors.Default;
            }
            return users;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            ReturnSelectedUser();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void lblFind_Click(object sender, EventArgs e)
        {
            Find();
        }
    }
}