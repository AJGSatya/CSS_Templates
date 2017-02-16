using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OAMPS.Office.Word.Models.ActiveDirectory;

namespace OAMPS.Office.Word.Views.Wizards.Popups
{
    public partial class PeoplePicker : BaseWizardForm
    {
        private static List<Principal> _cachedUsers;

        private static PrincipalSearcher _searcher;
        public UserPrincipalEx SelectedUser = null;
        private List<Principal> _users;

        public PeoplePicker(string name, Form owner)
        {
            InitializeComponent();
            txtFind.KeyDown += txtFind_KeyDown;
            lstUsers.DoubleClick += lstUsers_DoubleClick;
            Owner = owner;

            Setup(name);
        }

        public static List<Principal> CachedUsers
        {
            get { return _cachedUsers; }
        }

        private void Setup(string name)
        {
            txtFind.Text = name;
            if (String.IsNullOrEmpty(name))
                return;
            //TODO TEST loading from cache
            FindOld();
        }

        private void lstUsers_DoubleClick(object sender, EventArgs e)
        {
            ReturnSelectedUser();
        }

        private void ReturnSelectedUser()
        {
            SelectedUser = (UserPrincipalEx) lstUsers.SelectedItem;
            Close();
        }

        private void txtFind_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //FindUsersFromCache();
                FindOld();
            }
        }

        //private void PeoplePicker_Load(object sender, EventArgs e)
        //{

        //    if (String.IsNullOrEmpty(txtFind.Text))
        //        return;
        //    //TODO TEST loading from cache
        //    FindOld();
        //    //FindUsersFromCache();
        //    //   var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
        //    //Task.Factory.StartNew(() => LoadUsers(uiScheduler));
        //    // LoadUsers(uiScheduler);
        //}

        private void Find()
        {
            if (_users == null) return;

            var foundUsers = new List<Principal>();

            if (String.IsNullOrEmpty(txtFind.Text))
                foundUsers = _users;
            else
            {
                foundUsers.AddRange(
                    _users.Where(u => !String.IsNullOrEmpty(u.DisplayName))
                          .Where(u => u.DisplayName.Contains(txtFind.Text)));
            }

            lstUsers.DataSource = foundUsers;
            lstUsers.DisplayMember = "displayName";
            lstUsers.ValueMember = "displayName";
        }

        private void FindOld()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                if (Owner != null)
                    Owner.Cursor = Cursors.WaitCursor;

                using (var context = new PrincipalContext(ContextType.Domain))
                {
                    using (var userPrincipal = new UserPrincipalEx(context) {Enabled = true})
                    {
                        if (!String.IsNullOrEmpty(txtFind.Text))
                        {
                            userPrincipal.DisplayName = "*" + txtFind.Text + "*";
                        }
                        using (var searcher = new PrincipalSearcher(userPrincipal))
                        {
                            searcher.QueryFilter = userPrincipal;
                            _users = searcher.FindAll().ToList();
                            lstUsers.DataSource = _users;
                            lstUsers.DisplayMember = "displayName";
                            lstUsers.ValueMember = "displayName";
                            //if only 1 row returned=> auto select it
                            if (_users.Count == 1)
                                SelectedUser = (UserPrincipalEx) _users.FirstOrDefault();
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
        }

        public void FindUsersFromCache()
        {
            if (_cachedUsers != null)
            {
                Principal results = _cachedUsers.Find(x => x.DisplayName.Contains(txtFind.Text));
                lstUsers.DataSource = results;
                lstUsers.DisplayMember = "displayName";
                lstUsers.ValueMember = "displayName";
            }
        }


        public static void LoadUsersIntoCache()
        {
            TaskScheduler uiScheduler = TaskScheduler.Default;
            Task.Factory.StartNew(() =>
                {
                    if (_cachedUsers == null)
                    {
                        using (var context = new PrincipalContext(ContextType.Domain))
                        {
                            using (var userPrincipal = new UserPrincipalEx(context) {Enabled = true})
                            {
                                using (var searcher = new PrincipalSearcher(userPrincipal))
                                {
                                    _searcher = searcher;
                                    searcher.QueryFilter = userPrincipal;

                                    List<Principal> breakpoint = _searcher.FindAll().ToList();
                                    _cachedUsers = breakpoint;
                                }
                            }
                        }
                    }
                }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
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
            //FindUsersFromCache();
            FindOld();
        }
    }
}