﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using QlikviewEnhancedUserControl.ServiceReference;
using QlikviewEnhancedUserControl.ServiceSupport;
using BrightIdeasSoftware;
using System.Threading;
using System.Configuration;
using System.DirectoryServices.AccountManagement;


namespace QlikviewEnhancedUserControl
{
    public partial class Form1 : Form
    {
        QMSClient Client;
        List<string> QVS = new List<string>();
        DataTable dtSelectedUSers = new DataTable();
        DataTable dtTasks = new DataTable();
        DataTable dtUserDocuments = new DataTable();
        DataTable dtUsersAndDocs = new DataTable();
        DocumentNode dtDocSource = new DocumentNode();
        DocumentNode dtDocTarget = new DocumentNode();

        bool sourceAdded = false;
        bool targetAdded = false;
        Guid qvsId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid dscId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid qdsId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid remoteQvsId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid remoteDscId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid remoteQdsId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid remoteQmsId = new Guid("00000000-0000-0000-0000-000000000000");
        bool loadingVisible = false;
        string key = "";
        string qms = "";
        string server = "";
        TabPage tb1 = null;
        TabPage tb2 = null;
        TabPage tb3 = null;
        TabPage tb4 = null;
        int totalUserDocs = 0;
        int SelectedTab = 0;
        //BrightIdeasSoftware.OLVColumn deleteColumn = new BrightIdeasSoftware.OLVColumn();



        public Form1()
        {
            InitializeComponent();
            tb1 = tabControl1.TabPages[1];
            tb2 = tabControl1.TabPages[2];
            tb3 = tabControl1.TabPages[3];
            tb4 = tabControl1.TabPages[4];

            tabControl1.TabPages.Remove(tb1);
            tabControl1.TabPages.Remove(tb2);
            tabControl1.TabPages.Remove(tb3);
            tabControl1.TabPages.Remove(tb4);
            //DataTable dtServices = new DataTable();
            //DataColumn sID = new DataColumn("ID");
            //DataColumn sName = new DataColumn("Name");
            //DataColumn sType = new DataColumn("Type");
            //DataColumn sNodeCount = new DataColumn("NodeCount");
            //DataColumn sAddress = new DataColumn("Address");            
            //dtServices.Columns.Add(sName);
            //dtServices.Columns.Add(sType);
            //dtServices.Columns.Add(sAddress);
            //dtServices.Columns.Add(sNodeCount);
            //dtServices.Columns.Add(sID);    
        }


        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {

            //MessageBox.Show(dataGridView2.SelectedRows.Count.ToString());
            /* foreach (DataGridViewRow row in dataGridView2.SelectedRows)
             {
                 Guid[] ids = new Guid[1];
                 ids[0] = new Guid(row.Cells[4].Value.ToString());
                 var stat = Client.GetServiceStatuses(ids);

                 DataTable dtStats = new DataTable();
                 DataColumn dcMMessage = new DataColumn("Message");
                 DataColumn dcMId = new DataColumn("ID");
                 DataColumn dcMHost = new DataColumn("Host");
                 DataColumn dcMStatus = new DataColumn("Status");
                 dtStats.Columns.Add(dcMHost);
                 dtStats.Columns.Add(dcMStatus);
                 dtStats.Columns.Add(dcMMessage);
                 dtStats.Columns.Add(dcMId);

                 for (var i = 0; i < stat[0].MemberStatusDetails.Length; i++)
                 {

                     DataRow dr = dtStats.NewRow();
                     if (stat[0].MemberStatusDetails[i].Message.Length > 0)
                     {
                         dr["Message"] = stat[0].MemberStatusDetails[i].Message[0].ToString();
                     }
                     else
                     {
                         dr["Message"] = "no message";
                     }

                     dr["Status"] = stat[0].MemberStatusDetails[i].Status.ToString();
                     dr["Host"] = stat[0].MemberStatusDetails[i].Host.ToString();
                     dr["ID"] = stat[0].MemberStatusDetails[i].ID.ToString();
                     dtStats.Rows.Add(dr);                    
                 }

                 dataGridView3.DataSource = dtStats;

                

             }*/
        }

        private DataTable GetServicesStatuses()
        {
            DataTable dtStats = new DataTable();
            DataColumn dcMMessage = new DataColumn("Message");
            DataColumn dcMId = new DataColumn("MemberID");
            DataColumn dcMHost = new DataColumn("Host");
            DataColumn dcMStatus = new DataColumn("Status");
            DataColumn sID = new DataColumn("ID");
            DataColumn sName = new DataColumn("Name");
            DataColumn sType = new DataColumn("Type");
            DataColumn sNodeCount = new DataColumn("NodeCount");
            DataColumn sAddress = new DataColumn("Address");
            dtStats.Columns.Add(sName);
            dtStats.Columns.Add(sType);
            dtStats.Columns.Add(sAddress);
            //dtStats.Columns.Add(sNodeCount);            
            dtStats.Columns.Add(dcMHost);
            dtStats.Columns.Add(dcMStatus);
            dtStats.Columns.Add(dcMMessage);
            dtStats.Columns.Add(sID);
            dtStats.Columns.Add(dcMId);

            List<ServiceInfo> myServices = Client.GetServices(ServiceTypes.All);
            ServiceInfo remoteQMSService = Client.GetServices(ServiceTypes.RemoteQlikViewManagementService).FirstOrDefault();
            List<ServiceInfo> remoteServices = Client.RemoteGetServices(remoteQMSService.ID, ServiceTypes.All);
            remoteQmsId = remoteQMSService.ID;

            foreach (ServiceInfo service in remoteServices)
            {
                var t = service.Type;
                if (service.Type == ServiceTypes.QlikViewServer)
                {
                    remoteQvsId = service.ID;
                }

                if (service.Type == ServiceTypes.QlikViewDirectoryServiceConnector)
                {
                    remoteDscId = service.ID;
                }

                if (service.Type == ServiceTypes.QlikViewDistributionService)
                {
                    remoteQdsId = service.ID;
                }
            }            


            foreach (ServiceInfo service in myServices)
            {
                List<Guid> ids = new List<Guid>();
                ids.Add(service.ID);
                var stat = Client.GetServiceStatuses(ids);

                if (service.Type == ServiceTypes.QlikViewServer)
                {
                    qvsId = service.ID;
                }

                if (service.Type == ServiceTypes.QlikViewDirectoryServiceConnector)
                {
                    dscId = service.ID;
                }

                if (service.Type == ServiceTypes.QlikViewDistributionService)
                {
                    qdsId = service.ID;
                }

                //DataTable dtStats = new DataTable();

                if (stat.Count > 0)
                {

                    for (var i = 0; i < stat[0].MemberStatusDetails.Count; i++)
                    {

                        if (service.Type == ServiceTypes.QlikViewServer)
                        {
                            QVS.Add(stat[0].MemberStatusDetails[i].ID.ToString() + '_' + stat[0].MemberStatusDetails[i].Host);
                        }

                        DataRow dr = dtStats.NewRow();
                        //DataRow dr = dtServices.NewRow();
                        dr["ID"] = service.ID;
                        dr["Name"] = service.Name;
                        dr["Type"] = service.Type;
                        //dr["NodeCount"] = service.ClusterNodeCount;
                        dr["Address"] = service.Address;
                        //dtStats.Rows.Add(dr);

                        if (stat[0].MemberStatusDetails[i].Message.Count > 0)
                        {
                            dr["Message"] = stat[0].MemberStatusDetails[i].Message[0].ToString();
                        }
                        else
                        {
                            dr["Message"] = "no message";
                        }

                        dr["Status"] = stat[0].MemberStatusDetails[i].Status.ToString();
                        dr["Host"] = stat[0].MemberStatusDetails[i].Host.ToString();
                        dr["MemberID"] = stat[0].MemberStatusDetails[i].ID.ToString();
                        dtStats.Rows.Add(dr);
                    }
                }
            }

            return dtStats;
        }

        private DataTable GetActiveUserAndDocs(List<string> members)
        {
            DataTable dt = new DataTable();
            DataColumn TimeStamp = new DataColumn("timestamp");
            DataColumn Server = new DataColumn("server");
            DataColumn Document = new DataColumn("document");
            DataColumn UserId = new DataColumn("userid");
            dt.Columns.Add(Server);
            dt.Columns.Add(Document);
            dt.Columns.Add(UserId);
            dt.Columns.Add(TimeStamp);

            for (int m = 0; m < members.Count; m++)
            {
                var a = new Guid(members[m].Substring(0, members[m].IndexOf('_')));

                var t = new Dictionary<string, List<string>>();

                try
                {
                    t = Client.GetQVSDocumentsAndUsers(a, QueryTarget.ClusterMember);
                }
                catch (System.Exception ex)
                {
                    key = Client.GetTimeLimitedServiceKey();
                    ServiceKeyClientMessageInspector.ServiceKey = key;
                    t = Client.GetQVSDocumentsAndUsers(a, QueryTarget.ClusterMember);
                }

                foreach (KeyValuePair<string, List<string>> entry in t)
                {
                    var doc = entry.Key;
                    for (int b = 0; b < entry.Value.Count; b++)
                    {
                        var executiondate = DateTime.Now.ToString("yyyyMMddHHmmss");

                        DataRow dr = dt.NewRow();
                        dr["timestamp"] = executiondate.ToString();
                        dr["server"] = members[m].Substring(members[m].IndexOf('_') + 1, members[m].Length - members[m].IndexOf('_') - 1).Trim();
                        dr["document"] = doc;
                        dr["userid"] = entry.Value[b];
                        dt.Rows.Add(dr);
                    }
                }
            }

            return dt;
        }

        private DataTable GetUserDocuments()
        {
            //var qvsId = new Guid("ea7e3a82-3693-4fee-9ab2-3e9dd8a67148");

            DataTable dtDocs = new DataTable();
            DataColumn dName = new DataColumn("Name");
            DataColumn dPath = new DataColumn("RelativePath");
            DataColumn dType = new DataColumn("Type");
            DataColumn isOrphan = new DataColumn("IsOrphan");
            DataColumn isSubfolder = new DataColumn("IsSubFolder");
            DataColumn dTaskCount = new DataColumn("TaskCount");
            DataColumn dId = new DataColumn("ID");
            DataColumn fId = new DataColumn("FolderID");
            dType.DataType = typeof(DocumentType);
            dType.DataType = typeof(DocumentType);
            dtDocs.Columns.Add(dName);
            dtDocs.Columns.Add(dPath);
            dtDocs.Columns.Add(dType);
            dtDocs.Columns.Add(isOrphan);
            dtDocs.Columns.Add(isSubfolder);
            dtDocs.Columns.Add(dTaskCount);
            dtDocs.Columns.Add(fId);
            dtDocs.Columns.Add(dId);

            var userDocs = new List<DocumentNode>();
            try
            {
                Client.ClearQVSCache(QVSCacheObjects.UserDocumentList);
                userDocs = Client.GetUserDocuments(qvsId);                
            }
            catch (System.Exception ex)
            {
                key = Client.GetTimeLimitedServiceKey();
                ServiceKeyClientMessageInspector.ServiceKey = key;
                Client.ClearQVSCache(QVSCacheObjects.UserDocumentList);
                userDocs = Client.GetUserDocuments(qvsId);
            }

            foreach (var userDoc in userDocs)
            {
                var a = userDoc;
                DataRow dr = dtDocs.NewRow();
                dr["FolderID"] = a.FolderID;
                dr["ID"] = a.ID;
                dr["IsOrphan"] = a.IsOrphan;
                dr["IsSubFolder"] = a.IsSubFolder;
                dr["Name"] = a.Name;
                dr["RelativePath"] = a.RelativePath;
                dr["TaskCount"] = a.TaskCount;
                dr["Type"] = a.Type;
                dtDocs.Rows.Add(dr);
            }

            //dataListView2.DataSource = dtDocs;

            return dtDocs;
        }

        private DataTable GetDocumentsMetadata(DataTable dtDocs)
        {
            DataTable dt = dtDocs.Clone();
            DataColumn dcUser = new DataColumn("User");
            dt.Columns.Add(dcUser);
            dt.Columns["User"].SetOrdinal(1);

            if (SelectedTab == 3)
            {
                backgroundWorker1.ReportProgress(dtDocs.Rows.Count, "totalDocs");
            }

            for (var i = 0; i < dtDocs.Rows.Count; i++)
            {

                var documentNode = new DocumentNode();
                documentNode.Type = DocumentType.User;
                documentNode.FolderID = new Guid(dtDocs.Rows[i]["FolderID"].ToString());
                documentNode.ID = new Guid(dtDocs.Rows[i]["ID"].ToString());
                documentNode.Name = dtDocs.Rows[i]["Name"].ToString();
                documentNode.IsOrphan = Convert.ToBoolean(dtDocs.Rows[i]["IsOrphan"]);
                documentNode.IsSubFolder = Convert.ToBoolean(dtDocs.Rows[i]["IsSubFolder"]);
                documentNode.RelativePath = dtDocs.Rows[i]["RelativePath"].ToString();
                documentNode.TaskCount = Convert.ToInt32(dtDocs.Rows[i]["TaskCount"]);

                DocumentMetaData t = new DocumentMetaData();

                try
                {
                    t = Client.GetDocumentMetaData(documentNode, DocumentMetaDataScope.Authorization);
                }
                catch (System.Exception ex)
                {
                    key = Client.GetTimeLimitedServiceKey();
                    ServiceKeyClientMessageInspector.ServiceKey = key;
                    t = Client.GetDocumentMetaData(documentNode, DocumentMetaDataScope.Authorization);
                }

                if (t.Authorization.Access.Count > 0)
                {
                    for (var b = 0; b < t.Authorization.Access.Count; b++)
                    {
                        DataRow dr = dt.NewRow();
                        dr["User"] = t.Authorization.Access[b].UserName;
                        dr["FolderID"] = dtDocs.Rows[i]["FolderID"];
                        dr["ID"] = dtDocs.Rows[i]["ID"];
                        dr["IsOrphan"] = dtDocs.Rows[i]["IsOrphan"];
                        dr["IsSubFolder"] = dtDocs.Rows[i]["IsSubFolder"];
                        dr["Name"] = dtDocs.Rows[i]["Name"];
                        dr["RelativePath"] = dtDocs.Rows[i]["RelativePath"];
                        dr["TaskCount"] = dtDocs.Rows[i]["TaskCount"];
                        dr["Type"] = dtDocs.Rows[i]["Type"];
                        dt.Rows.Add(dr);
                    }
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    dr["User"] = "NONE";
                    dr["FolderID"] = dtDocs.Rows[i]["FolderID"];
                    dr["ID"] = dtDocs.Rows[i]["ID"];
                    dr["IsOrphan"] = dtDocs.Rows[i]["IsOrphan"];
                    dr["IsSubFolder"] = dtDocs.Rows[i]["IsSubFolder"];
                    dr["Name"] = dtDocs.Rows[i]["Name"];
                    dr["RelativePath"] = dtDocs.Rows[i]["RelativePath"];
                    dr["TaskCount"] = dtDocs.Rows[i]["TaskCount"];
                    dr["Type"] = dtDocs.Rows[i]["Type"];
                    dt.Rows.Add(dr);
                }

                if (SelectedTab == 3)
                {
                    backgroundWorker1.ReportProgress(i, "processeddocs");
                }
            }


            return dt;
        }

        private void GetActiveUsers()
        {
            DataTable dt = GetActiveUserAndDocs(QVS);
            dataListView2.DataSource = dt;

            for (var i = 0; i < dataListView2.Columns.Count; i++)
            {
                dataListView2.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetActiveUsers();

            numericUpDown1.Enabled = true;
            button16.Enabled = true;
            button16.Enabled = true;
            textBox4.Enabled = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            RefreshServices();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //timer1.Start();
            try
            {
                dataListView4.Clear();
                dataListView5.Clear();
                //dataListView10.Clear();

                DataTable dt = GetUserDocuments();
                dataListView4.DataSource = dt;
                //dataListView10.DataSource = dt;

                for (var i = 0; i < dataListView4.Columns.Count; i++)
                {
                    dataListView4.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                    //dataListView10.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }
                //DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                //dataGridView3.Columns.Add(chk);
                //chk.HeaderText = " ";
                //chk.Name = "chk";
                //dataGridView3.Columns[dataGridView3.Columns.Count-1].DisplayIndex = 0;
            }
            catch (System.Exception ex)
            {
                //key = Client.GetTimeLimitedServiceKey();
                //ServiceKeyClientMessageInspector.ServiceKey = key;
            }
            finally
            {
                //timer1.Stop();
            }

        }

        private void RefreshServices()
        {
            label3.Text = "";

            try
            {
                Client = new QMSClient("BasicHttpBinding_IQMS", qms);
                key = Client.GetTimeLimitedServiceKey();
                ServiceKeyClientMessageInspector.ServiceKey = key;

                DataTable dt = GetServicesStatuses();
                dataListView1.DataSource = dt;

                for (var i = 0; i < dataListView1.Columns.Count; i++)
                {
                    dataListView1.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }

                tabControl1.TabPages.Remove(tb1);
                tabControl1.TabPages.Remove(tb2);
                tabControl1.TabPages.Remove(tb3);
                tabControl1.TabPages.Remove(tb4);
                tabControl1.TabPages.Add(tb1);
                tabControl1.TabPages.Add(tb2);
                tabControl1.TabPages.Add(tb3);
                tabControl1.TabPages.Add(tb4);
            }
            catch (System.Exception ex)
            {
                dataListView1.Clear();
                tabControl1.TabPages.Remove(tb1);
                tabControl1.TabPages.Remove(tb2);
                tabControl1.TabPages.Remove(tb3);
                tabControl1.TabPages.Remove(tb4);
                label3.Text = "Cannot establish connection to the server.";
            }
        }

        /*private void exportNewData()
        {

            StringBuilder sb = new StringBuilder();

            string columnsHeader = "";

            for (int i = 0; i < dataGridView3.Columns.Count - 1; i++)
            { columnsHeader += dataGridView3.Columns[i].Name + ","; }

            sb.Append(columnsHeader + Environment.NewLine);

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (!row.IsNewRow)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        var val = "";
                        try
                        {
                            val = row.Cells[i].Value.ToString().Replace(",", " ");
                        }
                        catch (System.Exception ex)
                        {
                            val = "";
                        }
                        sb.Append(val + ","); }

                    sb.Append(Environment.NewLine);
                }
            }

            exportFileBrowser(sb);
        }*/

        private void exportFileBrowser(StringBuilder sb)
        {
            StreamWriter sw;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "CSV files (*.csv)|*.csv";
            sfd.InitialDirectory = @"s:\development\stefan\";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sw = new StreamWriter(sfd.FileName, false);
                using (sw)
                { sw.WriteLine(sb.ToString()); }

                MessageBox.Show("CSV file saved.");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //exportNewData();
            //OLVExporter ex = new OLVExporter(dataListView4, dataListView4.FilteredObjects);
            //string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
        }

        private void UserSearch()
        {
            dataListView6.Clear();

            DataTable dt = new DataTable();
            DataColumn dcId = new DataColumn("Id");
            DataColumn dcName = new DataColumn("Name");
            dt.Columns.Add(dcId);
            dt.Columns.Add(dcName);


            var users = new List<string>();
            users.Add(textBox1.Text);

            var u = textBox1.Text.Split(';');
            users = u.ToList<string>();

            foreach (var user in users)
            {
                var user1 = new List<string>();
                user1.Add(user.Trim());
                var userDetails = Client.LookupNames(dscId, user1);

                if (userDetails.Count > 0)
                {
                    for (var i = 0; i < userDetails.Count; i++)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Id"] = userDetails[i].Name;
                        dr["Name"] = userDetails[i].OtherProperty;
                        dt.Rows.Add(dr);
                    }
                }
            }

            dataListView6.DataSource = dt;
            for (var i = 0; i < dataListView6.Columns.Count; i++)
            {
                dataListView6.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            UserSearch();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddUserToList();

            //dataListView7.Columns.Add(deleteColumn);
        }

        private void AddUserToList()
        {
            foreach (DataRowView row in dataListView6.SelectedObjects)
            {
                bool present = false;

                for (int i = 0; i < dtSelectedUSers.Rows.Count; i++)
                {
                    if (dtSelectedUSers.Rows[i]["ID"].ToString() == row.Row["ID"].ToString())
                    {
                        present = true;
                        break;
                    }
                }

                if (present == false)
                {
                    DataRow dr = dtSelectedUSers.NewRow();
                    dr["Id"] = row.Row["ID"];
                    dr["Name"] = row.Row["Name"];
                    dtSelectedUSers.Rows.Add(dr);
                }
            }

            dataListView7.DataSource = dtSelectedUSers;
            for (var i = 0; i < dataListView7.Columns.Count; i++)
            {
                dataListView7.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);

            }
        }

        private void FilterDocs()
        {
            try
            {
                var field = "";
                if (rbtn_doc.Checked == true)
                {
                    field = "Name";
                }
                else
                {
                    field = "RelativePath";
                }

                try
                {
                    //(dataGridView3.DataSource as DataTable).DefaultView.RowFilter = string.Concat(field, " LIKE '%", textBox2.Text, "%'"); //"RelativePath LIKE '%[,]" + textBox2.Text + "[,]%'";            
                    (dataListView4.DataSource as DataTable).DefaultView.RowFilter = string.Concat(field, " LIKE '%", textBox2.Text, "%'"); //"RelativePath LIKE '%[,]" + textBox2.Text + "[,]%'";            
                }
                catch (System.Exception ex)
                {
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FilterDocs();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox2.Text = String.Empty;
            //(dataGridView3.DataSource as DataTable).DefaultView.RowFilter = string.Empty;
            (dataListView4.DataSource as DataTable).DefaultView.RowFilter = string.Empty;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //MessageBox.Show(tabControl1.SelectedIndex.ToString());
            label2.Text = "";
            label4.Text = "";
            DataColumn dcId = new DataColumn("Id");
            DataColumn dcName = new DataColumn("Name");
            dtSelectedUSers.Columns.Add(dcId);
            dtSelectedUSers.Columns.Add(dcName);
            //dtSelectedUSers.Columns.Add(deleteColumn);            

            dtTasks.Columns.Add("Name");
            dtTasks.Columns.Add("DocNodeName");
            dtTasks.Columns.Add("DocNodeFolderId");
            dtTasks.Columns.Add("DocNodeRelativePath");
            dtTasks.Columns.Add("DocName");
            dtTasks.Columns.Add("Category");
            dtTasks.Columns.Add("RelativePath");
            dtTasks.Columns.Add("TaskId");
            dtTasks.Columns.Add("DocId");


            numericUpDown1.Enabled = false;
            button16.Enabled = false;
            button16.Enabled = false;
            textBox4.Enabled = false;
            numericUpDown2.Maximum = decimal.MaxValue;

            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.textBox1, "Single user or multiple users separated with ';'. Use '*' for wildcard");
            ToolTip1.SetToolTip(this.button12, "ADD user from the left to the documents on the right");
            ToolTip1.SetToolTip(this.button13, "REMOVE user from the left from the documents on the right");
            ToolTip1.SetToolTip(this.textBox3, "For example: http://localhost:4799/QMS/Service");
            ToolTip1.SetToolTip(this.checkBox1, "If checked - append the missing users to target from source. If not checked - target will have the same users as source.");

            label3.Text = "";
            rbtn_doc.Checked = true;
            btn_AddSource.Enabled = false;
            btn_AddTarget.Enabled = false;
            //chb_AutoGetMetadata.Checked = true;

            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
            qms = config.AppSettings.Settings["qms"].Value.ToString();
            textBox3.Text = qms;

            //if (config.AppSettings.Settings["activedocspath"].Value.ToString().Length == 0)
            //{
            //    textBox4.Text = System.Reflection.Assembly.GetEntryAssembly().Location;
            //}
            //else
            //{
            textBox4.Text = config.AppSettings.Settings["activedocspath"].Value.ToString();
            numericUpDown1.Value = Convert.ToInt32(config.AppSettings.Settings["activedocsinterval"].Value);
            //}            

            server = qms.Substring(qms.IndexOf("//") + 2, qms.Length - qms.IndexOf("/QMS"));
            //string qms = "http://localhost:4799/QMS/Service";
            if (qms.Trim().Length > 0)
            {
                Refresh();
                button2.Enabled = true;
            }

            //dataListView7.CellEditActivation = ObjectListView.CellEditActivateMode.SingleClick;
            //dataListView7.CellEditStarting += ObjectListView1OnCellEditStarting;

            //deleteColumn.IsEditable = true;
            //deleteColumn.AspectGetter = delegate
            //{
            //return "Delete";
            //};
        }

        private void ObjectListView1OnCellEditStarting(object sender, CellEditEventArgs e)
        {
            // special cell edit handling for our delete-row
            //if (e.Column == deleteColumn)
            {
                //    e.Cancel = true;        // we don't want to edit anything
                //    dataListView7.RemoveObject(e.RowObject); // remove object
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            totalUserDocs = 0;
            dtUsersAndDocs.Clear();
            progressBar1.Visible = true;
            label1.Text = String.Empty;
            label1.Visible = true;
            button9.Enabled = false;
            userDocExport.Enabled = false;
            //timer1.Start();
            backgroundWorker1.RunWorkerAsync();

            //Thread.Sleep(5000);
            //DataTable dt = GetUserDocuments();
            //DataTable dt1 = GetDocumentsMetadata(dt);
            //dataListView3.DataSource = dt1;



            //backgroundWorker1.CancelAsync();


        }

        private void btn_AddSource_Click(object sender, EventArgs e)
        {
            sourceAdded = true;
            DataRowView obj = dataListView4.SelectedObject as DataRowView;
            label2.Text = obj.Row["Name"].ToString();
            button11.Enabled = true;
        }

        private void dataListView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataListView4.SelectedObjects.Count == 1 && sourceAdded == false)
            {
                btn_AddSource.Enabled = true;
            }
            else
            {
                btn_AddSource.Enabled = false;
            }

            if (dataListView4.SelectedObjects.Count == 1 && sourceAdded == true)
            {
                btn_AddTarget.Enabled = true;
            }
            else
            {
                btn_AddTarget.Enabled = false;
            }

            //if (chb_AutoGetMetadata.Checked == true)
            //{
            GetDocMetadata();
            //}


        }

        private void GetDocMetadata()
        {
            DataTable dt = new DataTable();
            DataColumn dName = new DataColumn("Name");
            DataColumn dPath = new DataColumn("RelativePath");
            DataColumn dType = new DataColumn("Type");
            DataColumn isOrphan = new DataColumn("IsOrphan");
            DataColumn isSubfolder = new DataColumn("IsSubFolder");
            DataColumn dTaskCount = new DataColumn("TaskCount");
            DataColumn dId = new DataColumn("ID");
            DataColumn fId = new DataColumn("FolderID");
            dType.DataType = typeof(DocumentType);
            dType.DataType = typeof(DocumentType);
            dt.Columns.Add(dName);
            dt.Columns.Add(dPath);
            dt.Columns.Add(dType);
            dt.Columns.Add(isOrphan);
            dt.Columns.Add(isSubfolder);
            dt.Columns.Add(dTaskCount);
            dt.Columns.Add(fId);
            dt.Columns.Add(dId);

            DataTable dt1 = new DataTable();
            var obj1 = dataListView4.SelectedObjects;
            for (var a = 0; a < obj1.Count; a++)
            {

                DataRowView obj = obj1[a] as DataRowView;
                DataRow dataBoundItem = obj.Row;
                dt.ImportRow(dataBoundItem);


                dt1 = GetDocumentsMetadata(dt);
                dataListView5.DataSource = dt1;

                for (var i = 0; i < dataListView5.Columns.Count; i++)
                {
                    dataListView5.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }
            }

            if (sourceAdded == false)
            {
                //dtDocSource = dt;
                dtDocSource.FolderID = new Guid(dt1.Rows[0]["FolderID"].ToString());
                dtDocSource.ID = new Guid(dt1.Rows[0]["ID"].ToString());
                dtDocSource.Name = dt1.Rows[0]["Name"].ToString();
                dtDocSource.RelativePath = dt1.Rows[0]["RelativePath"].ToString();
                dtDocSource.IsOrphan = Convert.ToBoolean(dt1.Rows[0]["IsOrphan"].ToString());
                dtDocSource.IsSubFolder = Convert.ToBoolean(dt1.Rows[0]["IsSubFolder"].ToString());
                dtDocSource.TaskCount = Convert.ToInt32(dt1.Rows[0]["TaskCount"].ToString());
                dtDocSource.Type = DocumentType.User;
            }
            else if (targetAdded == false)
            {
                dtDocTarget.FolderID = new Guid(dt1.Rows[0]["FolderID"].ToString());
                dtDocTarget.ID = new Guid(dt1.Rows[0]["ID"].ToString());
                dtDocTarget.Name = dt1.Rows[0]["Name"].ToString();
                dtDocTarget.RelativePath = dt1.Rows[0]["RelativePath"].ToString();
                dtDocTarget.IsOrphan = Convert.ToBoolean(dt1.Rows[0]["IsOrphan"].ToString());
                dtDocTarget.IsSubFolder = Convert.ToBoolean(dt1.Rows[0]["IsSubFolder"].ToString());
                dtDocTarget.TaskCount = Convert.ToInt32(dt1.Rows[0]["TaskCount"].ToString());
                dtDocTarget.Type = DocumentType.User;
            }
        }


        private void btn_AddTarget_Click(object sender, EventArgs e)
        {
            try
            {
                DataRowView obj = dataListView4.SelectedObject as DataRowView;
                label4.Text = obj.Row["Name"].ToString();
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
                checkBox1.Enabled = true;
                button10.Enabled = true;
                targetAdded = true;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            dtDocSource = null;
            dtDocTarget = null;
            sourceAdded = false;
            //btn_AddTarget.Enabled = false;
            //btn_AddSource.Enabled = false;
            label2.Text = "";
            label4.Text = "";
            button10.Enabled = false;
            checkBox1.Enabled = false;
            btn_AddSource.Enabled = false;
            btn_AddTarget.Enabled = false;
            sourceAdded = false;
            targetAdded = false;
        }

        private void btn_GetDocMetadata_Click(object sender, EventArgs e)
        {
            //GetDocMetadata();
        }

        private void chb_AutoGetMetadata_CheckedChanged(object sender, EventArgs e)
        {
            dataListView5.Clear();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            var obj1 = dataListView4.SelectedObjects;
            for (var a = 0; a < obj1.Count; a++)
            {
                DataRowView obj = obj1[a] as DataRowView;
                DataRow dataBoundItem = obj.Row;

                DocumentNode dn = new DocumentNode();
                dn.FolderID = new Guid(dataBoundItem["FolderID"].ToString());
                dn.ID = new Guid(dataBoundItem["ID"].ToString());
                dn.Name = dataBoundItem["Name"].ToString();
                dn.RelativePath = dataBoundItem["RelativePath"].ToString();
                dn.IsOrphan = Convert.ToBoolean(dataBoundItem["IsOrphan"].ToString());
                dn.IsSubFolder = Convert.ToBoolean(dataBoundItem["IsSubFolder"].ToString());
                dn.TaskCount = Convert.ToInt32(dataBoundItem["TaskCount"].ToString());
                dn.Type = DocumentType.User;

                var meta = Client.GetDocumentMetaData(dn, DocumentMetaDataScope.Authorization);
                foreach (object o in dataListView7.Objects)
                {
                    var b1 = o as DataRowView;
                    DataRow dataBoundItem1 = b1.Row;
                    DocumentAccessEntry dae = new DocumentAccessEntry();
                    dae.UserName = dataBoundItem1["Id"].ToString();
                    dae.DayOfWeekConstraints = new List<DayOfWeek>();

                    meta.Authorization.Access.Add(dae);
                }
                Client.SaveDocumentMetaData(meta);
            }

            dtSelectedUSers.Clear();
            GetDocMetadata();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            var obj1 = dataListView4.SelectedObjects;
            for (var a = 0; a < obj1.Count; a++)
            {
                DataRowView obj = obj1[a] as DataRowView;
                DataRow dataBoundItem = obj.Row;

                DocumentNode dn = new DocumentNode();
                dn.FolderID = new Guid(dataBoundItem["FolderID"].ToString());
                dn.ID = new Guid(dataBoundItem["ID"].ToString());
                dn.Name = dataBoundItem["Name"].ToString();
                dn.RelativePath = dataBoundItem["RelativePath"].ToString();
                dn.IsOrphan = Convert.ToBoolean(dataBoundItem["IsOrphan"].ToString());
                dn.IsSubFolder = Convert.ToBoolean(dataBoundItem["IsSubFolder"].ToString());
                dn.TaskCount = Convert.ToInt32(dataBoundItem["TaskCount"].ToString());
                dn.Type = DocumentType.User;

                var meta = Client.GetDocumentMetaData(dn, DocumentMetaDataScope.Authorization);
                foreach (object o in dataListView7.Objects)
                {
                    var b1 = o as DataRowView;
                    DataRow dataBoundItem1 = b1.Row;

                    for (var m = 0; m < meta.Authorization.Access.Count; m++)
                    {
                        if (meta.Authorization.Access[m].UserName == dataBoundItem1["Id"].ToString())
                        {
                            meta.Authorization.Access.RemoveAt(m);
                        }
                    }


                    //var b1 = o as DataRowView;
                    //DataRow dataBoundItem1 = b1.Row;
                    //DocumentAccessEntry dae = new DocumentAccessEntry();
                    //dae.UserName = dataBoundItem1["Id"].ToString();
                    //dae.DayOfWeekConstraints = new List<DayOfWeek>();
                    //dae.AccessMode = DocumentAccessEntryMode.Restricted;
                    //dae.IsAnonymous = false;
                    ////dae.UserName
                    //meta.Authorization.Access.RemoveAt(0);
                }
                Client.SaveDocumentMetaData(meta);
            }

            dtSelectedUSers.Clear();
            GetDocMetadata();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

            List<DocumentAccessEntry> forRemove = new List<DocumentAccessEntry>();
            List<string> toAdd = new List<string>();

            var source = Client.GetDocumentMetaData(dtDocSource, DocumentMetaDataScope.Authorization);
            var target = Client.GetDocumentMetaData(dtDocTarget, DocumentMetaDataScope.Authorization);

            if (target.Authorization.Access.Count == 0)
            {
                for (var s = 0; s < source.Authorization.Access.Count; s++)
                {
                    DocumentAccessEntry dae = new DocumentAccessEntry();
                    dae.UserName = source.Authorization.Access[s].UserName.ToString();
                    dae.DayOfWeekConstraints = new List<DayOfWeek>();
                    target.Authorization.Access.Add(dae);
                }
            }
            else if (checkBox1.Checked == false)
            {

                if (source.Authorization.Access.Count == 0)
                {
                    target.Authorization.Access.Clear();
                }
                else
                {

                    for (var t = 0; t < target.Authorization.Access.Count; t++)
                    {
                        bool exists = false;
                        for (var s = 0; s < source.Authorization.Access.Count; s++)
                        {
                            if (target.Authorization.Access[t].UserName == source.Authorization.Access[s].UserName)
                            {
                                exists = true;
                            }
                        }

                        if (exists == false)
                        {
                            forRemove.Add(target.Authorization.Access[t]);
                        }

                        for (var r = 0; r < forRemove.Count; r++)
                        {
                            target.Authorization.Access.Remove(forRemove[r]);
                        }

                    }

                    for (var s = 0; s < source.Authorization.Access.Count; s++)
                    {
                        bool exists = false;
                        for (var t = 0; t < target.Authorization.Access.Count; t++)
                        {
                            if (source.Authorization.Access[s].UserName == target.Authorization.Access[t].UserName)
                            {
                                exists = true;
                            }
                        }

                        if (exists == false)
                        {
                            toAdd.Add(source.Authorization.Access[s].UserName.ToString());
                        }
                    }

                    for (var a = 0; a < toAdd.Count; a++)
                    {
                        DocumentAccessEntry dae = new DocumentAccessEntry();
                        dae.UserName = toAdd[a].ToString();
                        dae.DayOfWeekConstraints = new List<DayOfWeek>();
                        target.Authorization.Access.Add(dae);
                    }
                }
            }
            else
            {
                for (var s = 0; s < source.Authorization.Access.Count; s++)
                {
                    bool exists = false;
                    for (var t = 0; t < target.Authorization.Access.Count; t++)
                    {
                        if (source.Authorization.Access[s].UserName == target.Authorization.Access[t].UserName)
                        {
                            exists = true;
                        }
                    }

                    if (exists == false)
                    {
                        toAdd.Add(source.Authorization.Access[s].UserName.ToString());
                    }
                }

                for (var a = 0; a < toAdd.Count; a++)
                {
                    DocumentAccessEntry dae = new DocumentAccessEntry();
                    dae.UserName = toAdd[a].ToString();
                    dae.DayOfWeekConstraints = new List<DayOfWeek>();
                    target.Authorization.Access.Add(dae);
                }
            }

            Client.SaveDocumentMetaData(target);
            //btn_AddSource.Enabled = false;
            btn_AddTarget.Enabled = false;
            label2.Text = "";
            label4.Text = "";
            //label2.Enabled = false;
            //label4.Enabled = false;
            checkBox1.Enabled = false;
            button10.Enabled = false;
            button11.Enabled = false;
            sourceAdded = false;
            targetAdded = false;

            GetDocMetadata();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
            config.AppSettings.Settings["qms"].Value = textBox3.Text;
            config.Save(ConfigurationSaveMode.Modified);
            qms = textBox3.Text;
            label3.Text = "";
            server = qms.Substring(qms.IndexOf("//") + 2, qms.Length - qms.IndexOf("/QMS") - 3);
        }

        private void userDocExport_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            saveFileDialog1.FileName = server + "_docs-and-users_" + dt.ToString("yyyyMMdd-HHmmss") + ".csv";
            //saveFileDialog1.ShowDialog();

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (dataListView3.Columns.Count > 0)
                {
                    OLVExporter ex = new OLVExporter(dataListView3, dataListView3.FilteredObjects);
                    string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
                    string fileName = saveFileDialog1.FileName;
                    System.IO.File.WriteAllText(fileName, test);
                }
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            dtSelectedUSers.Clear();
        }

        private void dataListView6_DoubleClick(object sender, EventArgs e)
        {
            AddUserToList();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                DataTable dt = GetUserDocuments();
                dtUsersAndDocs = GetDocumentsMetadata(dt);
                backgroundWorker1.ReportProgress(100, "done");
            }
            catch (System.Exception ex)
            {
                backgroundWorker1.ReportProgress(100, ex.Message);
            }
            finally
            {

            }

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Visible = false;
            label1.Visible = false;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {


            if (e.UserState.ToString() == "done")
            {
                //timer1.Stop();
                button9.Enabled = false;
                userDocExport.Enabled = false;

                dataListView3.DataSource = dtUsersAndDocs;
                for (var i = 0; i < dataListView3.Columns.Count; i++)
                {
                    dataListView3.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }

                button9.Enabled = true;
                userDocExport.Enabled = true;
            }

            if (e.UserState.ToString() == "error")
            {
                //timer1.Stop();
                button9.Enabled = true;
                userDocExport.Enabled = true;
                progressBar1.Visible = false;
                label1.Visible = false;
            }

            if (e.UserState.ToString() == "totalDocs")
            {
                progressBar1.Maximum = e.ProgressPercentage;
                totalUserDocs = e.ProgressPercentage;
                label1.Text = "0 / " + e.ProgressPercentage;
            }

            if (e.UserState.ToString() == "processeddocs")
            {
                progressBar1.Increment(1);
                label1.Text = progressBar1.Value + " / " + totalUserDocs;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {

                notifyIcon1.ShowBalloonTip(500);
                this.Hide();
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void restoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            AboutBox1 box = new AboutBox1();
            box.ShowDialog();

            /*PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, "systems\\GRP-Qlikview-GrBI-Mob-Users-G");
            var users = group.GetMembers();

            if (group != null)
            {
                // iterate over members
                foreach (Principal p in group.GetMembers())
                {
                    //Console.WriteLine("{0}: {1}", p.StructuralObjectClass, p.DisplayName);

                    // do whatever you need to do to those members
                    UserPrincipal theUser = p as UserPrincipal;

                    if (theUser != null)
                    {
                        if (theUser.IsAccountLockedOut())
                        {

                        }
                        else
                        {

                        }
                    }
                }
            }*/
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedTab = tabControl1.SelectedIndex;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserSearch();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                FilterDocs();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            AboutBox1 box = new AboutBox1();
            box.ShowDialog();
        }

        private void rbtn_path_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rbtn_doc_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = server + "_active-docs.csv";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = saveFileDialog1.FileName;
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                config.AppSettings.Settings["activedocspath"].Value = textBox4.Text;
                config.Save(ConfigurationSaveMode.Modified);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            timer1.Interval = (Int32)(numericUpDown1.Value * 1000);

            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
            config.AppSettings.Settings["activedocsinterval"].Value = numericUpDown1.Value.ToString();
            config.Save(ConfigurationSaveMode.Modified);
            //button16.Enabled = false;
            if (button16.Text == "Start")
            {
                timer1.Start();
                button16.Text = "Stop";
                numericUpDown1.Enabled = false;
                button17.Enabled = false;
                //button16.Enabled = false;
                textBox4.Enabled = false;
            }
            else
            {
                timer1.Stop();
                button16.Text = "Start";
                numericUpDown1.Enabled = true;
                //button16.Enabled = true;
                button17.Enabled = true;
                textBox4.Enabled = true;
            }
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            GetActiveUsers();
            //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            if (dataListView2.Columns.Count > 0)
            {
                OLVExporter ex = new OLVExporter(dataListView2, dataListView2.FilteredObjects);
                ex.IncludeColumnHeaders = false;
                string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
                string fileName = textBox4.Text;
                System.IO.File.AppendAllText(fileName, test);
            }
            //}
        }

        private void button18_Click(object sender, EventArgs e)
        {
            var t1 = new QlikviewEnhancedUserControl.ServiceReference.TaskStatusFilter();

            var t = Client.GetTaskStatuses(t1, TaskStatusScope.Extended);
            //var t = Client.GetTasks(qdsId);
            //var t = Client.GetTaskStatus(new Guid("e42adf3d-22d7-4a42-bcba-06f04bbc8959"), TaskStatusScope.All);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            saveFileDialog1.FileName = server + "_docs-and-users_" + dt.ToString("yyyyMMdd-HHmmss") + ".csv";
            //saveFileDialog1.ShowDialog();

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (dataListView5.Columns.Count > 0)
                {
                    OLVExporter ex = new OLVExporter(dataListView5, dataListView5.FilteredObjects);
                    string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
                    string fileName = saveFileDialog1.FileName;
                    System.IO.File.WriteAllText(fileName, test);
                }
            }
        }

        private void btnBatchImport_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            DataTable batchFileData = ConvertCSVtoDataTable(openFileDialog1.FileName);

            DataView view = new DataView(batchFileData);
            DataTable distinctValues = view.ToTable(true, "Document", "Area");

            string stats = "";

            for (var a = 0; a < distinctValues.Rows.Count; a++)
            {
                string query = "Document = '" + distinctValues.Rows[a][0] + "' and Area = '" + distinctValues.Rows[a][1] + "'";
                DataRow[] results = batchFileData.Select(query);


                for (var i = 0; i < dataListView4.GetItemCount(); i++)
                {
                    var obj = dataListView4.Items[i];
                    var items = obj.SubItems;
                    //var t = items[1];
                    //var a1 = items[0].Text.ToString().ToLower();
                    //var a2 = items[1].Text.ToString().ToLower();
                    //var a3 = distinctValues.Rows[a][0].ToString().ToLower();
                    //var a4 = distinctValues.Rows[a][1].ToString().ToLower();

                    if (items[0].Text.ToString().ToLower() == distinctValues.Rows[a][0].ToString().ToLower() + ".qvw" && items[1].Text.ToString().ToLower() == distinctValues.Rows[a][1].ToString().ToLower())
                    {
                        //MessageBox.Show(items[6].Text.ToString());
                        DocumentNode dn = new DocumentNode();
                        dn.FolderID = new Guid(items[6].Text.ToString());
                        dn.ID = new Guid(items[7].Text.ToString());
                        dn.Name = items[0].Text.ToString();
                        dn.RelativePath = items[1].Text.ToString();
                        dn.IsOrphan = Convert.ToBoolean(items[3].Text.ToString());
                        dn.IsSubFolder = Convert.ToBoolean(items[4].Text.ToString());
                        dn.TaskCount = Convert.ToInt32(items[5].Text.ToString());
                        dn.Type = DocumentType.User;

                        var meta = Client.GetDocumentMetaData(dn, DocumentMetaDataScope.Authorization);

                        for (var b = 0; b < results.Length; b++)
                        {
                            DocumentAccessEntry dae = new DocumentAccessEntry();
                            dae.UserName = results[b][0].ToString();
                            dae.DayOfWeekConstraints = new List<DayOfWeek>();

                            meta.Authorization.Access.Add(dae);
                        }
                        Client.SaveDocumentMetaData(meta);

                        stats = stats + "Document: " + items[0].Text.ToString() + "; Added users: " + results.Length.ToString() + System.Environment.NewLine;
                    }

                }
            }

            MessageBox.Show(stats);


        }

        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            string[] headers = sr.ReadLine().Split(',');
            DataTable dt = new DataTable();

            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }

            while (!sr.EndOfStream)
            {
                string[] rows = sr.ReadLine().Split(',');
                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }

            sr.Close();
            return dt;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = folderBrowserDialog2.SelectedPath;
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                config.AppSettings.Settings["usersdocpath"].Value = textBox5.Text;
                config.Save(ConfigurationSaveMode.Modified);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            timer2.Interval = (Int32)(numericUpDown2.Value * 1000);

            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
            config.AppSettings.Settings["usersdocinterval"].Value = numericUpDown2.Value.ToString();
            config.Save(ConfigurationSaveMode.Modified);

            if (button21.Text == "Start")
            {
                timer2.Start();
                button21.Text = "Stop";
                numericUpDown2.Enabled = false;
                button20.Enabled = false;
                textBox5.Enabled = false;
                ExportUserDocs();
            }
            else
            {
                timer2.Stop();
                button21.Text = "Start";
                numericUpDown2.Enabled = true;
                button20.Enabled = true;
                textBox5.Enabled = true;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            ExportUserDocs();
        }

        private void ExportUserDocs()
        {
            dtUsersAndDocs.Clear();
            DataTable dt = GetUserDocuments();
            dtUsersAndDocs = GetDocumentsMetadata(dt);
            dataListView3.DataSource = dtUsersAndDocs;

            DateTime dtime = DateTime.Now;
            string fileName = textBox5.Text + "\\" + server + "_docs-and-users_" + dtime.ToString("yyyyMMdd-HHmmss") + ".csv";

            OLVExporter ex = new OLVExporter(dataListView3, dataListView3.FilteredObjects);
            string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
            System.IO.File.WriteAllText(fileName, test);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            DataTable dtTasks = new DataTable();
            dtTasks.Columns.Add("Name");
            dtTasks.Columns.Add("DocName");
            dtTasks.Columns.Add("Category");
            dtTasks.Columns.Add("RelativePath");
            dtTasks.Columns.Add("TaskId");
            dtTasks.Columns.Add("DocId");
            

            var remoteTasks = Client.RemoteGetTasks(remoteQmsId, remoteQdsId);

            foreach (TaskInfo remoteTask in remoteTasks)
            //for (var i = 0; i < 10; i++ )
            {
                //TaskInfo remoteTask = remoteTasks[i];
                var t = Client.RemoteGetDocumentTask(remoteQmsId, remoteTask.ID, DocumentTaskScope.All);

                if (t != null)
                {
                    DataRow row = dtTasks.NewRow();
                    row["Name"] = remoteTask.Name.ToString();
                    row["DocName"] = t.Document.Name.ToString();
                    row["Category"] = t.DocumentInfo.Category.ToString();
                    row["RelativePath"] = t.Document.RelativePath.ToString();
                    row["TaskId"] = remoteTask.ID.ToString();
                    row["DocId"] = t.Document.ID;
                    dtTasks.Rows.Add(row);
                }
            }

            dataListView8.DataSource = dtTasks;

            for (var i = 0; i < dataListView8.Columns.Count; i++)
            {
                dataListView8.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }

        }

        private void dataListView8_DoubleClick(object sender, EventArgs e)
        {

        }

        private void button23_Click(object sender, EventArgs e)
        {
            

            foreach (object o in dataListView9.Objects)
            {
                var b1 = o as DataRowView;
                DataRow dataBoundItem1 = b1.Row;

                var a = Client.GetSourceDocumentNodes(qdsId, new Guid(dataBoundItem1["DocNodeFolderId"].ToString()), dataBoundItem1["DocNodeRelativePath"].ToString());

                for (var i = 0; i < a.Count; i++)
                {
                    if (a[i].IsSubFolder == false)
                    {
                        if(a[i].ID.ToString() == dataBoundItem1["DocId"].ToString()) {
                            Client.ImportDocumentTask(remoteQmsId, new Guid(dataBoundItem1["TaskId"].ToString()), qdsId, a[i]);
                        }
                    }
                }

                    
                //DocumentAccessEntry dae = new DocumentAccessEntry();
                //dae.UserName = dataBoundItem1["Id"].ToString();
                //dae.DayOfWeekConstraints = new List<DayOfWeek>();

                
            }


            var t = dataListView10.Objects;
            DataTable r = t as DataTable;
            dtTasks.Clear();
            
            //var a = Client.GetSourceDocumentNodes(qdsId, new Guid(""), "");
            //Client.ImportDocumentTask(remoteQmsId, new Guid(""), qdsId, );
        }

        private void button24_Click(object sender, EventArgs e)
        {
            foreach (DataRowView row in dataListView8.SelectedObjects)
            {
                bool present = false;

                for (int i = 0; i < dtTasks.Rows.Count; i++)
                {
                    if (dtTasks.Rows[i]["TaskId"].ToString() == row.Row["TaskId"].ToString())
                    {
                        present = true;
                        break;
                    }
                }

                if (present == false)
                {
                    DataRow dr = dtTasks.NewRow();
                    dr["Name"] = row.Row["Name"];
                    //dr["DocName"] = row.Row["DocName"];
                    //dr["Category"] = row.Row["Category"];
                    //dr["RelativePath"] = row.Row["RelativePath"];
                    dr["TaskId"] = row.Row["TaskId"];                    
                    //dr["DocId"] = row.Row["DocId"];
                    var t = dataListView10.SelectedObjects;
                    DataRowView r = t[0] as DataRowView;
                    dr["DocNodeName"] = r["Name"];
                    dr["DocId"] = r["ID"];
                    dr["DocNodeFolderId"] = r["FolderID"];
                    dr["DocNodeRelativePath"] = r["RelativePath"];
                    //Client.GetD
                    dtTasks.Rows.Add(dr);

                    var a = Client.GetTasksForDocument(new Guid(r["ID"].ToString()));
                    //a[0].
                }
            }

            dataListView9.DataSource = dtTasks;
            for (var i = 0; i < dataListView9.Columns.Count; i++)
            {
                dataListView9.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);

            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            foreach (DataRowView row in dataListView9.SelectedObjects)
            {
                for (var i = 0; i < dtTasks.Rows.Count; i++)
                {
                    DataRow dr = dtTasks.Rows[i];
                    if (dr["TaskId"].ToString() == row.Row["TaskId"].ToString())
                    {
                        dtTasks.Rows.Remove(dr);
                    }
                }
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            DataTable dtDocs = new DataTable();
            DataColumn dName = new DataColumn("Name");
            DataColumn dPath = new DataColumn("RelativePath");
            DataColumn dType = new DataColumn("Type");
            DataColumn isOrphan = new DataColumn("IsOrphan");
            DataColumn isSubfolder = new DataColumn("IsSubFolder");
            DataColumn dTaskCount = new DataColumn("TaskCount");
            DataColumn dId = new DataColumn("ID");
            DataColumn fId = new DataColumn("FolderID");
            dType.DataType = typeof(DocumentType);
            dType.DataType = typeof(DocumentType);
            dtDocs.Columns.Add(dName);
            dtDocs.Columns.Add(dPath);
            dtDocs.Columns.Add(dType);
            dtDocs.Columns.Add(isOrphan);
            dtDocs.Columns.Add(isSubfolder);
            dtDocs.Columns.Add(dTaskCount);
            dtDocs.Columns.Add(fId);
            dtDocs.Columns.Add(dId);

            var sourceDocs = new List<DocumentNode>();
            try
            {
                //Client.ClearQVSCache(QVSCacheObjects.UserDocumentList);
                sourceDocs = Client.GetSourceDocuments(qdsId);
            }
            catch (System.Exception ex)
            {
                key = Client.GetTimeLimitedServiceKey();
                ServiceKeyClientMessageInspector.ServiceKey = key;
                //Client.ClearQVSCache(QVSCacheObjects.UserDocumentList);
                sourceDocs = Client.GetSourceDocuments(qdsId);
            }

            foreach (var sourceDoc in sourceDocs)
            {
                var a = sourceDoc;
                DataRow dr = dtDocs.NewRow();
                dr["FolderID"] = a.FolderID;
                dr["ID"] = a.ID;
                dr["IsOrphan"] = a.IsOrphan;
                dr["IsSubFolder"] = a.IsSubFolder;
                dr["Name"] = a.Name;
                dr["RelativePath"] = a.RelativePath;
                dr["TaskCount"] = a.TaskCount;
                dr["Type"] = a.Type;
                dtDocs.Rows.Add(dr);
            }

            dataListView10.DataSource = dtDocs;

        }
    }
}
