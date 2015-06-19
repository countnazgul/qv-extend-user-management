using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using QV_UserManagement.ServiceReference;
using QV_UserManagement.ServiceSupport;
using BrightIdeasSoftware;
using System.Threading;


namespace QV_UserManagement
{
    public partial class Form1 : Form
    {
        QMSClient Client;
        List<string> QVS = new List<string>();
        DataTable dtSelectedUSers = new DataTable();
        DataTable dtUserDocuments = new DataTable();
        bool sourceAdded = false;
        Guid qvsId = new Guid("00000000-0000-0000-0000-000000000000");
        Guid dscId = new Guid("00000000-0000-0000-0000-000000000000");
        bool loadingVisible = false;
        string key = "";

        public Form1()
        {
            InitializeComponent();

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

            
            Client = new QMSClient("BasicHttpBinding_IQMS");
            key = Client.GetTimeLimitedServiceKey();
            ServiceKeyClientMessageInspector.ServiceKey = key;            
           
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

                //DataTable dtStats = new DataTable();
   

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
                var t = Client.GetQVSDocumentsAndUsers(a, QueryTarget.ClusterMember);

                foreach (KeyValuePair<string, List<string>> entry in t)
                {
                    var doc = entry.Key;
                    for (int b = 0; b < entry.Value.Count; b++)
                    {
                        var executiondate = DateTime.Now.ToString("yyyyMMddHHmmss");

                        DataRow dr = dt.NewRow();
                        dr["timestamp"] = executiondate.ToString();
                        dr["server"] =  members[m].Substring(members[m].IndexOf('_') + 1, members[m].Length - members[m].IndexOf('_') - 1).Trim();
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
                userDocs = Client.GetUserDocuments(qvsId);
            }
            catch (System.Exception ex)
            {
                key = Client.GetTimeLimitedServiceKey();
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

            for (var i = 0; i < dtDocs.Rows.Count; i++)
            {

                var documentNode = new DocumentNode();
                documentNode.FolderID = new Guid(dtDocs.Rows[i]["FolderID"].ToString());
                documentNode.ID = new Guid(dtDocs.Rows[i]["ID"].ToString());
                documentNode.Name = dtDocs.Rows[i]["Name"].ToString();
                documentNode.IsOrphan = Convert.ToBoolean(dtDocs.Rows[i]["IsOrphan"]);
                documentNode.IsSubFolder = Convert.ToBoolean(dtDocs.Rows[i]["IsSubFolder"]);
                documentNode.RelativePath = dtDocs.Rows[i]["RelativePath"].ToString();
                documentNode.TaskCount = Convert.ToInt32(dtDocs.Rows[i]["TaskCount"]);
                documentNode.Type = DocumentType.User;

                var t = Client.GetDocumentMetaData(documentNode, DocumentMetaDataScope.Authorization);

                if (t.Authorization.Access.Count > 0)
                {
                    for (var b = 0; b < t.Authorization.Access.Count; b++)
                    {
                        DataRow dr = dt.NewRow();
                        dr["FolderID"] = dtDocs.Rows[i]["FolderID"];
                        dr["ID"] = dtDocs.Rows[i]["ID"];
                        dr["IsOrphan"] = dtDocs.Rows[i]["IsOrphan"];
                        dr["IsSubFolder"] = dtDocs.Rows[i]["IsSubFolder"];
                        dr["Name"] = dtDocs.Rows[i]["Name"];
                        dr["RelativePath"] = dtDocs.Rows[i]["RelativePath"];
                        dr["TaskCount"] = dtDocs.Rows[i]["TaskCount"];
                        dr["Type"] = dtDocs.Rows[i]["Type"];
                        dr["User"] = t.Authorization.Access[b].UserName;
                        dt.Rows.Add(dr);
                    }
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    dr["FolderID"] = dtDocs.Rows[i]["FolderID"];
                    dr["ID"] = dtDocs.Rows[i]["ID"];
                    dr["IsOrphan"] = dtDocs.Rows[i]["IsOrphan"];
                    dr["IsSubFolder"] = dtDocs.Rows[i]["IsSubFolder"];
                    dr["Name"] = dtDocs.Rows[i]["Name"];
                    dr["RelativePath"] = dtDocs.Rows[i]["RelativePath"];
                    dr["TaskCount"] = dtDocs.Rows[i]["TaskCount"];
                    dr["Type"] = dtDocs.Rows[i]["Type"];
                    dr["User"] = "NONE";
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = GetActiveUserAndDocs(QVS);
            dataListView2.DataSource = dt;

            for (var i = 0; i < dataListView2.Columns.Count; i++)
            {
                dataListView2.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = GetServicesStatuses();
            dataListView1.DataSource = dt;

            for (var i = 0; i < dataListView1.Columns.Count; i++)
            {
                dataListView1.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Start();
            try
            {
                DataTable dt = GetUserDocuments();
                dataListView4.DataSource = dt;

                for (var i = 0; i < dataListView4.Columns.Count; i++)
                {
                    dataListView4.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }
                //DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                //dataGridView3.Columns.Add(chk);
                //chk.HeaderText = " ";
                //chk.Name = "chk";
                //dataGridView3.Columns[dataGridView3.Columns.Count-1].DisplayIndex = 0;
            }
            catch (System.Exception ex)
            {
                timer1.Stop();
            }
            finally
            {
                timer1.Stop();
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
            OLVExporter ex = new OLVExporter(dataListView4, dataListView4.FilteredObjects);
            string test = ex.ExportTo(OLVExporter.ExportFormat.CSV);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //var dscId = new Guid("295f4414-8d2e-44a0-8b7b-c91c25b6da66");
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

                        //label1.Text = user[0].Name + " ," + user[0].OtherProperty;
                }
            }

            




            /*else
            {
                DataRow dr = dt.NewRow();
                dr["Id"] = "Not found";
                dr["Name"] = "Not found";
                dt.Rows.Add(dr);
            }*/

            dataListView6.DataSource = dt;
            for (var i = 0; i < dataListView6.Columns.Count; i++)
            {
                dataListView6.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        

        private void button6_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataColumn dcId = new DataColumn("Id");
            DataColumn dcName = new DataColumn("Name");
            dt.Columns.Add(dcId);
            dt.Columns.Add(dcName);

            foreach (DataRowView row in dataListView6.SelectedObjects)
            {
                DataRow dr = dt.NewRow();
                dr["Id"] = row.Row["ID"];
                dr["Name"] = row.Row["Name"];
                dt.Rows.Add(dr);
            }

            dataListView7.DataSource = dt;
            for (var i = 0; i < dataListView7.Columns.Count; i++)
            {
                dataListView7.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }

        }

        private void button7_Click(object sender, EventArgs e)
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
            
            //(dataGridView3.DataSource as DataTable).DefaultView.RowFilter = string.Concat(field, " LIKE '%", textBox2.Text, "%'"); //"RelativePath LIKE '%[,]" + textBox2.Text + "[,]%'";            
            (dataListView4.DataSource as DataTable).DefaultView.RowFilter = string.Concat(field, " LIKE '%", textBox2.Text, "%'"); //"RelativePath LIKE '%[,]" + textBox2.Text + "[,]%'";            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            
            //(dataGridView3.DataSource as DataTable).DefaultView.RowFilter = string.Empty;
            (dataListView4.DataSource as DataTable).DefaultView.RowFilter = string.Empty;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            rbtn_doc.Checked = true;
            btn_AddSource.Enabled = false;
            btn_AddTarget.Enabled = false;
            chb_AutoGetMetadata.Checked = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable dt = GetUserDocuments();
            DataTable dt1 = GetDocumentsMetadata(dt);
            dataListView3.DataSource = dt1;

            for (var i = 0; i < dataListView3.Columns.Count; i++)
            {
                dataListView3.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void btn_AddSource_Click(object sender, EventArgs e)
        {
            sourceAdded = true;
            DataRowView obj = dataListView4.SelectedObject as DataRowView;
            label2.Text = obj.Row["Name"].ToString();
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

            if (chb_AutoGetMetadata.Checked == true)
            {
                GetDocMetadata();
            }


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


            var obj1 = dataListView4.SelectedObjects;
            for (var a = 0; a < obj1.Count; a++)
            {

                DataRowView obj = obj1[a] as DataRowView;
                DataRow dataBoundItem = obj.Row;
                dt.ImportRow(dataBoundItem);


                DataTable dt1 = GetDocumentsMetadata(dt);
                dataListView5.DataSource = dt1;

                for (var i = 0; i < dataListView5.Columns.Count; i++)
                {
                    dataListView5.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
                }
            }


        }


        private void btn_AddTarget_Click(object sender, EventArgs e)
        {
            DataRowView obj = dataListView4.SelectedObject as DataRowView;
            label4.Text = obj.Row["Name"].ToString();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            sourceAdded = false;
            btn_AddTarget.Enabled = false;
            btn_AddSource.Enabled = false;
            label2.Text = "";
            label4.Text = "";
        }

        private void btn_GetDocMetadata_Click(object sender, EventArgs e)
        {
            GetDocMetadata();
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
                                var b1 = o as DataRowView ;
                                DataRow dataBoundItem1 = b1.Row;
                                DocumentAccessEntry dae = new DocumentAccessEntry();
                                dae.UserName = dataBoundItem1["Id"].ToString();
                                dae.DayOfWeekConstraints = new List<DayOfWeek>();

                                meta.Authorization.Access.Add(dae);                                
                            }
                Client.SaveDocumentMetaData(meta);           
            }
            
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
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            //label1.Text = "Loading ...";
            //label1.ForeColor = Color.Red;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (loadingVisible == false)
            {
                label1.Visible = true;
                loadingVisible = true;
            }
            else
            {
                label1.Visible = false;
                loadingVisible = false;
            }
        }

    }
}
