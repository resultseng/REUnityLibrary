using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using REUnityLibrary;
using Hyland.Unity;


namespace UserForms
{
    public partial class BillsTestForm1 : Form
    {
        private Hyland.Unity.Application _app;
        private Hyland.Unity.Document _doc;
        private REUnityLibrary.DocumentKeywordLibrary _kw;

        public BillsTestForm1()
        {
            InitializeComponent();

            Login();

            _doc = _app.Core.GetDocumentByID(12800);
            //textBox1.Text = _doc.Name;

            _kw = new DocumentKeywordLibrary(_app, _doc);
        }


        private bool Login()
        {
            bool result = false;
            try
            {
                string appServer = "http://10.1.5.15/9sf_appserver/service.asmx";

                AuthenticationProperties authProps = Hyland.Unity.Application.CreateOnBaseAuthenticationProperties(appServer,
                        "MANAGER", "PASSWORD", "9secondfoods");

                _app = Hyland.Unity.Application.Connect(authProps);
                if (_app != null)
                {
                    _app.Diagnostics.Write("Invoice service logged in");
                    lblSessionID.Text = _app.SessionID.ToString();
                    result = true;
                }

            }
            catch (UnityAPIException unityEx)
            {
                MessageBox.Show("connection error: " + unityEx.Message);
            }

            return result;
        }

        private void Disconnect()
        {
            try
            {
                if (_app != null)
                {
                    _app.Diagnostics.Write("Invoice service disconnecting");

                    if (_app.IsConnected)
                    {
                        _app.Disconnect();
                        _app = null;
                    }
                }

            }
            catch (UnityAPIException unityEx)
            {
                MessageBox.Show("Disconnect error: " + unityEx.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //listBox1.Items.Add("fake = " + _kw.DocumentTypeContainsKeywordType("fake").ToString());
            //listBox1.Items.Add("Invoice Amount = " + _kw.DocumentTypeContainsKeywordType("Invoice Amount").ToString());

            foreach (ReadOnlyKeyword item in _kw.AllKeywordsList)
            {
                listBox1.Items.Add(item.Name + " - " + item.Value + " - " + item.RecordType.ToString());
            }

            listBox1.Items.Add((_kw.AllKeywordsList.Contains(new ReadOnlyKeyword { Name = "no exist kw" })).ToString());
            listBox1.Items.Add((_kw.AllKeywordsList.Contains(new ReadOnlyKeyword { Name = "Vendor Name" })).ToString());
            listBox1.Items.Add((_kw.AllKeywordsList.Contains(new ReadOnlyKeyword { Name = "Vendor Name", Value = "COMPUTERS ARE US" })).ToString());
            listBox1.Items.Add((_kw.AllKeywordsList.Contains(new ReadOnlyKeyword { Name = "Vendor Name", Value = "COMPUTERS ARE US", RecordType = Hyland.Unity.RecordType.StandAlone })).ToString());

            listBox1.Items.Add((_kw.AllKeywordsList.Exists(x => x.Name == "Vendor Name")));
            listBox1.Items.Add((_kw.AllKeywordsList.Exists(x => x.Name == "Vendor XName")));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            REUnityLibrary.Utility util1 = new REUnityLibrary.Utility(_app);
            REUnityLibrary.Utility util2 = new REUnityLibrary.Utility(_app, "no item");
            REUnityLibrary.Utility util3 = new REUnityLibrary.Utility(_app, "DiagnosticLevel");

            util1.SetDiagnosticLevel();
            listBox1.Items.Add(_app.Diagnostics.Level.ToString());
        }


    }
}
