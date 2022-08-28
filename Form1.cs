using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Runtime.InteropServices.ComTypes;
using Word = Microsoft.Office.Interop.Word;

using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;


namespace hackathon
{
    public partial class Asistent : Form
    {
        string filePath = string.Empty;

        IFirebaseConfig config = new FirebaseConfig
        {
            AuthSecret= "i9bm6AYdelKKJsBgBsZsVZbUi9fnya85li1DbeLH",
            BasePath= "https://hackathon-b45a8-default-rtdb.europe-west1.firebasedatabase.app/"
        };

        IFirebaseClient client;

        public Asistent()
        {
            InitializeComponent();
        }

        private void Asistent_Load(object sender, EventArgs e)
        {
            client = new FireSharp.FirebaseClient(config);
        }

        private void Choice_File_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;

                var fileStream = openFileDialog.OpenFile();

                using (StreamReader reader = new StreamReader(fileStream))
                {
                    Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();
                    object File = filePath;
                    object nullobject = System.Reflection.Missing.Value; Microsoft.Office.Interop.Word.Application wordobject = new Microsoft.Office.Interop.Word.Application();
                    wordobject.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; Microsoft.Office.Interop.Word._Document docs = wordObject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject); docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();
                    this.richTextBox1.Paste();
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                }
            }
        }

        private async void Processing_Click(object sender, EventArgs e)
        {
            var data = new Data
            {
                Doc = richTextBox1.Text
            };

            SetResponse response = await client.SetAsync("Documents/"+richTextBox1.Text, data);
            Data result = response.ResultAs<Data>();

            MessageBox.Show("Data inserted");
        }
    }
}
