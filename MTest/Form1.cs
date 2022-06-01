using System;
using System.Windows.Forms;
namespace MTest {
    public partial class Form1 : Form {
        FileConvertLogic fcl = FileConvertLogic.GetInstance();
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            ;
            var sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files | *.xlsx";
            sfd.DefaultExt = "xlsx";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                var path = sfd.FileName;
                fcl.CreateFile(path);
                MessageBox.Show("Zakończono tworzenie pliku.");

            }

        }

        private void button2_Click(object sender, EventArgs e) {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Xml Files | *.xml";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                var path = ofd.FileName;

                fcl.OpenFile(path);
                textBox1.Text = path;


            }


        }
    }
}
