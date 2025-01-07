using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Linealizar;

namespace LinealizarHojasExel
{
    public partial class SecondForm : Form
    {
        string _fullfileInpath;
        FolderBrowserDialog dirVir;
        Form1 _main;
        string directoryOutPath = null;

        public SecondForm(Form1 main, string fullfileInpath)
        {
            InitializeComponent();
            _fullfileInpath = fullfileInpath;
            _main = main;
            dirVir = new FolderBrowserDialog();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            dirVir.Description = "Seleccionar el Directorio de salida.";

            DialogResult dirSelected = dirVir.ShowDialog(this);

            if (dirSelected == DialogResult.OK)
            {
                directoryOutPath = dirVir.SelectedPath;
            }
        }

        private bool MakeFileProperties()
        {
            int dimension;
            if (directoryOutPath != null
                && int.TryParse(textBoxDimesion.Text, out dimension))
            {
                string[] spliting = _fullfileInpath.Split('\\');

                MakeFile(_fullfileInpath, directoryOutPath,spliting[spliting.Length-1], dimension);
                return true;
            }
            else
            {
                MessageBox.Show("Para comezar, complete la dirección del fichero de salida y la dimensión de las matrices.");
                return false;
            }
        }

        private void MakeFile(string inFilePath, string fulloutpath, string fileOutName, int dimension)
        {
            ReadExecDoc linealizar;
            linealizar = new ReadExecDoc(inFilePath, fulloutpath, fileOutName, dimension);
            progressBar1.Maximum = (int)linealizar.Percent/100;
            foreach (var percent in linealizar.Numbers())
            {
                progressBar1.Value = (int)percent / 100;
            }

            linealizar.CloseDocucument();
        }

        private void ButtonMake_Click(object sender, EventArgs e)
        {
            try
            {
                if (MakeFileProperties())
                {
                    Close();
                    _main.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
            
        }
    }
}
