using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace extractChart
{
    public partial class Form1 : Form
    {

        #region Variables

        string DirectorioTemp;

        #endregion

        #region Constructor

        public Form1()
        {
            InitializeComponent();

            DirectorioTemp = Path.Combine(Environment.CurrentDirectory, "temp");

            var x = 0;
        }

        #endregion

        #region Metodos

        private void extraerimagen()
        {
            try
            {
                object lMissing = System.Reflection.Missing.Value;

                var lExcelApp = new Excel.Application()
                {
                    Visible = true
                };

                string lArchivo = Environment.CurrentDirectory + "\\test.xlsx";

                Excel.Workbook lLibro =
                        lExcelApp.Workbooks.Open(lArchivo, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing, lMissing);

                for (int i = 1; i <= lLibro.Worksheets.Count; i++)// i = num de hoja
                {
                    if (lLibro.Worksheets[i].ChartObjects().Count() > 0)
                    {
                        for (int j = 1; j <= lLibro.Worksheets[i].ChartObjects().Count(); j++)// j = numero de grafico
                        {
                            string lNombreGraf = $"grafico{i}-{j}";

                            string lNombreArchivo = Path.Combine(DirectorioTemp, $"{lNombreGraf}.jpg");

                            lLibro.Worksheets[i].ChartObjects(j).Activate();
                            lLibro.Worksheets[i].ChartObjects(j).Chart.Export(lNombreArchivo, "JPG", false);

                            cmbGrafico.Items.Add(lNombreGraf);
                        }
                    }
                }

                lLibro.Close();
                lExcelApp.Quit();
                lExcelApp = null;

                if (cmbGrafico.Items.Count > 0)
                {
                    btnExportar.Enabled = true;
                    cmbGrafico.Enabled = true;
                    this.cmbGrafico.SelectedIndexChanged += new System.EventHandler(this.cmbGrafico_SelectedIndexChanged);
                    cmbGrafico.SelectedIndex = 0;
                }
            }
            catch (Exception lExcp)
            {
                throw new Exception("Error al extraer la imagen\n" + lExcp.Message);
            }
        }

        private void insertarimagen(string pArchivoImagen)
        {
            try
            {
                object lMissing = System.Reflection.Missing.Value;

                Word.Application lWordApp = new Word.Application()
                {
                    Visible = true
                };

                string lArchivo = Path.Combine(Environment.CurrentDirectory, "test.docx");

                Word.Document lDocumento = lWordApp.Documents.Add(lArchivo);

                object lBookMark = "marcador_imagen";
                lDocumento.Bookmarks.get_Item(ref lBookMark).Range.InlineShapes.AddPicture(pArchivoImagen);

                //lDocumento.PrintOut();

                lWordApp = null;
            }
            catch (Exception lExcp)
            {
                throw new Exception("Error al insertar la imagen en el doc\n" + lExcp.Message);
            }
        }

        #endregion

        #region Botones

        private void btnExtraer_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(DirectorioTemp) == false)
                {
                    Directory.CreateDirectory(DirectorioTemp);
                }

                cmbGrafico.Items.Clear();
                imgGrafico.Image = null;

                extraerimagen();
            }
            catch (Exception lExcp)
            {
                MessageBox.Show(lExcp.Message);
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            string lArchivo = Path.Combine(DirectorioTemp, $"{cmbGrafico.Text}.jpg");
            insertarimagen(lArchivo);
        }

        #endregion

        #region Eventos

        private void cmbGrafico_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                
                string lArchivo = Path.Combine(DirectorioTemp,  $"{cmbGrafico.Text}.jpg");

                string lExtract = cmbGrafico.Text.Replace("grafico", "");

                lblHoja.Text = lExtract.Substring(0, lExtract.IndexOf('-'));

                lblGraf.Text = lExtract.Substring(lExtract.IndexOf('-') + 1);

                if (File.Exists(lArchivo) == true)
                {
                    Image img;
                    using (var bmpTemp = new Bitmap(lArchivo))
                    {
                        img = new Bitmap(bmpTemp);
                    }
                    imgGrafico.Image = img;
                }
            }
            catch (Exception lExcp)
            {
                MessageBox.Show(lExcp.Message);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Directory.Exists(DirectorioTemp) == true)
            {
                //var x = Directory.GetFiles(DirectorioTemp);

                //foreach(string archivo in x)
                //{
                //    File.Delete(archivo);
                //}

                Directory.Delete(DirectorioTemp, true);
            }

        }
        
        #endregion

    }
}