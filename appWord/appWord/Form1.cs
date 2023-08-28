using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace appWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialogo = new SaveFileDialog();
            dialogo.Filter = "Documentos de Word (*.docx)|*.docx"; // Filtro para seleccionar solo archivos de Word
            dialogo.DefaultExt = ".docx"; // Extensión predeterminada
            dialogo.AddExtension = true; // Agregar automáticamente la extensión si no se proporciona

            if (dialogo.ShowDialog() == DialogResult.OK)
            {
                string ruta = dialogo.FileName; // Obtener la ruta seleccionada por el usuario
                string dato = txtDato.Text;

                try
                {
                    var wordApp = new Word.Application();
                    wordApp.Visible = true;
                    var doc = wordApp.Documents.Add();
                    doc.Content.Text = dato;
                    doc.SaveAs2(ruta);
                    doc.Close();

                    MessageBox.Show("Documento guardado exitosamente.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar el documento:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
