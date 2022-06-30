using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using TikaOnDotNet.TextExtraction;
using BitMiracle.Docotic.Pdf;
using Emgu;
using Emgu.CV;
using Emgu.CV.Util;
using Emgu.CV.OCR;
using Emgu.CV.Structure;
using Emgu.Util;







namespace diplom
{
    public partial class Form1 : Form
    {   
        int count = 0;
        string path;
        string tessdata;
        public string target_word;

        bool s1=true;
        


        public Form1()
        {
            InitializeComponent();

        }

        public void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {

            if (radioButton1.Checked == true)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "pdf files (*.pdf)|*.pdf";
                if (ofd.ShowDialog() == DialogResult.OK)
                {   
                    textBox1.Text = ofd.FileName;
                }

            }
            else if (checkBox1.Checked == true || checkBox2.Checked == true || checkBox3.Checked == true)
            {
                FolderBrowserDialog q = new FolderBrowserDialog();
                if (q.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = q.SelectedPath;
                }
            }
            else
            {
                MessageBox.Show("Выберите что искать", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }




        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                if (!System.IO.File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Файл не найден", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (textBox1.Text.EndsWith("pdf"))
                    {
                        
                        await pdftoimage();
                    }
                    else
                    {
                        MessageBox.Show("Указанный файл не pdf формата", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            else
            {
                if (!Directory.Exists(textBox1.Text))
                {
                    MessageBox.Show("Директория не найдена", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    progressBar1.Maximum = 0;
                    await OnClickSearch(textBox1.Text);                   

                }
            }


        }


        private async Task OnClickSearch(string directory)
        {
            listBox1.Items.Clear();

            s1 = true;

            if (checkBox1.Checked == true || checkBox2.Checked == true || checkBox3.Checked == true)
            {
                string[] alltxt = Directory.GetFiles(directory, "*.txt", SearchOption.AllDirectories);
                string[] allpdf = Directory.GetFiles(directory, "*.pdf", SearchOption.AllDirectories);
                string[] alldocx = Directory.GetFiles(directory, "*.docx", SearchOption.AllDirectories);
               
                
                

                if (checkBox3.Checked == true)
                {   
                    
                    if (alltxt.Length != 0)
                    {
                        
                        
                        await SearchTxt(alltxt);
                        count += alltxt.Length;
                        

                        
                    }
                }
                if (checkBox1.Checked == true)
                {
                    
                    if (allpdf.Length != 0)
                    {
                        
                        await SearchPdf(allpdf);
                        count += allpdf.Length;
                    }
                }
                if (checkBox2.Checked == true)
                {
                    
                    if (alldocx.Length != 0)
                    {
                        
                        await SearchDocx(alldocx);
                        count += alldocx.Length;
                    }
                }

                if (count == 0)
                {
                    MessageBox.Show("Совпадений нет", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            else
            {
                MessageBox.Show("Выберите что искать", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        private void Form1_Load(object sender, EventArgs e)
        {

            textBox1.Text = Directory.GetCurrentDirectory();
            tessdata = $"{Directory.GetCurrentDirectory()}\\testdata\\";

        }

        private async Task SearchTxt(string[] alltxt)
        {
            string target_word = textBox2.Text;
            progressBar1.Maximum += alltxt.Length;


            if (textBox2.Text == "")
            {
                foreach (string file in alltxt)
                {
                    
                   
                    if (s1 == false)
                    {
                        break;
                    }
                    Invoke((Action)(() => listBox1.Items.Add(file)));
                    Invoke((Action)(() => progressBar1.Value += 1));
                }
                
            }
            else
            {

                foreach (string file in alltxt)
                {
                    
                    if (s1 == false)
                    {
                        break;
                    }
                    using (StreamReader sr = new StreamReader(file))
                    {
                        while (!sr.EndOfStream)
                        {
                            string str = await sr.ReadLineAsync();
                            if (str.IndexOf(target_word, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                Invoke((Action)(() => listBox1.Items.Add(file)));
                                break;
                            }
                            

                        }
                    }
                    Invoke((Action)(() => progressBar1.Value += 1));
                }
                
            }
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string str = listBox1.SelectedItem.ToString();
                Process.Start(str);
                listBox1.SetSelected(listBox1.SelectedIndex, false);
            }
            catch
            {

            }

        }

        private Task SearchPdf(string[] allpdf)
        {
            string target_word = textBox2.Text;
            progressBar1.Maximum += allpdf.Length;

            return Task.Run(() =>
                {
                    if (textBox2.Text == "")
                    {
                        foreach (string file in allpdf)
                        {
                            if (s1 == false)
                            {
                                break;
                            }
                            Invoke((Action)(() => listBox1.Items.Add(file)));
                            Invoke((Action)(() => progressBar1.Value += 1));

                        }
                        
                    }
                    else
                    {
                        SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                        foreach (string file in allpdf)
                        {
                            if (s1 == false)
                            {
                                break;
                            }

                            f.OpenPdf(file);
                            string str = f.ToText();
                            if (str.IndexOf(target_word, StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                Invoke((Action)(() => listBox1.Items.Add(file)));
                            }
                            Invoke((Action)(() => progressBar1.Value += 1));
                        }
                        
                    }


                    
                });
            


             
        }


        private Task SearchDocx(string[] alldocx)
        {
            string target_word = textBox2.Text;
            progressBar1.Maximum += alldocx.Length;


            if (textBox2.Text =="")
                {
                    return Task.Run(() => {
                        foreach (string file in alldocx)
                        {

                            if (s1 == false)
                            {
                                break;
                            }
                            Invoke((Action)(() => listBox1.Items.Add(file)));
                            Invoke((Action)(() => progressBar1.Value += 1));
                        }
                    });
                    

                }
                else
                {
                    return Task.Run(() => {
                        var textExtractor = new TextExtractor();
                        foreach (string file in alldocx)

                        {
                            if (s1 == false)
                            {
                                break;
                            }
                            try
                            {
                                string str = textExtractor.Extract(file).Text;
                                if (str.IndexOf(target_word, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                {
                                    Invoke((Action)(() => listBox1.Items.Add(file)));
                                }
                                
                            }
                            catch(TextExtractionException)
                            {
                                continue;
                                
                            }
                            Invoke((Action)(() => progressBar1.Value += 1));

                        }
                    });
                    


                }
                
            
        }



        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {

                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                MessageBox.Show("Можно выбрать только один файл pdf", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                radioButton1.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                radioButton1.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                radioButton1.Checked = false;
            }

        }


        private Task pdftoimage()            
        {
            listBox1.Items.Clear();
            Invoke((Action)(() => progressBar1.Maximum = 0));
            s1 = true;

            BitMiracle.Docotic.LicenseManager.AddLicenseData("6ECD4-38S0B-ZJVO3-E7B40-QTBJD");
            string target_word = textBox2.Text;
            string file = textBox1.Text;
            path = DateTime.Now.ToString();
            path = path.Replace(":", "-");
            Directory.CreateDirectory(path);        
            
           
                return Task.Run(() =>
                {
                    try
                    {   
                        using (var pdf = new PdfDocument(file))
                        {
                            PdfDrawOptions options = PdfDrawOptions.Create();
                            options.BackgroundColor = new PdfRgbColor(255, 255, 255);
                            options.HorizontalResolution = 300;
                            options.VerticalResolution = 300;
                            Invoke((Action)(() => progressBar1.Maximum += pdf.PageCount));
                            
                            for (int i = 0; i < pdf.PageCount; ++i)
                            {
                                
                                if (s1 == false)
                                {
                                    break;
                                }
                                pdf.Pages[i].Save($@"{path}\page_{i}.png", options);
                                string page = $@"{path}\page_{i}.png";


                                if(textBox2.Text=="")
                                {
                                    Invoke((Action)(() => listBox1.Items.Add($@"{Directory.GetCurrentDirectory()}\{page}")));
                                    Invoke((Action)(() => progressBar1.Value += 1));
                                }
                                else
                                {   
                                    Tesseract tesseract = new Tesseract(@tessdata, "rus", OcrEngineMode.TesseractLstmCombined);

                                    tesseract.SetImage(new Image<Bgr, byte>(page));

                                    tesseract.Recognize();

                                    string text = tesseract.GetUTF8Text();
                                    if (text.IndexOf(target_word, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                    {
                                        Invoke((Action)(() => listBox1.Items.Add($@"{Directory.GetCurrentDirectory()}\{page}")));

                                    }

                                    tesseract.Dispose();
                                    Invoke((Action)(() => progressBar1.Value += 1));
                                }
                                
                            }
                            


                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Ошибка выполнения", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                    }
                    
                    
                });
            
            
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            s1 = false;
                      
        }
    }
}