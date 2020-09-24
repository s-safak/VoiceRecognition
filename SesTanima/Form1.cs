using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Threading;
using System.Runtime.InteropServices;


namespace SesTanima
{
    public partial class Form1 : Form
    {
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;
        private const int WM_SETTEXT = 0x0C;
        private const int WM_KEYDOWN = 100;
        private const int WM_KEYUP = 101;
        private const int WM_CHAR = 258;
        private const int WM_BUTTONDOWN = 0x0100;
        private const int WM_BUTTONUP = 0x0101;
        const uint WM_MOUSEMOVE = 0x0200;

        [DllImport("User32.dll")]
        public static extern Int32 FindWindow(String lpClassName, String lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, string lParam);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
        //[DllImport("user32", EntryPoint = "PostMessageA")]
        //private static extern int PostMessage(int hWnd, int uMsg, int wParam, int lParam);
        [DllImport("User32.Dll", EntryPoint = "PostMessageA")]
        static extern bool PostMessage(int hWnd, uint msg, int wParam, int lParam);
        [DllImport("User32.dll")]
        public static extern int SendMessage(int hWnd, uint wMsg, IntPtr wParam, IntPtr lParam);


        // Global Deðiþkenler
        private SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();

        // Yapýcý metot
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadGrammars();
            StartRecognition();
        }
        // Metotlar
        // Tanýma motoru tarafýndan tanýnmasý gereken kelimeleri belirtiyoruz.
        private void LoadGrammars() 
        {
            Choices choices = new Choices(new string[] {textBox2.Text , textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text});
            GrammarBuilder grammarBuilder = new GrammarBuilder(choices);
            Grammar grammar = new Grammar(grammarBuilder);
            recognizer.LoadGrammar(grammar);
        }

        // Ses tanýma iþlemi sýrasýnda ve sonrasýnda meydana gelecek olaylarý belirtiyoruz.
        // Tanýma iþlemini baþlatýyoruz.
        private void StartRecognition() 
        {

            // Belirli sesleri tanýma iþlemindeki ana olaylar
            recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);
            recognizer.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRecognitionRejected);
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            recognizer.RecognizeCompleted += new EventHandler<RecognizeCompletedEventArgs>(recognizer_RecognizeCompleted);

            // recognizer.SetInputToDefaultAudioDevice() Bu metotun bir Thread içinde çalýþtýrýlmasý gerekmektedir!
            // Ses tanýma iþlemini baþlatýyoruz.
            Thread t1 = new Thread(delegate()
            {            
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.RecognizeAsync(RecognizeMode.Single);
            });

            t1.Start();

        }


        // Olaylar
        // Kullanýcý konuþmaya baþladýðý anda tetiklenen olay
        private void recognizer_SpeechDetected(object sender, SpeechDetectedEventArgs e) 
        {
            textBox1.Text = "Ses Tanýnýyor";
            
        }

        // Kullanýcýnýn konuþtuðu kelimeler gramerde bulunuyorsa tetiklenen olay
        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) 
        {
            if (e.Result.Text == textBox2.Text)
            {
                IntPtr ip1 = new IntPtr(0x31); //1
                IntPtr ip2 = new IntPtr(0x12); //ALT
                int ihwnd = FindWindow(null, "World of Warcraft");

                SendMessage(ihwnd, WM_BUTTONDOWN, ip2, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONDOWN, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip2, IntPtr.Zero);
                pictureBox1.Image = Properties.Resources.lightsOn;
                textBox1.Text = "";
                Thread.Sleep(2);
                pictureBox1.Image = Properties.Resources.lightsOff;
            }
            else if (e.Result.Text == textBox3.Text)
            {
                IntPtr ip1 = new IntPtr(0x32); //2
                IntPtr ip2 = new IntPtr(0x12); //ALT
                int ihwnd = FindWindow(null, "World of Warcraft");

                SendMessage(ihwnd, WM_BUTTONDOWN, ip2, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONDOWN, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip2, IntPtr.Zero);
                pictureBox1.Image = Properties.Resources.lightsOn;
                textBox1.Text = "";
                Thread.Sleep(2);
                pictureBox1.Image = Properties.Resources.lightsOff;
            }
            else if (e.Result.Text == textBox4.Text)
            {
                IntPtr ip1 = new IntPtr(0x33); //3
                IntPtr ip2 = new IntPtr(0x12); //ALT
                int ihwnd = FindWindow(null, "World of Warcraft");

                SendMessage(ihwnd, WM_BUTTONDOWN, ip2, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONDOWN, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip2, IntPtr.Zero);
                pictureBox1.Image = Properties.Resources.lightsOn;
                textBox1.Text = "";
                Thread.Sleep(2);
                pictureBox1.Image = Properties.Resources.lightsOff;
            }
            else if (e.Result.Text == textBox5.Text)
            {
                IntPtr ip1 = new IntPtr(0x34); //4
                IntPtr ip2 = new IntPtr(0x12); //ALT
                int ihwnd = FindWindow(null, "World of Warcraft");

                SendMessage(ihwnd, WM_BUTTONDOWN, ip2, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONDOWN, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip2, IntPtr.Zero);
                pictureBox1.Image = Properties.Resources.lightsOn;
                textBox1.Text = "";
                Thread.Sleep(2);
                pictureBox1.Image = Properties.Resources.lightsOff;
            }
            else if (e.Result.Text == textBox6.Text)
            {
                IntPtr ip1 = new IntPtr(0x35); //5
                IntPtr ip2 = new IntPtr(0x12); //ALT
                int ihwnd = FindWindow(null, "World of Warcraft");

                SendMessage(ihwnd, WM_BUTTONDOWN, ip2, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONDOWN, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip1, IntPtr.Zero);
                SendMessage(ihwnd, WM_BUTTONUP, ip2, IntPtr.Zero);
                pictureBox1.Image = Properties.Resources.lightsOn;
                textBox1.Text = "";
                Thread.Sleep(2);
                pictureBox1.Image = Properties.Resources.lightsOff;
            }

            textBox1.Text = e.Result.Text;

        }

        // Konuþulan kelimeler gramerde bulumuyorsa tetiklenen olay
        private void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e) 
        {
            textBox1.Text = "Ses Tanýma Ýþlemi Baþarýsýz.";
        }

        // Tanýma iþlemi baþarýlý olsun veya olmasýný sonuçlandýðýnda tetiklenen olay
        private void recognizer_RecognizeCompleted(object sender, RecognizeCompletedEventArgs e) 
        {
            recognizer.RecognizeAsync();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {                
                LoadGrammars();                
                textBox2.Enabled = false;
            }
            else textBox2.Enabled = true;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                LoadGrammars();
                textBox3.Enabled = false;
            }
            else textBox3.Enabled = true;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                LoadGrammars();
                textBox4.Enabled = false;
            }
            else textBox4.Enabled = true;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                LoadGrammars();
                textBox5.Enabled = false;
            }
            else textBox5.Enabled = true;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                LoadGrammars();
                textBox6.Enabled = false;
            }
            else textBox6.Enabled = true;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void textBox5_Click(object sender, EventArgs e)
        {
            textBox5.Text = "";
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
        }

    }
}