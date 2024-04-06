using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using InputManager;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace CommandeVocal
{

    public partial class Form1 : Form
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);
        SpeechRecognitionEngine recognization = new SpeechRecognitionEngine(new CultureInfo("fr-FR"));
        SpeechRecognitionEngine confirmation = new SpeechRecognitionEngine();
        public bool resolve { get; set; } = false;
        public string typeCommande { get; set; } = "null";
        public Form1()
        {

            InitializeComponent();
            richTextBox1.Text += "Initialisation";
            comboBox1.SelectedIndex = 0;


            Choices cmds = new Choices();
            cmds.Add(new string[] { "Salut ça va", "Que fait tu", "Arrêter commandes vocales", "Groupe d'arme suivant",
                                    "Train d'atterisage","Moteur","Accélérer","Nettoyer","Ouvre Opera","Insulte Nico",
                                    "Oui","Non","verification"});
            GrammarBuilder builder = new GrammarBuilder();
            builder.Append(cmds);
            Grammar grammar = new Grammar(builder);

            recognization.LoadGrammarAsync(grammar);
            recognization.SetInputToDefaultAudioDevice();
            recognization.SpeechRecognized += recognized_voice;


            Grammar grammarConf = new Grammar(builder);
            confirmation.LoadGrammarAsync(grammarConf);
            confirmation.SetInputToDefaultAudioDevice();
            confirmation.SpeechRecognized += confirmation_voice;
            

        }
        public async void confirmation_voice(object sender, SpeechRecognizedEventArgs e)
        {

            string resultat = e.Result.Text;
            if (resultat == "Oui")
            {
                richTextBox1.Text = "\n POSITIVE !!!";
                resolve = true;
            }
            else if (resultat == "Non")
            {
                richTextBox1.Text = "\n confirmation négative !";
                resolve = false;
            }
        }
        private async void recognized_voice(object sender, SpeechRecognizedEventArgs e){

            string resultat = e.Result.Text;

            if (typeCommande == "Elite Dangerous") {

                switch (e.Result.Text)
                {
                    case "Salut ça va":
                        richTextBox1.Text += "\n Je vais bien, merci.";
                        break;
                    case "Que fait tu":
                        richTextBox1.Text += "\n Je suis en cours de maj";
                        break;
                    case "Arrêter commandes vocales":
                        richTextBox1.Text += "\n Arrêt en cours";
                        button2_Click(this, new EventArgs());
                        break;
                    case "Groupe d'arme suivant":
                        Keyboard.KeyPress(Keys.N);
                        //WriteExampleText(this);
                        break;
                    case "Train d'atterisage":
                        Keyboard.KeyPress(Keys.L);
                        break;
                    case "Moteur":
                        Keyboard.KeyPress(Keys.Up);
                        break;
                    case "Accélérer":
                        //TODO
                        break;
                }
            }
            if (typeCommande == "Bureau")
            {
                //BUREAU
                if (resultat == "Nettoyer")
                {
                    richTextBox1.Text = "";
                }
                if (resultat == "Ouvre Opera")
                {
                    string winpath = Environment.GetEnvironmentVariable("windir");
                    string path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

                    Process.Start("C:/Users/Ady/AppData/Local/Programs/Opera GX/launcher.exe");
                }
                if (resultat == "Conctacte Nico")
                {
                    confirmation.Recognize(TimeSpan.FromSeconds(3));

                    richTextBox1.Text += "\n Verification: " + resolve;
                    if (resolve)
                    {
                        richTextBox1.Text += "\n CHEF OUI CHEF ";
                        conctactNico(this);
                        confirmation.RecognizeAsyncStop();
                    }
                    else
                    {
                        richTextBox1.Text += "\n annulation...";
                        confirmation.RecognizeAsyncStop();

                    }
                }
                if (resultat == "verification")
                {
                    confirmation.Recognize(TimeSpan.FromSeconds(3));

                    richTextBox1.Text += "\n Verification: " + resolve;
                    if (resolve)
                    {
                        richTextBox1.Text += "\n OMG RESOLVED: ";
                        confirmation.RecognizeAsyncStop();
                    }
                    else
                    {
                        richTextBox1.Text += "\n NOPE";
                        confirmation.RecognizeAsyncStop();

                    }
                }
            }

        }



        private void button1_Click(object sender, EventArgs e)
        {
            recognization.RecognizeAsync(RecognizeMode.Multiple);
            richTextBox1.Text += "\n Debut de l'écoute.";
            button2.Enabled = true;
            button1.Enabled = false;
            

        }
        private void button2_Click(object sender, EventArgs e)
        {
            recognization.RecognizeAsyncStop();
            richTextBox1.Text += "\n Fin de l'écoute.";
            button2.Enabled = false;
            button1.Enabled = true;

        }

        private void ExempleKeyboard(object sender)
        {
            Keyboard.KeyDown(Keys.Z);
            Thread.Sleep(1000);
            Keyboard.KeyUp(Keys.Z);
            Keyboard.KeyDown(Keys.S);
            Thread.Sleep(1000);
            Keyboard.KeyUp(Keys.S);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Text += "\n Boutton 3 !";
            if (comboBox1.SelectedIndex != -1)
            {
                richTextBox1.Text += "WOW TU A SELECTIONE" + comboBox1.SelectedIndex;
                if(comboBox1.SelectedIndex == 0)
                {
                    richTextBox1.Text += "Mode: Bureau";
                    typeCommande = "Bureau";
                }
                if (comboBox1.SelectedIndex == 1)
                {
                    richTextBox1.Text += "Mode: Elite Dangerous";
                    typeCommande = "Elite Dangerous";
                }
            }
            else
            {
                MessageBox.Show("NON MAIS HO LA");
            }
        }
        private void conctactNico(object sender)
        {
            //MessageBox.Show("HELLO THERE");
            Mouse.Move(1630, 1060);
            Mouse.PressButton(Mouse.MouseKeys.Left);
            Thread.Sleep(10);
            Keyboard.KeyPress(Keys.Escape);
            Thread.Sleep(10);
            Keyboard.KeyDown(Keys.LControlKey);
            Keyboard.KeyDown(Keys.K);
            Thread.Sleep(20);
            Keyboard.KeyUp(Keys.LControlKey);
            Keyboard.KeyUp(Keys.K);
            Thread.Sleep(20);

            List<Keys> victime = new List<Keys>();
            victime.Append(Keys.N);
            victime.Append(Keys.I);
            victime.Append(Keys.C);
            victime.Append(Keys.O);

            int[] myArray = new int[4];
            Keys[] insulte01 = new Keys[] { Keys.N, Keys.I, Keys.C, Keys.O, Keys.Space, Keys.T, Keys.Space, Keys.M, Keys.O, Keys.C, Keys.H, Keys.E}; //pas trop d'inspi...
            //todo gerer la reconnaisance vocal d'un message custom

            Keyboard.KeyPress(Keys.N);
            Thread.Sleep(10);
            Keyboard.KeyPress(Keys.I);
            Thread.Sleep(10);
            Keyboard.KeyPress(Keys.C);
            Thread.Sleep(10);
            Keyboard.KeyPress(Keys.O);
            Thread.Sleep(200);
            Keyboard.KeyPress(Keys.Enter);
            Thread.Sleep(10);


            foreach (Keys k in insulte01)
            {
                Keyboard.KeyPress(k);
                Thread.Sleep(20);
            }
            Thread.Sleep(200);
            Keyboard.KeyPress(Keys.Enter);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            conctactNico(this);
        }

        private void deccolageAvantPoste()
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
