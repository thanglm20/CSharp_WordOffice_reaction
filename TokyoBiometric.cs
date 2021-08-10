using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Graph = Microsoft.Office.Interop.Graph;
using Microsoft.Office.Interop.Graph;
using System.Runtime.InteropServices;
using System.Threading;
using System.Security.AccessControl;
using System.Security.Principal;
/*
  
  Cách Cài đặt trực tiếp Microsoft.Office.Interop.Word.dll vào Windows.
  Sao chép tệp .DLL vào thư mục C:\Windows\System32 (nếu sử dụng HĐH 32 bit)
  Sao chép tệp .DLL vào thư mục C:\Windows\SysWOW64 (nếu sử dụng HĐH 64 bit)
  Cài đặt DLL đã được hoàn thành!

    WT / WS / WP / WE / WI / WD / WX: đại bàng
    AT/AS: chim công
    UL: chim cú
    RL: bồ câu

    data.smk ( đai bàng) , config.daa ( cú) , user.rtk ( công) , biotech.toy ( bồ câu) 
 
 */
namespace TokyoBiometric
{
    public partial class TokyoBiometric : Form 
    {
       
        public TokyoBiometric()
        {
            InitializeComponent();
        }
        bool g_flCalculated = false;
        string g_pathReferenceData = @"C:\Program Files (x86)\TokyoBiometric\TokyoBiometric\TokyoBiometric";
        //string g_pathReferenceData = @"D:\MyProject\TokyoBiometric\data\config";
        string[] g_arListPathImages = new string[20];
        Word.Document g_myWordDoc = null;
        Word.Application g_wordApp = null;
        public static void SetFullAccessPermissionsForEveryone(string directoryPath)
        {
            //Everyone Identity
            IdentityReference everyoneIdentity = new SecurityIdentifier(WellKnownSidType.WorldSid,
                                                       null);

            DirectorySecurity dir_security = Directory.GetAccessControl(directoryPath);

            FileSystemAccessRule full_access_rule = new FileSystemAccessRule(everyoneIdentity,
                            FileSystemRights.FullControl, InheritanceFlags.ContainerInherit |
                             InheritanceFlags.ObjectInherit, PropagationFlags.None,
                             AccessControlType.Allow);
            dir_security.AddAccessRule(full_access_rule);

            Directory.SetAccessControl(directoryPath, dir_security);
        }
        private void Form1_Load(object sender, EventArgs e)
        {


            SetFullAccessPermissionsForEveryone(g_pathReferenceData);

            //g_wordApp =  new Word.Application();
            cbbL1.Text = "AT";
            cbbL2.Text = "AT";
            cbbL3.Text = "AT";
            cbbL4.Text = "AT";
            cbbL5.Text = "AT";
            cbbR1.Text = "AT";
            cbbR2.Text = "AT";
            cbbR3.Text = "AT";
            cbbR4.Text = "AT";
            cbbR5.Text = "AT";

        }
        
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            //object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        
        private bool checkTextBoxIsNumber()
        {

            if (float.TryParse(txtL1.Text.ToString(), out _) != true || (float.Parse(txtL1.Text.ToString()) > 100.0 || float.Parse(txtL1.Text.ToString()) <= 0))
            { txtL1.BackColor = Color.Red; return false; }
            else { txtL1.BackColor = Color.White; }

            if (float.TryParse(txtL2.Text.ToString(), out _) != true || (float.Parse(txtL2.Text.ToString()) > 100.0 || float.Parse(txtL2.Text.ToString()) <= 0))
            { txtL2.BackColor = Color.Red; return false;} 
            else{ txtL2.BackColor = Color.White; }

            if (float.TryParse(txtL3.Text.ToString(), out _) != true || (float.Parse(txtL3.Text.ToString()) > 100.0 || float.Parse(txtL3.Text.ToString()) <= 0))
            { txtL3.BackColor = Color.Red; return false; }
            else { txtL3.BackColor = Color.White; }

            if (float.TryParse(txtL4.Text.ToString(), out _) != true || (float.Parse(txtL4.Text.ToString()) > 100.0 || float.Parse(txtL4.Text.ToString()) <= 0))
            { txtL4.BackColor = Color.Red; return false; }
            else { txtL4.BackColor = Color.White; }

            if (float.TryParse(txtL5.Text.ToString(), out _) != true || (float.Parse(txtL5.Text.ToString()) > 100.0 || float.Parse(txtL5.Text.ToString()) <= 0))
            { txtL5.BackColor = Color.Red; return false; }
            else { txtL5.BackColor = Color.White; }


            if (float.TryParse(txtR1.Text.ToString(), out _) != true || (float.Parse(txtR1.Text.ToString()) > 100.0 || float.Parse(txtR1.Text.ToString()) <= 0))
            { txtR1.BackColor = Color.Red; return false; }
            else { txtR1.BackColor = Color.White; }

            if (float.TryParse(txtR2.Text.ToString(), out _) != true || (float.Parse(txtR2.Text.ToString()) > 100.0 || float.Parse(txtR2.Text.ToString()) <= 0))
            { txtR2.BackColor = Color.Red; return false; }
            else { txtR2.BackColor = Color.White; }

            if (float.TryParse(txtR3.Text.ToString(), out _) != true || (float.Parse(txtR3.Text.ToString()) > 100.0 || float.Parse(txtR3.Text.ToString()) <= 0))
            { txtR3.BackColor = Color.Red; return false; }
            else { txtR3.BackColor = Color.White; }

            if (float.TryParse(txtR4.Text.ToString(), out _) != true || (float.Parse(txtR4.Text.ToString()) > 100.0 || float.Parse(txtR4.Text.ToString()) <= 0))
            { txtR4.BackColor = Color.Red; return false; }
            else { txtR4.BackColor = Color.White; }

            if (float.TryParse(txtR5.Text.ToString(), out _) != true || (float.Parse(txtR5.Text.ToString()) > 100.0 || float.Parse(txtR5.Text.ToString()) <= 0))
            { txtR5.BackColor = Color.Red; return false; }
            else { txtR5.BackColor = Color.White; }

            return true;
        }
        public static Tuple<string, string> getCharacter(string[] arCharacters)
        {
            Dictionary<string, int> mp = new Dictionary<string, int>();
            for(int i = 0; i < arCharacters.Length; i++)
            {
                if (arCharacters[i] == "WT" || arCharacters[i] == "WS" || arCharacters[i] == "WP" ||
                        arCharacters[i] == "WE" || arCharacters[i] == "WI" || arCharacters[i] == "WD" || arCharacters[i] == "WX")
                {
                    arCharacters[i] = "dai_bang";
                }
                else if (arCharacters[i] == "AT" || arCharacters[i] == "AS")
                {
                    arCharacters[i] = "chim_cong";
                }
                else if (arCharacters[i] == "UL")
                {
                    arCharacters[i] = "chim_cu";
                }
                else if (arCharacters[i] == "RL")
                {
                    arCharacters[i] = "bo_cau";
                }
            }     
            // Traverse through array elements and
            // count frequencies
            for (int i = 0; i < arCharacters.Length; i++)
            {
                if (mp.ContainsKey(arCharacters[i]))
                {
                    var val = mp[arCharacters[i]];
                    mp.Remove(arCharacters[i]);
                    mp.Add(arCharacters[i], val + 1);
                }
                else
                {
                    mp.Add(arCharacters[i], 1);
                }
            }

            var myList = mp.ToList();
            myList.Sort((pair1, pair2) => pair1.Value.CompareTo(pair2.Value));
            // Traverse through map and print frequencies
            //foreach (KeyValuePair<string, int> entry in myList)
            //{
            //    Console.WriteLine(entry.Key + " " + entry.Value);
            //}
            //Console.WriteLine(myList.Count);
            var characters = Tuple.Create("", "");
            if (myList.Count == 0)
            {
                return Tuple.Create(" ", " ");
            }
            else if (myList.Count == 1)
            {
                return Tuple.Create(myList[0].Key, myList[0].Key);
            }
            else if (myList.Count > 1)
            {
                int max = myList.Count - 1;
                return Tuple.Create(myList[max].Key, myList[max -1].Key);
            }
            return characters;
        }
        private void calculate()
        {
            try
            {
                if(checkTextBoxIsNumber() == true)
                {

                    float L1 = float.Parse(txtL1.Text.ToString());
                    float L2 = float.Parse(txtL2.Text.ToString());
                    float L3 = float.Parse(txtL3.Text.ToString());
                    float L4 = float.Parse(txtL4.Text.ToString());
                    float L5 = float.Parse(txtL5.Text.ToString());
                    float R1 = float.Parse(txtR1.Text.ToString());
                    float R2 = float.Parse(txtR2.Text.ToString());
                    float R3 = float.Parse(txtR3.Text.ToString());
                    float R4 = float.Parse(txtR4.Text.ToString());
                    float R5 = float.Parse(txtR5.Text.ToString());

                    float A27 = L1 + L2 + L3 + L4 + L5 + R1 + R2 + R3 + R4 + R5;
                    float A1 = (L1 + R1) / A27 * 100;
                    float A2 = (L2 + R2) / A27 * 100;
                    float A3 = (L3 + R3) / A27 * 100;
                    float A4 = (L4 + R4) / A27 * 100;
                    float A5 = (L5 + R5) / A27 * 100;
                    float A6 = L1 + L2 + L3 + L4 + L5;
                    float A7 = R1 + R2 + R3 + R4 + R5;
                    float A8 = A7 / (A6 + A7) * 100;
                    float A9 = A6 / (A6 + A7) * 100;
                    float A10 = R1 / A27 * 100;
                    float A11 = L1 / A27 * 100;
                    float A12 = R2 / A27 * 100;
                    float A13 = L2 / A27 * 100;
                    float A14 = R3 / A27 * 100;
                    float A15 = L3 / A27 * 100;
                    float A16 = R4 / A27 * 100;
                    float A17 = L4 / A27 * 100;
                    float A18 = R5 / A27 * 100;
                    float A19 = L5 / A27 * 100;
                    float A20 = (A18 + A19) / (A14 + A15 + A16 + A17 + A18 + A19) * 100;
                    float A21 = (A16 + A17) / (A14 + A15 + A16 + A17 + A18 + A19) * 100;
                    float A22 = (A14 + A15) / (A14 + A15 + A16 + A17 + A18 + A19) * 100;
                    float A23 = (A10 + A11 + A12);
                    float A24 = (A12 + A13);
                    float A25 = (A11 + A12);
                    float A26 = (A13 + A18 + A19);
                    double A28 = (A17 + A16) / A27 * 1.1 * 100;
                    double A29 = (A18 + A13) / A27 * 1.05 * 100;
                    double A30 = (A18 + A19) / A27 * 1.14 * 100;
                    double A31 = (A14 + A15) / A27 * 1.21 * 100;
                    double A32 = (A24 + A26) / (A24 + A25 + A26 + A27) * 100;
                    double A33 = (A23 + A25) / (A24 + A25 + A26 + A27) * 100;
                    double A34 = A16 / A7 * 0.93 * 100;
                    double A35 = A12 / A7 * 1.25 * 100;

                    txtA1.Text = Math.Round(A1, 2).ToString();
                    txtA2.Text = Math.Round(A2, 2).ToString();
                    txtA3.Text = Math.Round(A3, 2).ToString();
                    txtA4.Text = Math.Round(A4, 2).ToString();
                    txtA5.Text = Math.Round(A5, 2).ToString();
                    txtA6.Text = Math.Round(A6, 2).ToString();
                    txtA7.Text = Math.Round(A7, 2).ToString();
                    txtA8.Text = Math.Round(A8, 2).ToString();
                    txtA9.Text = Math.Round(A9, 2).ToString();
                    txtA10.Text = Math.Round(A10, 2).ToString();
                    txtA11.Text = Math.Round(A11, 2).ToString();
                    txtA12.Text = Math.Round(A12, 2).ToString();
                    txtA13.Text = Math.Round(A13, 2).ToString();
                    txtA14.Text = Math.Round(A14, 2).ToString();
                    txtA15.Text = Math.Round(A15, 2).ToString();
                    txtA16.Text = Math.Round(A16, 2).ToString();
                    txtA17.Text = Math.Round(A17, 2).ToString();
                    txtA18.Text = Math.Round(A18, 2).ToString();
                    txtA19.Text = Math.Round(A19, 2).ToString();
                    txtA20.Text = Math.Round(A20, 2).ToString();
                    txtA21.Text = Math.Round(A21, 2).ToString();
                    txtA22.Text = Math.Round(A22, 2).ToString();
                    txtA23.Text = Math.Round(A23, 2).ToString();
                    txtA24.Text = Math.Round(A24, 2).ToString();
                    txtA25.Text = Math.Round(A25, 2).ToString();
                    txtA26.Text = Math.Round(A26, 2).ToString();
                    txtA27.Text = Math.Round(A27, 2).ToString();
                    txtA28.Text = Math.Round(A28, 2).ToString();
                    txtA29.Text = Math.Round(A29, 2).ToString();
                    txtA30.Text = Math.Round(A30, 2).ToString();
                    txtA31.Text = Math.Round(A31, 2).ToString();
                    txtA32.Text = Math.Round(A32, 2).ToString();
                    txtA33.Text = Math.Round(A33, 2).ToString();
                    txtA34.Text = Math.Round(A34, 2).ToString();
                    txtA35.Text = Math.Round(A35, 2).ToString();

                    // define character of customer
                    string[] arCharacters = { cbbL1.Text, cbbL2.Text, cbbL3.Text, cbbL4.Text, cbbL5.Text, cbbR1.Text, cbbR2.Text, cbbR3.Text, cbbR4.Text, cbbR5.Text };
                    Tuple<string, string> characters = getCharacter(arCharacters);
                    //Console.WriteLine("Character: {0}, {1}", characters.Item1.ToString(), characters.Item2.ToString());
                    if(characters.Item1.ToString() == "chim_cong")
                    {
                        txtA36.Text = "Chim công";
                    }
                    else if (characters.Item1.ToString() == "dai_bang")
                    {
                        txtA36.Text = "Đại bàng";
                    }
                    else if (characters.Item1.ToString() == "chim_cu")
                    {
                        txtA36.Text = "Chim cú";
                    }
                    else if (characters.Item1.ToString() == "bo_cau")
                    {
                        txtA36.Text = "Bồ câu";
                    }
                    //----------------------------------
                    if (characters.Item2.ToString() == "chim_cong")
                    {
                        txtA37.Text = "Chim công";
                    }
                    else if (characters.Item2.ToString() == "dai_bang")
                    {
                        txtA37.Text = "Đại bàng";
                    }
                    else if (characters.Item2.ToString() == "chim_cu")
                    {
                        txtA37.Text = "Chim cú";
                    }
                    else if (characters.Item2.ToString() == "bo_cau")
                    {
                        txtA37.Text = "Bồ câu";
                    }

                    

                }    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Tính toán thông tin lỗi!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        //Creeate the Doc Method
        private void exportDocx(object filename)
        {

           
            try
            {
                // data.smk(đai bàng) , config.daa(cú) , user.rtk(công) , biotech.toy(bồ câu)

                object missing = Missing.Value;
                //object filename = Path.Combine(Directory.GetCurrentDirectory(), @"assets\", path.ToString());
                //Console.WriteLine(path);
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    g_wordApp.Visible = false;

                    g_myWordDoc = g_wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);
                    g_myWordDoc.Activate();


                    
                    // fill in customer's information
                    this.FindAndReplace(g_wordApp, "<B1>", txtFullName.Text);
                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    foreach (Microsoft.Office.Interop.Word.Section section in g_myWordDoc.Sections)
                    {
                        Microsoft.Office.Interop.Word.Range footerRange = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        footerRange.Find.Text = "<B1>";
                        footerRange.Find.Replacement.Text = txtFullName.Text;
                        footerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                    }
                    this.FindAndReplace(g_wordApp, "<B2>", txtBirthday.Text);
                    this.FindAndReplace(g_wordApp, "<B3>", txtEmail.Text);
                    this.FindAndReplace(g_wordApp, "<B4>", txtPhoneNumber.Text);
                    this.FindAndReplace(g_wordApp, "<B5>", txtAddress.Text);
                    this.FindAndReplace(g_wordApp, "<B6>", txtSex.Text);
                    //replace picture Profile
                    if (g_arListPathImages[0] != null)
                        foreach (Word.Range rng in g_myWordDoc.StoryRanges)
                        {
                            while (rng.Find.Execute("<pic_profile>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                            {
                                rng.Select();
                                g_wordApp.Selection.InlineShapes.AddPicture(g_arListPathImages[0], false, true);
                            }
                        }
                    //////////////////////////////////////////
                    
                    insertChart();
                   
                    
                    // replace biometric data
                    this.FindAndReplace(g_wordApp, "<A1>", txtA1.Text);
                    this.FindAndReplace(g_wordApp, "<A2>", txtA2.Text);
                    this.FindAndReplace(g_wordApp, "<A3>", txtA3.Text);
                    this.FindAndReplace(g_wordApp, "<A4>", txtA4.Text);
                    this.FindAndReplace(g_wordApp, "<A5>", txtA5.Text);
                    this.FindAndReplace(g_wordApp, "<A6>", txtA6.Text);
                    this.FindAndReplace(g_wordApp, "<A7>", txtA7.Text);
                    this.FindAndReplace(g_wordApp, "<A8>", txtA8.Text);
                    this.FindAndReplace(g_wordApp, "<A9>", txtA9.Text);

                    this.FindAndReplace(g_wordApp, "<A10>", txtA10.Text);
                    this.FindAndReplace(g_wordApp, "<A11>", txtA11.Text);
                    this.FindAndReplace(g_wordApp, "<A12>", txtA12.Text);
                    this.FindAndReplace(g_wordApp, "<A13>", txtA13.Text);
                    this.FindAndReplace(g_wordApp, "<A14>", txtA14.Text);
                    this.FindAndReplace(g_wordApp, "<A15>", txtA15.Text);
                    this.FindAndReplace(g_wordApp, "<A16>", txtA16.Text);
                    this.FindAndReplace(g_wordApp, "<A17>", txtA17.Text);
                    this.FindAndReplace(g_wordApp, "<A18>", txtA18.Text);
                    this.FindAndReplace(g_wordApp, "<A19>", txtA19.Text);

                    this.FindAndReplace(g_wordApp, "<A20>", txtA20.Text);
                    this.FindAndReplace(g_wordApp, "<A21>", txtA21.Text);
                    this.FindAndReplace(g_wordApp, "<A22>", txtA22.Text);
                    this.FindAndReplace(g_wordApp, "<A23>", txtA23.Text);
                    this.FindAndReplace(g_wordApp, "<A24>", txtA24.Text);
                    this.FindAndReplace(g_wordApp, "<A25>", txtA25.Text);
                    this.FindAndReplace(g_wordApp, "<A26>", txtA26.Text);
                    this.FindAndReplace(g_wordApp, "<A27>", txtA27.Text);
                    this.FindAndReplace(g_wordApp, "<A28>", txtA28.Text);
                    this.FindAndReplace(g_wordApp, "<A29>", txtA29.Text);

                    this.FindAndReplace(g_wordApp, "<A30>", txtA30.Text);
                    this.FindAndReplace(g_wordApp, "<A31>", txtA31.Text);
                    this.FindAndReplace(g_wordApp, "<A32>", txtA32.Text);
                    this.FindAndReplace(g_wordApp, "<A33>", txtA33.Text);
                    this.FindAndReplace(g_wordApp, "<A34>", txtA34.Text);
                    this.FindAndReplace(g_wordApp, "<A35>", txtA35.Text);
                    this.FindAndReplace(g_wordApp, "<A36>", txtA36.Text);
                    this.FindAndReplace(g_wordApp, "<A37>", txtA37.Text);

                    this.FindAndReplace(g_wordApp, "<C1>", cbbL1.Text);
                    this.FindAndReplace(g_wordApp, "<C2>", cbbL2.Text);
                    this.FindAndReplace(g_wordApp, "<C3>", cbbL3.Text);
                    this.FindAndReplace(g_wordApp, "<C4>", cbbL4.Text);
                    this.FindAndReplace(g_wordApp, "<C5>", cbbL5.Text);
                    this.FindAndReplace(g_wordApp, "<C6>", cbbR1.Text);
                    this.FindAndReplace(g_wordApp, "<C7>", cbbR2.Text);
                    this.FindAndReplace(g_wordApp, "<C8>", cbbR3.Text);
                    this.FindAndReplace(g_wordApp, "<C9>", cbbR4.Text);
                    this.FindAndReplace(g_wordApp, "<C10>", cbbR5.Text);


                    

                    float[] arAi = { float.Parse(txtA20.Text.ToString()), float.Parse(txtA21.Text.ToString()), float.Parse(txtA22.Text.ToString()) };
                    var myList = arAi.ToList();
                    myList.Sort((pair1, pair2) => pair1.CompareTo(pair2));
                    string[] arTextAi = { "Có thể", "Nên làm", "Tốt nhất" };
                    string A20i = "", A21i = "", A22i = "";

                    for (int i = 0; i < myList.Count; i++)
                        if (float.Parse(txtA20.Text.ToString()) == myList[i])
                        {
                            A20i = arTextAi[i];
                            myList[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList.Count; i++)
                        if (float.Parse(txtA21.Text.ToString()) == myList[i])
                        {
                            A21i = arTextAi[i];
                            myList[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList.Count; i++)
                        if (float.Parse(txtA22.Text.ToString()) == myList[i])
                        {
                            A22i = arTextAi[i];
                            myList[i] = -1;
                            break;
                        }

                    
                    this.FindAndReplace(g_wordApp, "<A20i>", A20i);
                    this.FindAndReplace(g_wordApp, "<A21i>", A21i);
                    this.FindAndReplace(g_wordApp, "<A22i>", A22i);
                    // Calculate A38-A41
                    string[] arCharacters_ = { cbbL1.Text, cbbL2.Text, cbbL3.Text, cbbL4.Text, cbbL5.Text, cbbR1.Text, cbbR2.Text, cbbR3.Text, cbbR4.Text, cbbR5.Text };
                    int countA38 = 0, countA39 = 0, countA40 = 0, countA41 = 0;
                    for (int i = 0; i < arCharacters_.Length; i++)
                    {
                        if (arCharacters_[i] == "WT" || arCharacters_[i] == "WS" || arCharacters_[i] == "WP" ||
                        arCharacters_[i] == "WE" || arCharacters_[i] == "WI" || arCharacters_[i] == "WD" || arCharacters_[i] == "WX")
                        {
                            countA38++;
                        }
                        else if (arCharacters_[i] == "AT" || arCharacters_[i] == "AS")
                        {
                            countA40++;
                        }
                        else if (arCharacters_[i] == "UL")
                        {
                            countA39++;
                        }
                        else if (arCharacters_[i] == "RL")
                        {
                            countA41++;
                        }
                    }
                    double A38 = countA38 * 1.0 / arCharacters_.Length * 100.0;
                    double A39 = countA39 * 1.0 / arCharacters_.Length * 100.0;
                    double A40 = countA40 * 1.0 / arCharacters_.Length * 100.0;
                    double A41 = 100.0 - A38 - A39 - A40;
                    this.FindAndReplace(g_wordApp, "<A38>", Math.Round(A38, 2).ToString());
                    this.FindAndReplace(g_wordApp, "<A39>", Math.Round(A39, 2).ToString());
                    this.FindAndReplace(g_wordApp, "<A40>", Math.Round(A40, 2).ToString());
                    this.FindAndReplace(g_wordApp, "<A41>", Math.Round(A41, 2).ToString());


                    // Calculate A10i -A19i
                    float[] arA10_A19 = {float.Parse(txtA10.Text.ToString()), float.Parse(txtA11.Text.ToString()), float.Parse(txtA12.Text.ToString()), float.Parse(txtA13.Text.ToString()), float.Parse(txtA14.Text.ToString()),
                                        float.Parse(txtA15.Text.ToString()), float.Parse(txtA16.Text.ToString()), float.Parse(txtA17.Text.ToString()), float.Parse(txtA18.Text.ToString()), float.Parse(txtA19.Text.ToString()) };
                    var myList_page35 = arA10_A19.ToList();
                    myList_page35.Sort((pair1, pair2) => pair1.CompareTo(pair2));

                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA10.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A10i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }

                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA11.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A11i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }

                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA12.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A12i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }

                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA13.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A13i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA14.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A14i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA15.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A15i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA16.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A16i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA17.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A17i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA18.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A18i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }
                    for (int i = 0; i < myList_page35.Count; i++)
                        if (float.Parse(txtA19.Text.ToString()) == myList_page35[i])
                        {

                            this.FindAndReplace(g_wordApp, "<A19i>", (i + 1).ToString());
                            myList_page35[i] = -1;
                            break;
                        }


                    // Replace fringerprint image to do41
                    for (int i = 1; i <= 10; i++)
                    {
                        if(g_arListPathImages[i] != null)
                        foreach (Word.Range rng in g_myWordDoc.StoryRanges)
                        {
                            string name_finger = "";
                            if (i < 6)
                            {
                                string pre_name_finger = "<Pic_L>";
                                string index_finger = (i).ToString();
                                name_finger = pre_name_finger.Insert(6, index_finger);
                            }
                            else
                            {
                                string pre_name_finger = "<Pic_R>";
                                string index_finger = (i - 5).ToString();
                                name_finger = pre_name_finger.Insert(6, index_finger);
                            }
                            while (rng.Find.Execute(name_finger, Forward: true, Wrap: WdFindWrap.wdFindContinue))
                            {
                                rng.Select();
                                g_wordApp.Selection.InlineShapes.AddPicture(g_arListPathImages[i], false, true);
                            }
                        }
                    }
                }
                else
                {
                    
                    MessageBox.Show("Không tìm thấy file Temp.docx!");
                }


                

                SaveFileDialog save_file = new SaveFileDialog();
                save_file.Filter = "Documents (*.docx)|*.docx";
                save_file.Title = "Save file";
                if (save_file.ShowDialog() == DialogResult.OK)
                    if (save_file.FileName != "")
                    {
                        lbExporting.Visible = false;
                        //System.IO.FileStream fs = (System.IO.FileStream)save_file.OpenFile();
                        object fs = Path.GetFullPath(save_file.FileName);
                        g_myWordDoc.SaveAs2(ref fs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);
                        MessageBox.Show("Lưu báo cáo thành công", "Thông báo");

                        g_myWordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            
                        g_wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        if (g_myWordDoc != null)
                        {
                            Marshal.FinalReleaseComObject(g_myWordDoc);
                            g_myWordDoc = null;
                        }
                        if (g_wordApp != null)
                        {
                            Marshal.FinalReleaseComObject(g_wordApp);
                            g_wordApp = null;
                        }

                    }   
                //}
                //else
                //{
                //    //g_myWordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                //    g_myWordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                //   // g_wordApp.Quit();
                //    g_wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                //    MessageBox.Show("Khong luu bao cao", "Thông báo");
                //}
            }
            catch(Exception ex)
            {

                lbExporting.Visible = false;
                if (g_myWordDoc != null)
                {
                    g_myWordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(g_myWordDoc);
                    g_myWordDoc = null;
                }
                if (g_wordApp != null)
                {
                    g_wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(g_wordApp);
                    g_wordApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                MessageBox.Show(ex.ToString(), "Xuất báo cáo lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
            
        }

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            lbExporting.Visible = false;

            if (g_flCalculated == true)
            {

                g_wordApp = new Word.Application();

                string file_data = "";

                if (txtA36.Text == "Chim công")
                {
                    file_data = g_pathReferenceData + @"\" + "user.rtk";
                }
                else if (txtA36.Text == "Bồ câu")
                {
                    file_data = g_pathReferenceData + @"\" + "biotech.toy";

                }
                else if (txtA36.Text == "Đại bàng")
                {
                    file_data = g_pathReferenceData + @"\" + "data.smk";

                }
                else if (txtA36.Text == "Chim cú")
                {
                    file_data = g_pathReferenceData + @"\" + "config.data";

                }
                lbExporting.Visible = true;
                exportDocx(file_data);
            }
            else
            {
                MessageBox.Show("Bạn chưa tính toán số liệu" , "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //formWait = new Waiting(exportDocx)
            //formWait.Show();
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void btnL1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[1] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[1]);
                    picL1.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnL2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[2] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[2]);
                    picL2.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnL3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[3] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[3]);
                    picL3.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnL4_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[4] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[4]);
                    picL4.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnL5_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[5] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[5]);
                    picL5.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnR1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[6] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[6]);
                    picR1.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnR2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[7] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[7]);
                    picR2.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnR3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[8] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[8]);
                    picR3.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnR4_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                //open_file_report.InitialDirectory;
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[9] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[9]);
                    picR4.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnR5_Click(object sender, EventArgs e)
        {
            try
            {
                 OpenFileDialog open_file_report = new OpenFileDialog();
                 open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg" ;
                 if (open_file_report.ShowDialog() == DialogResult.OK)
                    {
                        string fileName = open_file_report.FileName;
                        g_arListPathImages[10] = System.IO.Path.GetFullPath(fileName);
                        Image img = Image.FromFile(g_arListPathImages[10]);
                        picR5.Image = img;
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc file báo cáo tham chiếu lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnProfile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open_file_report = new OpenFileDialog();
                open_file_report.Filter = "Image Files (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg";
                if (open_file_report.ShowDialog() == DialogResult.OK)
                {
                    string fileName = open_file_report.FileName;
                    g_arListPathImages[0] = System.IO.Path.GetFullPath(fileName);
                    Image img = Image.FromFile(g_arListPathImages[0]);
                    picProfile.Image = img;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Đọc ảnh đại diện người lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            
            calculate();
            g_flCalculated = true;
            lbExporting.Visible = false;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            g_flCalculated = false;
            lbExporting.Visible = false;
            this.Refresh();
            Array.Clear(g_arListPathImages, 0, g_arListPathImages.Length);
            cbbL1.Text = "AT";
            cbbL2.Text = "AT";
            cbbL3.Text = "AT";
            cbbL4.Text = "AT";
            cbbL5.Text = "AT";
            cbbR1.Text = "AT";
            cbbR2.Text = "AT";
            cbbR3.Text = "AT";
            cbbR4.Text = "AT";
            cbbR5.Text = "AT";

            txtA1.Text = "";
            txtA2.Text = "";
            txtA3.Text = "";
            txtA4.Text = "";
            txtA5.Text = "";
            txtA6.Text = "";
            txtA7.Text = "";
            txtA8.Text = "";
            txtA9.Text = "";
            txtA10.Text = "";
            txtA11.Text = "";
            txtA12.Text = "";
            txtA13.Text = "";
            txtA14.Text = "";
            txtA15.Text = "";
            txtA16.Text = "";
            txtA17.Text = "";
            txtA18.Text = "";
            txtA19.Text = "";
            txtA20.Text = "";
            txtA21.Text = "";
            txtA22.Text = "";
            txtA23.Text = "";
            txtA24.Text = "";
            txtA25.Text = "";
            txtA26.Text = "";
            txtA27.Text = "";
            txtA28.Text = "";
            txtA29.Text = "";
            txtA30.Text = "";
            txtA31.Text = "";
            txtA32.Text = "";
            txtA33.Text = "";
            txtA34.Text = "";
            txtA35.Text = "";

            txtL1.Text = "0";
            txtL2.Text = "0";
            txtL3.Text = "0";
            txtL4.Text = "0";
            txtL5.Text = "0";
            txtR1.Text = "0";
            txtR2.Text = "0";
            txtR3.Text = "0";
            txtR4.Text = "0";
            txtR5.Text = "0";

            txtFullName.Text = "";
            txtBirthday.Text = "";
            txtEmail.Text = "";
            txtAddress.Text = "";
            txtSex.Text = "";
            txtPhoneNumber.Text = "";

            picProfile.Image = null;
            picL1.Image = null;
            picL2.Image = null;
            picL3.Image = null;
            picL4.Image = null;
            picL5.Image = null;
            picR1.Image = null;
            picR2.Image = null;
            picR3.Image = null;
            picR4.Image = null;
            picR5.Image = null;

        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {

            HelpWindows help = new HelpWindows();
            help.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutWindows about = new AboutWindows();
            about.Show();
        }







        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }



        
        string createPraphChart_Page28()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {
               
                xlWorkSheet.Cells[1, 1] = "Thùy Trước Trán";
                xlWorkSheet.Cells[1, 2] = txtA1.Text;


                xlWorkSheet.Cells[2, 1] = "Thùy Trán";
                xlWorkSheet.Cells[2, 2] = txtA2.Text;


                xlWorkSheet.Cells[3, 1] = "Thùy Đỉnh";
                xlWorkSheet.Cells[3, 2] = txtA3.Text;

                xlWorkSheet.Cells[4, 1] = "Thùy Thái Dương";
                xlWorkSheet.Cells[4, 2] = txtA4.Text;

                xlWorkSheet.Cells[5, 1] = "Thùy Chẩm";
                xlWorkSheet.Cells[5, 2] = txtA5.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 260);
                Excel.Chart chartPage = myChart.Chart;

                chartPage.PlotArea.Interior.Color = System.Drawing.Color.LightYellow;

                chartRange = xlWorkSheet.get_Range("A1", "B5");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 4; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }
                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_5vung_naobo.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_5vung_naobo.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createGraphChart_Page34()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {

                xlWorkSheet.Cells[2, 1] = "Quản lý";
                xlWorkSheet.Cells[2, 2] = txtA10.Text;

                xlWorkSheet.Cells[3, 1] = "Lãnh đạo";
                xlWorkSheet.Cells[3, 2] = txtA11.Text;

                xlWorkSheet.Cells[4, 1] = "Lập trình, LOGIC";
                xlWorkSheet.Cells[4, 2] = txtA12.Text;

                xlWorkSheet.Cells[5, 1] = "Tưởng tượng";
                xlWorkSheet.Cells[5, 2] = txtA13.Text;

                xlWorkSheet.Cells[6, 1] = "Vân động tinh";
                xlWorkSheet.Cells[6, 2] = txtA4.Text;

                xlWorkSheet.Cells[7, 1] = "Vận động thô";
                xlWorkSheet.Cells[7, 2] = txtA15.Text;

                xlWorkSheet.Cells[8, 1] = "Ngôn luận";
                xlWorkSheet.Cells[8, 2] = txtA16.Text;

                xlWorkSheet.Cells[9, 1] = "Thính giác";
                xlWorkSheet.Cells[9, 2] = txtA17.Text;

                xlWorkSheet.Cells[10, 1] = "Hình ảnh";
                xlWorkSheet.Cells[10, 2] = txtA18.Text;

                xlWorkSheet.Cells[11, 1] = "Đọc, quan sát";
                xlWorkSheet.Cells[11, 2] = txtA19.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 600);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B11");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 10; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_10chucnang.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_10chucnang.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createGraphChart_Page35()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {


                xlWorkSheet.Cells[2, 1] = "L1";
                xlWorkSheet.Cells[2, 2] = txtA10.Text;

                xlWorkSheet.Cells[3, 1] = "L2";
                xlWorkSheet.Cells[3, 2] = txtA11.Text;

                xlWorkSheet.Cells[4, 1] = "L3";
                xlWorkSheet.Cells[4, 2] = txtA12.Text;

                xlWorkSheet.Cells[5, 1] = "L4";
                xlWorkSheet.Cells[5, 2] = txtA13.Text;

                xlWorkSheet.Cells[6, 1] = "L5";
                xlWorkSheet.Cells[6, 2] = txtA14.Text;

                xlWorkSheet.Cells[7, 1] = "R1";
                xlWorkSheet.Cells[7, 2] = txtA15.Text;

                xlWorkSheet.Cells[8, 1] = "R2";
                xlWorkSheet.Cells[8, 2] = txtA16.Text;

                xlWorkSheet.Cells[9, 1] = "R3";
                xlWorkSheet.Cells[9, 2] = txtA17.Text;

                xlWorkSheet.Cells[10, 1] = "R4";
                xlWorkSheet.Cells[10, 2] = txtA18.Text;

                xlWorkSheet.Cells[11, 1] = "R5";
                xlWorkSheet.Cells[11, 2] = txtA19.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 480, 120);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B11");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 10; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_bieudo_chucnang.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_bieudo_chucnang.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createPieChart_Page41()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {

                double A1 = Convert.ToDouble(txtA1.Text.ToString());
                double A2 = Convert.ToDouble(txtA2.Text.ToString());

                double A1_percent = Math.Round(A1 / (A1 + A2) * 100, 2);
                double A2_percnet = 100.0 - A1_percent;

                //add data 
                xlWorkSheet.Cells[1, 1] = "NĂNG ĐỘNG";
                xlWorkSheet.Cells[1, 2] = A1_percent.ToString() + "%";

                xlWorkSheet.Cells[2, 1] = "PHÂN TÍCH";
                xlWorkSheet.Cells[2, 2] = A2_percnet.ToString() + "%";

                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(50, 50, 150, 150);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A1", "B2");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.ChartType = Excel.XlChartType.xlPie;


                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_thuytruoc_thuytran.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "pie_page_41.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;

        }
        string createGraphChart_Page43()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {


                xlWorkSheet.Cells[2, 1] = "ÂM VƯC (L4)";
                xlWorkSheet.Cells[2, 2] = txtA16.Text;

                xlWorkSheet.Cells[3, 1] = "NGÔN NGỮ (R4)";
                xlWorkSheet.Cells[3, 2] = txtA17.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 150, 100);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B3");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 2; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_l4_r4.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_l4_r4.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createGraphChart_Page44()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {



                xlWorkSheet.Cells[2, 1] = "HÌNH ẢNH (L5)";
                xlWorkSheet.Cells[2, 2] = txtA18.Text;

                xlWorkSheet.Cells[3, 1] = "QUAN SÁT (R5)";
                xlWorkSheet.Cells[3, 2] = txtA19.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 150, 100);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B3");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 2; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_l5_r5.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_l5_r5.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }

        string createGraphChart_Page45()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {



                xlWorkSheet.Cells[2, 1] = "VẬN ĐỘNG THÔ (L3)";
                xlWorkSheet.Cells[2, 2] = txtA14.Text;

                xlWorkSheet.Cells[3, 1] = "VẬN ĐỘNG TINH (R3)";
                xlWorkSheet.Cells[3, 2] = txtA15.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 150, 100);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B3");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 2; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_l3_r3.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_l3_r3.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }


        string createGraphChart_Page48()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {

                xlWorkSheet.Cells[2, 1] = "Logic";
                xlWorkSheet.Cells[2, 2] = txtA28.Text;

                xlWorkSheet.Cells[3, 1] = "Ngôn ngữ";
                xlWorkSheet.Cells[3, 2] = txtA29.Text;

                xlWorkSheet.Cells[4, 1] = "Hướng nội";
                xlWorkSheet.Cells[4, 2] = txtA30.Text;

                xlWorkSheet.Cells[5, 1] = "Hướng ngoại";
                xlWorkSheet.Cells[5, 2] = txtA31.Text;

                xlWorkSheet.Cells[6, 1] = "Vận động";
                xlWorkSheet.Cells[6, 2] = txtA32.Text;

                xlWorkSheet.Cells[7, 1] = "Thị giác";
                xlWorkSheet.Cells[7, 2] = txtA33.Text;

                xlWorkSheet.Cells[8, 1] = "Thiên nhiên";
                xlWorkSheet.Cells[8, 2] = txtA34.Text;

                xlWorkSheet.Cells[9, 1] = "Âm nhạc";
                xlWorkSheet.Cells[9, 2] = txtA35.Text;
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 400, 480);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B9");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 8; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }
                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_dathongminh.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_dathongminh.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }


        string createGraphChart_Page57()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {



                xlWorkSheet.Cells[2, 1] = "IQ";
                xlWorkSheet.Cells[2, 2] = txtA23.Text;

                xlWorkSheet.Cells[3, 1] = "EQ";
                xlWorkSheet.Cells[3, 2] = txtA24.Text;

                xlWorkSheet.Cells[4, 1] = "AQ";
                xlWorkSheet.Cells[4, 2] = txtA25.Text;

                xlWorkSheet.Cells[5, 1] = "CQ";
                xlWorkSheet.Cells[5, 2] = txtA26.Text;


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 300, 180);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B5");
                chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 4; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_4q.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_4q.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }


        string createGraphChart_Page63()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {


                Random r = new Random();


                xlWorkSheet.Cells[2, 1] = "Kỹ năng giao tiếp";
                xlWorkSheet.Cells[2, 2] = Math.Round((float.Parse(txtA16.Text.ToString()) + float.Parse(txtA19.Text.ToString())) * 0.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[3, 1] = "Kỹ năng phân tich";
                xlWorkSheet.Cells[3, 2] = Math.Round((float.Parse(txtA12.Text.ToString()) * 2 + float.Parse(txtA16.Text.ToString())) * 0.7, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[4, 1] = "Khả năng quyết đoán";
                xlWorkSheet.Cells[4, 2] = Math.Round((float.Parse(txtA18.Text.ToString()) + float.Parse(txtA19.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[5, 1] = "Kỹ năng làm việc nhóm";
                xlWorkSheet.Cells[5, 2] = Math.Round((float.Parse(txtA16.Text.ToString()) + float.Parse(txtA14.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[6, 1] = "Kỹ năng lãnh đạo";
                xlWorkSheet.Cells[6, 2] = Math.Round(float.Parse(txtA19.Text.ToString()) * 2.3, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[7, 1] = "Lập kế hoạch chiến lược";
                xlWorkSheet.Cells[7, 2] = Math.Round((float.Parse(txtA10.Text.ToString()) + float.Parse(txtA11.Text.ToString())) * 1.05, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[8, 1] = "Khả năng tập trung";
                xlWorkSheet.Cells[8, 2] = Math.Round((float.Parse(txtA12.Text.ToString()) + float.Parse(txtA19.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[9, 1] = "Cách tiếp cận sáng tạo";
                xlWorkSheet.Cells[9, 2] = Math.Round((float.Parse(txtA13.Text.ToString()) + float.Parse(txtA18.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[10, 1] = "Khả năng quan sát";
                xlWorkSheet.Cells[10, 2] = Math.Round(float.Parse(txtA18.Text.ToString()) * 1.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[11, 1] = "Tuân thủ chất lượng";
                xlWorkSheet.Cells[11, 2] = Math.Round(float.Parse(txtA12.Text.ToString()) * 2.05, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[12, 1] = "Quản lý khủng hoảng";
                xlWorkSheet.Cells[12, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 2.1, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[13, 1] = "Thiết lập mục tiêu";
                xlWorkSheet.Cells[13, 2] = Math.Round((float.Parse(txtA12.Text.ToString()) + float.Parse(txtA13.Text.ToString())), MidpointRounding.AwayFromZero);
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 600);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B13");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                //chartPage.DataTable.Font.Size = 5;
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 12; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_kynang_quanly.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_kynang_quanly.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }

        string createGraphChart_Page64()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {


                Random r = new Random();


                xlWorkSheet.Cells[2, 1] = "Pháp luật";
                xlWorkSheet.Cells[2, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 0.5, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[3, 1] = "Báo chí";
                xlWorkSheet.Cells[3, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 0.75, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[4, 1] = "Tài chính";
                xlWorkSheet.Cells[4, 2] = Math.Round(float.Parse(txtA11.Text.ToString()) * 1.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[5, 1] = "Thương mại";
                xlWorkSheet.Cells[5, 2] = Math.Round(float.Parse(txtA11.Text.ToString()) * 1.7, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[6, 1] = "Máy tính";
                xlWorkSheet.Cells[6, 2] = Math.Round(float.Parse(txtA13.Text.ToString()) * 1.5, MidpointRounding.AwayFromZero) ;

                xlWorkSheet.Cells[7, 1] = "Kinh tế học";
                xlWorkSheet.Cells[7, 2] = Math.Round(float.Parse(txtA19.Text.ToString()) * 1.4, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[8, 1] = "Toán học";
                xlWorkSheet.Cells[8, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 1.3, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[9, 1] = "Vật lý";
                xlWorkSheet.Cells[9, 2] = Math.Round(float.Parse(txtA13.Text.ToString()) * 1.2, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[10, 1] = "Hoá học";
                xlWorkSheet.Cells[10, 2] = Math.Round(float.Parse(txtA13.Text.ToString()) * 1.05, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[11, 1] = "Sinh học";
                xlWorkSheet.Cells[11, 2] = Math.Round(float.Parse(txtA18.Text.ToString()) * 0.8, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[12, 1] = "Địa lý";
                xlWorkSheet.Cells[12, 2] = Math.Round(float.Parse(txtA18.Text.ToString()) * 0.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[13, 1] = "Quản lý";
                xlWorkSheet.Cells[13, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 0.95, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[14, 1] = "Lịch sử";
                xlWorkSheet.Cells[14, 2] = Math.Round(float.Parse(txtA19.Text.ToString()) * 0.95, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[15, 1] = "Tiếng Anh";
                xlWorkSheet.Cells[15, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 1.1, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[16, 1] = "Tiếng Pháp";
                xlWorkSheet.Cells[16, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 1.1, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[17, 1] = "Tiếng Việt";
                xlWorkSheet.Cells[17, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 1.5, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[18, 1] = "Nghệ thuật";
                xlWorkSheet.Cells[18, 2] = Math.Round(float.Parse(txtA13.Text.ToString()) * 1.1, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[19, 1] = "Kế toán";
                xlWorkSheet.Cells[19, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 0.88, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[20, 1] = "Đối ngoại";
                xlWorkSheet.Cells[20, 2] = Math.Round(float.Parse(txtA18.Text.ToString()) * 0.95, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[21, 1] = "Nhân văn";
                xlWorkSheet.Cells[21, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 0.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[22, 1] = "Giáo dục công dân";
                xlWorkSheet.Cells[22, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 0.7, MidpointRounding.AwayFromZero);

                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 600);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B22");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 21; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_monhoc.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_monhoc.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createGraphChart_Page65()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {

                Random r = new Random();


                xlWorkSheet.Cells[2, 1] = "Quản lý";
                xlWorkSheet.Cells[2, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 2.8, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[3, 1] = "Marketing & Bán hàng";
                xlWorkSheet.Cells[3, 2] = Math.Round((float.Parse(txtA10.Text.ToString()) + float.Parse(txtA12.Text.ToString()) + float.Parse(txtA19.Text.ToString())) * 0.7, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[4, 1] = "Tài chính";
                xlWorkSheet.Cells[4, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) + float.Parse(txtA12.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[5, 1] = "Nhân sự";
                xlWorkSheet.Cells[5, 2] = Math.Round(float.Parse(txtA11.Text.ToString()) + float.Parse(txtA18.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[6, 1] = "Lập kế hoạch";
                xlWorkSheet.Cells[6, 2] = Math.Round((float.Parse(txtA13.Text.ToString()) + float.Parse(txtA18.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[7, 1] = "Hoạt động";
                xlWorkSheet.Cells[7, 2] = Math.Round((float.Parse(txtA14.Text.ToString()) + float.Parse(txtA15.Text.ToString())) * 1.2, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[8, 1] = "Luật pháp";
                xlWorkSheet.Cells[8, 2] = Math.Round(float.Parse(txtA10.Text.ToString()) * 1.8, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[9, 1] = "Quản trị";
                xlWorkSheet.Cells[9, 2] = Math.Round((float.Parse(txtA10.Text.ToString()) + float.Parse(txtA11.Text.ToString())), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[10, 1] = "Nghiên cứu & Phát triển";
                xlWorkSheet.Cells[10, 2] = Math.Round((float.Parse(txtA14.Text.ToString()) + float.Parse(txtA16.Text.ToString())) * 1.28, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[11, 1] = "Chế tạo";
                xlWorkSheet.Cells[11, 2] = Math.Round((float.Parse(txtA12.Text.ToString()) + float.Parse(txtA13.Text.ToString())) * 1.15, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[12, 1] = "Mua sắm";
                xlWorkSheet.Cells[12, 2] = Math.Round(float.Parse(txtA17.Text.ToString()) + float.Parse(txtA18.Text.ToString()), MidpointRounding.AwayFromZero);


                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 600);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B12");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                chartPage.Legend.Delete();
                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();
                for (int i = 1; i <= 11; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));

                }
                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_nghenghiep.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                //chartPage.CopyPicture();
                //xlWorkSheet.Paste();
                ////This image has decent resolution
                //xlWorkSheet.Shapes.Item(xlWorkSheet.Shapes.Count).Copy();
                ////Save the image
                //System.Windows.Media.Imaging.BitmapEncoder enc = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                //enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                //pathPage = g_pathReferenceData + @"\" + "chart_nghenghiep.jpg";
                //using (System.IO.MemoryStream outStream = new System.IO.MemoryStream())
                //{
                //    enc.Save(outStream);
                //    System.Drawing.Image pic = new System.Drawing.Bitmap(outStream);
                //    pic.Save(pathPage);
                //}

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }

            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        string createGraphChart_Page66()
        {
            string pathPage = "";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            try
            {
                Random r = new Random();

                xlWorkSheet.Cells[2, 1] = "Viết nhật ký";
                xlWorkSheet.Cells[2, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 0.3, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[3, 1] = "Thể dục thẩm mỹ";
                xlWorkSheet.Cells[3, 2] = Math.Round(float.Parse(txtA14.Text.ToString()) * 1.05, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[4, 1] = "Thư pháp";
                xlWorkSheet.Cells[4, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) + float.Parse(txtA18.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[5, 1] = "Cờ vua";
                xlWorkSheet.Cells[5, 2] = Math.Round(float.Parse(txtA11.Text.ToString()) + float.Parse(txtA12.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[6, 1] = "Máy tính";
                xlWorkSheet.Cells[6, 2] = Math.Round(float.Parse(txtA12.Text.ToString()) + float.Parse(txtA13.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[7, 1] = "Nấu ăn";
                xlWorkSheet.Cells[7, 2] = Math.Round(float.Parse(txtA14.Text.ToString()) + float.Parse(txtA13.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[8, 1] = "Nhảy";
                xlWorkSheet.Cells[8, 2] = Math.Round(float.Parse(txtA14.Text.ToString()) * 1.4, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[9, 1] = "Tranh luận";
                xlWorkSheet.Cells[9, 2] = Math.Round(float.Parse(txtA16.Text.ToString()) * 0.5, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[10, 1] = "Sân khấu";
                xlWorkSheet.Cells[10, 2] = Math.Round(float.Parse(txtA14.Text.ToString()) * 0.7, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[11, 1] = "Vẽ / hội hoạ";
                xlWorkSheet.Cells[11, 2] = Math.Round(float.Parse(txtA19.Text.ToString()) + float.Parse(txtA13.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[12, 1] = "Âm nhạc";
                xlWorkSheet.Cells[12, 2] = Math.Round(float.Parse(txtA17.Text.ToString()) * 1.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[13, 1] = "Nhiếp ảnh";
                xlWorkSheet.Cells[13, 2] = Math.Round(float.Parse(txtA18.Text.ToString()) * 0.9, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[14, 1] = "Trò chơi ngoài trời";
                xlWorkSheet.Cells[14, 2] = Math.Round(float.Parse(txtA15.Text.ToString()) + float.Parse(txtA14.Text.ToString()), MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[15, 1] = "Đi bộ dài";
                xlWorkSheet.Cells[15, 2] = Math.Round(float.Parse(txtA15.Text.ToString()) * 1.5, MidpointRounding.AwayFromZero);

                xlWorkSheet.Cells[16, 1] = "Đạp xe";
                xlWorkSheet.Cells[16, 2] = Math.Round(float.Parse(txtA14.Text.ToString()) * 1.2, MidpointRounding.AwayFromZero);




                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(20, 20, 450, 600);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = xlWorkSheet.get_Range("A2", "B16");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlBarClustered; // page 34
                
                chartPage.Legend.Delete();
                var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

                chartPage.ApplyDataLabels(Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue);
                // set color 
                Excel.Series series1 = (Excel.Series)chartPage.SeriesCollection(1);
                series1.HasDataLabels = true;
                Random rColor = new Random();    
                for(int i = 1; i <=15 ; i++)
                {
                    Excel.Point pti = series1.Points(i);
                    pti.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(rColor.Next(5, 256), rColor.Next(5, 256), rColor.Next(5, 256)));
                   
                }

                // export chart to image
                pathPage = g_pathReferenceData + @"\" + "chart_hoatdong.png";
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    Excel.ChartObjects chartObjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                    foreach (Excel.ChartObject co in chartObjects)
                    {
                        Excel.Chart chart = (Excel.Chart)co.Chart;
                        //                  app.Goto(co, true);
                        chart.Export(pathPage, "PNG", false);
                    }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false,  misValue,  misValue);
                xlApp.Quit(); 
                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                
            }
            catch (Exception ex)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                if (xlWorkSheet != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }
                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }
                if (xlApp != null)
                {

                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                MessageBox.Show(ex.ToString(), "Tạo biểu đồ lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return pathPage;
        }
        private void insertChart()
        {

            string pathImagePage28 = createPraphChart_Page28();
            string pathImagePage34 = createGraphChart_Page34();
            string pathImagePage35 = createGraphChart_Page35();
            string pathImagePage41 = createPieChart_Page41();
            string pathImagePage43 = createGraphChart_Page43();
            string pathImagePage44 = createGraphChart_Page44();
            string pathImagePage45 = createGraphChart_Page45();
            string pathImagePage48 = createGraphChart_Page48();
            string pathImagePage57 = createGraphChart_Page57();
            string pathImagePage63 = createGraphChart_Page63();
            string pathImagePage64 = createGraphChart_Page64();
            string pathImagePage65 = createGraphChart_Page65();
            string pathImagePage66 = createGraphChart_Page66();
            if (pathImagePage28!= null)
                foreach (Word.Range rng in g_myWordDoc.StoryRanges)
                {
                    while (rng.Find.Execute("<chart_5vung_naobo>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage28, false, true);
                    }
                    while (rng.Find.Execute("<chart_10chucnang>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage34, false, true);
                    }
                    while (rng.Find.Execute("<chart_bieudo_chucnang>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage35, false, true);

                    }
                    while (rng.Find.Execute("<chart_thuytruoc_thuytran>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage41, false, true);

                    }
                    while (rng.Find.Execute("<chart_l4_r4>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage43, false, true);

                    }
                    while (rng.Find.Execute("<chart_l5_r5>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage44, false, true);

                    }
                    while (rng.Find.Execute("<chart_l3_r3>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage45, false, true);

                    }
                    while (rng.Find.Execute("<chart_dathongminh>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage48, false, true);

                    }
                    while (rng.Find.Execute("<chart_4q>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage57, false, true);

                    }
                    while (rng.Find.Execute("<chart_kynang_quanly>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage63, false, true);

                    }
                    while (rng.Find.Execute("<chart_monhoc>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage64, false, true);

                    }
                    while (rng.Find.Execute("<chart_nghenghiep>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage65, false, true);

                    }
                    while (rng.Find.Execute("<chart_hoatdong>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
                    {
                        rng.Select();
                        g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage66, false, true);

                    }
                }
            ////
            //string pathImagePage34 = createGraphChart_Page34();
            //if (pathImagePage34 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_10chucnang>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage34, false, true);
            //        }
            //    }
            ////
            //string pathImagePage35 = createGraphChart_Page35();
            //if (pathImagePage35 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_bieudo_chucnang>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage35, false, true);

            //        }
            //    }
            ////
            //string pathImagePage41 = createPieChart_Page41();
            //if (pathImagePage41 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_thuytruoc_thuytran>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage41, false, true);

            //        }
            //    }
            ////
            //string pathImagePage43 = createGraphChart_Page43();
            //if (pathImagePage43 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_l4_r4>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage43, false, true);

            //        }
            //    }

            //string pathImagePage44 = createGraphChart_Page44();
            //if (pathImagePage44 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_l5_r5>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage44, false, true);

            //        }
            //    }

            //string pathImagePage45 = createGraphChart_Page45();
            //if (pathImagePage44 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_l3_r3>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage45, false, true);

            //        }
            //    }

            //string pathImagePage48 = createGraphChart_Page48();
            //if (pathImagePage48 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_dathongminh>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage48, false, true);

            //        }
            //    }

            //string pathImagePage57 = createGraphChart_Page57();
            //if (pathImagePage57 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_4q>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage57, false, true);

            //        }
            //    }

            //string pathImagePage63 = createGraphChart_Page63();
            //if (pathImagePage63 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_kynang_quanly>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage63, false, true);

            //        }
            //    }
            //string pathImagePage64 = createGraphChart_Page64();
            //if (pathImagePage64 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_monhoc>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage64, false, true);

            //        }
            //    }
            //string pathImagePage65 = createGraphChart_Page65();
            //if (pathImagePage65 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_nghenghiep>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage65, false, true);

            //        }
            //    }
            //string pathImagePage66 = createGraphChart_Page66();
            //if (pathImagePage66 != null)
            //    foreach (Word.Range rng in g_myWordDoc.StoryRanges)
            //    {
            //        while (rng.Find.Execute("<chart_hoatdong>", Forward: true, Wrap: WdFindWrap.wdFindContinue))
            //        {
            //            rng.Select();
            //            g_wordApp.Selection.InlineShapes.AddPicture(pathImagePage66, false, true);

            //        }
            //    }

            if (File.Exists(pathImagePage28))
            {
                File.Delete(pathImagePage28);
            }
            if (File.Exists(pathImagePage34))
            {
                File.Delete(pathImagePage34);
            }
            if (File.Exists(pathImagePage35))
            {
                File.Delete(pathImagePage35);
            }
            if (File.Exists(pathImagePage41))
            {
                File.Delete(pathImagePage41);
            }
            if (File.Exists(pathImagePage43))
            {
                File.Delete(pathImagePage43);
            }
            if (File.Exists(pathImagePage44))
            {
                File.Delete(pathImagePage44);
            }
            if (File.Exists(pathImagePage45))
            {
                File.Delete(pathImagePage45);
            }
            if (File.Exists(pathImagePage48))
            {
                File.Delete(pathImagePage48);
            }
            if (File.Exists(pathImagePage57))
            {
                File.Delete(pathImagePage57);
            }
            if (File.Exists(pathImagePage63))
            {
                File.Delete(pathImagePage63);
            }
            if (File.Exists(pathImagePage64))
            {
                File.Delete(pathImagePage64);
            }
            if (File.Exists(pathImagePage65))
            {
                File.Delete(pathImagePage65);
            }
            if (File.Exists(pathImagePage66))
            {
                File.Delete(pathImagePage66);
            }

        }
        static UnmanagedMemoryStream GetResourceStream(string resName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var strResources = assembly.GetName().Name + ".g.resources";
            var rStream = assembly.GetManifestResourceStream(strResources);
            var resourceReader = new System.Resources.ResourceReader(rStream);
            var items = resourceReader.OfType<System.Collections.DictionaryEntry>();
            var stream = items.First(x => (x.Key as string) == resName.ToLower()).Value;
            return (UnmanagedMemoryStream)stream;
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            MessageBox.Show(path.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
    }
}
