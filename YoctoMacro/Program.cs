using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using MailKit.Net.Imap;
using MailKit.Search;


namespace YoctoMacro
{
    class Program
    {
        static YAnButton macro1, macro2, macro3, macro4, macro5,
            macro6, macro7, macro8, macro9;
        static YAnButton mute;
        static YColorLedCluster leds;
        static YQuadratureDecoder quadraVolume;
        static YModule maxiKnob;
        static string press = "0";
        static string errmsg = "";
        static double prevVal = 0;
        static string username = "your_mail";
        static string password = "your_password";
        static bool newMail = false, seqOn = false, newCommande = false, seqCommande = false;
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]

        public static extern void keybd_event(uint bVk, uint bScan, uint dwFlags, uint dwExtraInfo);
        static void BtnPress(YFunction fct, string text)
        {
            var uri= "";
            var psi = new System.Diagnostics.ProcessStartInfo();
            if (text != "0")
            {
                press = fct.FriendlyName;
                string[] splitval = press.Split('.');
                press = splitval[1];

                switch (press)
                {
                    case "prev": //Prev Song
                        keybd_event(0xB1, 0, 0, 0);
                        leds.set_hslColor(3, 1, 0x000000);
                        leds.hsl_move(3, 1, 0xBCFF2A, 500);
                        YAPI.Sleep(100, ref errmsg);
                        leds.set_hslColor(2, 1, 0x000000);
                        leds.hsl_move(2, 1, 0xBCFF2A, 500);
                        YAPI.Sleep(100, ref errmsg);
                        leds.set_hslColor(1, 1, 0x000000);
                        leds.hsl_move(1, 1, 0xBCFF2A, 500);
                        break;
                    case "play": //Play Pause Song
                        keybd_event(0xB3, 0, 0, 0);
                        leds.set_hslColor(2, 1, 0x000000);
                        leds.hsl_move(2, 1, 0xBCFF2A, 500);
                        YAPI.Sleep(100, ref errmsg);
                        leds.set_hslColor(3, 1, 0x000000);
                        leds.hsl_move(3, 1, 0xBCFF2A, 500);
                        leds.set_hslColor(1, 1, 0x000000);
                        leds.hsl_move(1, 1, 0xBCFF2A, 500);
                        break;
                    case "next": //Next Song
                        keybd_event(0xB0, 0, 0, 0);
                        leds.set_hslColor(1,1,0x000000);
                        leds.hsl_move(1, 1, 0xBCFF2A, 500);
                        YAPI.Sleep(100, ref errmsg);
                        leds.set_hslColor(2, 1, 0x000000);
                        leds.hsl_move(2, 1, 0xBCFF2A, 500);
                        YAPI.Sleep(100, ref errmsg);
                        leds.set_hslColor(3, 1, 0x000000);
                        leds.hsl_move(3, 1, 0xBCFF2A, 500);
                        break;
                    case "1": //Open Postbox
                        Process.Start("C:\\Program Files (x86)\\Postbox\\postbox.exe");
                        break;
                    case "2": //Open Spotify
                        Process.Start("C:\\Users\\yocto\\AppData\\Local\\Microsoft\\WindowsApps\\Spotify.exe");
                        break;
                    case "3": //Open Firefox
                        Process.Start("C:\\Program Files\\Mozilla Firefox\\firefox.exe");
                        break;
                    case "4": //Open Yoctopuce.com
                        uri = "https://172.17.17.77/FR/admin/project_list.php";
                        psi.UseShellExecute = true;
                        psi.FileName = uri;
                        Process.Start(psi);
                        break;
                    case "5": //Open virtualHub Config
                        Process[] proc = Process.GetProcessesByName("yoctoFlashSrvV3");
                        if (proc.Length >= 1)
                        {
                            proc[0].Kill();
                        }
                        Process.Start("C:\\Program Files (x86)\\Yoctopuce\\VirtualHub\\vhubconfig.exe");
                        break;
                    case "6":
                        Process[] processe;
                        processe = Process.GetProcessesByName("yoctoFlashSrvV3");
                        if (processe.Length == 0)
                        {
                            processe = Process.GetProcessesByName("VirtualHub");
                            if (processe.Length >= 1)
                            {
                                processe[0].Kill();
                                processe = Process.GetProcessesByName("vhubconfig");
                                if (processe.Length >= 1)
                                {
                                    processe[0].Kill();
                                }
                            }

                            Process.Start(@"D:\yoctopuce\yoctoprod\tools\yoctoFlashSrvV3\yoctoFlashSrvV3.exe");
                        }                                              
                        break;
                    case "mute": //Mute Volume
                        keybd_event(0xAD, 0, 0, 0); 
                        break;
                }
            }
        }
        static void QuadraChange(YFunction fct, string text)
        {
            
            double val = quadraVolume.get_currentValue();

            if (prevVal < val)
            {
                keybd_event(0xAE, 0, 0, 0); //Volume Down
            }
            else if(prevVal > val)
            {
                keybd_event(0xAF, 0, 0, 0); //Volume Up
            }
            prevVal = val;            
        } 
        
        static void LEDAnime()
        {
            Process[] process;
            if(newMail)
            { 
                if (!seqOn)
                {
                    leds.resetBlinkSeq(0);
                    leds.addHslMoveToBlinkSeq(0, 0xBCFFB0, 500);
                    leds.addHslMoveToBlinkSeq(0, 0xBCFFB0, 250);
                    leds.addHslMoveToBlinkSeq(0, 0xBCFF2A, 500);
                    leds.addHslMoveToBlinkSeq(0, 0xBCFF2A, 750);
                    leds.linkLedToBlinkSeq(4, 1, 0, 0);
                    leds.startBlinkSeq(0);
                    seqOn = true;
                }  
            }
            else
            {
                leds.stopBlinkSeq(0);
                seqOn = false;
                leds.hsl_move(4, 1, 0xBCFF2A, 500);
            }
            
            process = Process.GetProcessesByName("Spotify");
            if (process.Length >= 1)
            {
                leds.hsl_move(5, 1, 0x57FF10, 500);
            }
            else
            {
                leds.hsl_move(5, 1, 0xBCFF2A, 500);
            }
            if (newCommande)
            {
                if (!seqCommande)
                {
                    leds.resetBlinkSeq(1);
                    leds.addHslMoveToBlinkSeq(1, 0xFFFF2A, 500);
                    leds.addHslMoveToBlinkSeq(1, 0xFFFF2A, 500);
                    leds.addHslMoveToBlinkSeq(1, 0xBCFF2A, 500);
                    leds.addHslMoveToBlinkSeq(1, 0xBCFF2A, 500);
                    leds.linkLedToBlinkSeq(7, 1, 1, 0);
                    leds.startBlinkSeq(1);
                    seqCommande = true;
                }
            }
            else
            {
                leds.stopBlinkSeq(1);
                seqCommande = false;
                leds.hsl_move(7, 1, 0xBCFF2A, 500);
            }
            process = Process.GetProcessesByName("vhubconfig");
            if (process.Length >= 1)
            {
                leds.hsl_move(8, 1, 0xA0FF7A, 500);
            }
            else
            {
                leds.hsl_move(8, 1, 0xBCFF2A, 500);
            }

            process = Process.GetProcessesByName("yoctoFlashSrvV3");
            if (process.Length >= 1)
            {
                leds.hsl_move(9, 1, 0xBC0061, 500);
            }
            else
            {
                leds.hsl_move(9, 1, 0xBCFF2A, 500);
            }
            
        }
        
        static void NewMail()
        {
            int nbMail = 0;
            int nbCommande = 0;
            using var client = new ImapClient();
            bool result = false;
            client.Connect("mail.infomaniak.com", 143, false);
            client.Authenticate(username, password);

            var inbox = client.Inbox;
            
            inbox.Open(MailKit.FolderAccess.ReadOnly);
            var query = SearchQuery.NotSeen.And(SearchQuery.SubjectContains("New invoice payment"));
            foreach (var uid in inbox.Search(query))
            {
                nbCommande++;
            }
            inbox.Open(MailKit.FolderAccess.ReadOnly);
            foreach (var uid in inbox.Search(SearchQuery.NotSeen))
            {
                nbMail++;
            }
            client.Disconnect(true);
            if (nbCommande >= 1)
            {
                if (!newCommande)
                {
                    newCommande = true;
                    seqCommande = false;
                }
            }
            else{ newCommande = false; }
            if (nbMail >=1 && nbCommande!=nbMail)
            {
                if (!newMail)
                {
                    newMail = true;
                    seqOn = false;
                }
            }
            else { newMail = false;}
           
        }
        static void Main(string[] args)
        {
            
            if (YAPI.RegisterHub("usb", ref errmsg) != YAPI.SUCCESS)
            {
                if (YAPI.RegisterHub("127.0.0.1", ref errmsg) != YAPI.SUCCESS)
                {
                        Console.WriteLine("RegisterHub error: " + errmsg);
                        Environment.Exit(0);
                }
             }

            #region Init
            maxiKnob = YModule.FirstModule();
            macro1 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton1");
            macro2 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton2");
            macro3 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton3");
            macro4 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton4");
            macro5 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton5");
            macro6 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton6");
            macro7 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton7");
            macro8 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton8");
            macro9 = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton9");
            mute = YAnButton.FindAnButton(maxiKnob.FriendlyName + ".anButton10");
            leds = YColorLedCluster.FindColorLedCluster(maxiKnob.FriendlyName + ".colorLedCluster");
            quadraVolume = YQuadratureDecoder.FindQuadratureDecoder(maxiKnob.FriendlyName + ".quadratureDecoder1");
            prevVal = quadraVolume.get_currentValue();
            macro1.registerValueCallback(BtnPress);
            macro2.registerValueCallback(BtnPress);
            macro3.registerValueCallback(BtnPress);
            macro4.registerValueCallback(BtnPress);
            macro5.registerValueCallback(BtnPress);
            macro6.registerValueCallback(BtnPress);
            macro7.registerValueCallback(BtnPress);
            macro8.registerValueCallback(BtnPress);
            macro9.registerValueCallback(BtnPress);
            mute.registerValueCallback(BtnPress);
            quadraVolume.registerValueCallback(QuadraChange);
            leds.set_hslColor(1, 9, 0xBCFF2A);
            #endregion
            
            while (true)
            {
                LEDAnime();
                NewMail();
                leds.stopBlinkSeq(4);
                YAPI.Sleep(1000,ref errmsg);
            }
        }
    }
}
