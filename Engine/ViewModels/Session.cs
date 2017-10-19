using System;
using System.Collections.Generic;
using System.Linq;
using Engine.EventArgs;
using Engine.Models;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace Engine.ViewModels
{
    public class Session : BaseNotificationClass
    {
        public event EventHandler<MessageEventArgs> OnMessageRaised;

        public List<DWG> _dwg = new List<DWG>();
        public List<AVersion> _versions = new List<AVersion>();
        
        private Process acadProc = new Process();

        public string _stringToReturn = "";

        public List<AVersion> Versions
        {
            get { return _versions; }
            private set
            {
                _versions = value;
                OnPropertyChanged(nameof(Versions));
            }
        }

        public List<DWG> CurrentDWGList
        {
            get { return _dwg; }
            set
            {
                _dwg = value;
                OnPropertyChanged(nameof(CurrentDWGList));
            }
        }

        public AVersion CurrentVersion { get; set; }

        public void RefreshVersions()
        {
            _versions.Add(new AVersion(1, "AutoCAD R13", "\"C:\\r13\\win\\acad.exe\"", "C:\\r13\\win\\acad.exe", true));
            _versions.Add(new AVersion(2, "AutoCAD 14", "\"C:\\Program Files\\Autodesk\\AutoCAD 2014\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2014\\acad.exe", true));
            _versions.Add(new AVersion(3, "AutoCAD 17 R1", "\"C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe", true));
            _versions.Add(new AVersion(4, "AutoCAD 17", "\"C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe", false));
        }

        public void RefreshList(List<string> list)
        {
            foreach (string file in list)
            {
                string dwg_name = file.Split('\\').ToList().Last();
                string name = dwg_name.Split('.').ToList().First();
                string job_num = name.Split('-').ToList().First();
                string page_num = name.Split('-').ToList().Last();
                CurrentDWGList.Add(new DWG(job_num,page_num,name,file));
            }
        }

        private void RaiseMessage(string message)
        {
            OnMessageRaised?.Invoke(this, new MessageEventArgs(message));
        }

        public void GoButton()
        {
            if (CurrentDWGList.Count() > 0) //check whether the user selected dwgs or not
            {
                if (CurrentVersion == null) // check if user selected Autocad version
                {
                    RaiseMessage("Please Select an AutoCAD Version");
                    return;
                }
                if (!File.Exists(CurrentVersion.CleanPathToVersion)) // check if selected version exists
                {
                    RaiseMessage("The AutoCAD version selected wasn't found in this machine");
                    return;
                }
                string directoryName = Path.GetDirectoryName(CurrentDWGList.First().PathToFile);
                string path = "C:\\r13\\Script.scr";

                if (!Directory.Exists("C:\\r13\\Script.scr")) // check if r13 folder exists
                {
                    RaiseMessage("The Script would be generated in C:\\r13\\, looks like this folder doesn't exist in your computer." +
                        " In order to use this tool please create it.");
                    return;
                }
                if (!File.Exists(directoryName + "\\TITLE.dwg"))
                {
                    RaiseMessage("The TITLE block wasnt found. Please make sure the folder in which " +
                        "your drawings are located include a 'TITLE.dwg' file.");
                    return;
                }

                if (!File.Exists(path))//create script file
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(path))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter(path))
                {
                    int count = 0;
                    foreach (DWG item in CurrentDWGList)
                    {
                        sw.WriteLine("FILEOPEN");
                        if (count < 1)
                        {
                            sw.WriteLine("Y");
                        }
                        sw.WriteLine(item.PathToFile);
                        sw.WriteLine("(load " + '"' + "DDTITLE" + '"' + ")(TITLE 'INS)");
                        if (CurrentVersion.RequiresJobNum)
                        {
                            sw.WriteLine(item.PageID);
                        }
                        sw.WriteLine("_qsave");
                        sw.WriteLine("");
                        sw.WriteLine("");
                        count += 1;
                    }
                    sw.WriteLine("_quit");
                }
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = CurrentVersion.PathToVersion + "/b \"C:\\r13\\Script.scr\"";
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                acadProc.EnableRaisingEvents = true;
                acadProc.StartInfo = startInfo;
                acadProc.Start();
                acadProc.WaitForExit();

                RaiseMessage("Done!!");
            }
            else
            {
                RaiseMessage("First select the dwgs to update");
            }
        }

        public string SelectedDWGS()
        {
            foreach(DWG dwg in CurrentDWGList)
            {
                _stringToReturn += dwg.FileName + ";";
            }
            return _stringToReturn;
        }
    }
}
