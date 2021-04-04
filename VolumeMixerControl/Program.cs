using System;
using CSCore.CoreAudioAPI;
using System.Diagnostics;
using System.Linq;
using System.Globalization;
using System.Text;

namespace VolumeMixerControl
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            int argsCount = args.Length;
            if (argsCount == 3)
            {
                if (args[0] == "changeVolume")
                {
                    changeSoftwareVolume(args[1], float.Parse(args[2], CultureInfo.InvariantCulture.NumberFormat));

                }
            }
            else if (argsCount == 2)
            {
                if (args[0] == "getMusicSoftware")
                    Console.WriteLine(getMusicFromSoftware(args[1]));
            }
            else if (argsCount == 1)
            {
                if (args[0] == "getSoftwaresNames")
                    Console.WriteLine(getSoftwaresNames());
                else if (args[0] == "getSoftwaresIDs")
                    Console.WriteLine(getSoftwaresIDs());
            }

            else
            {
                Console.WriteLine("Volume Mixer Control");
                Console.WriteLine("Help:");
                Console.WriteLine("----------");
                Console.WriteLine("getSoftwaresNames : Get all softwares names who produce sound.");
                Console.WriteLine("----------");
                Console.WriteLine("changeVolume [SoftwareName] [Value] : Add or decrease percentage of volume of a software");
                Console.WriteLine("[SoftwareName] : Name of the taget Software");
                Console.WriteLine("[Value] : -100 to 100 %");
                Console.WriteLine("Exemple: changeVolume chrome 10 : Add 10% to Google Chrome volume");
                Console.WriteLine("----------");
                Console.WriteLine("getMusicSoftware [Name]: Get window name of the software");
                Console.WriteLine("Exemple: getMusicSoftware spotify");
            }



        }

        private static bool changeSoftwareVolume(string software, float volume)
        {
            if (software == "System")
                software = "Idle";
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    int[] IDsound = new int[sessionEnumerator.Count];
                    int SoftwareCount = 0;
                    foreach (var session in sessionEnumerator)
                    {
                        using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                        using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            IDsound[SoftwareCount] = sessionControl.ProcessID;
                            SoftwareCount++;
                        }
                    }

                    string[] SoftwareName = new string[SoftwareCount];
                    for (int a = 0; a < SoftwareCount; a++)
                    {
                        SoftwareName[a] = Process.GetProcessById(IDsound[a]).ProcessName;
                        if (SoftwareName[a] == software)
                        {
                            changeSoftwareIDVolume(IDsound[a], volume);
                            break;
                        }
                    }
                }
            }
            return true;
        }

        private static bool changeSoftwareIDVolume(int id, float volume)
        {
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    int[] IDsound = new int[sessionEnumerator.Count];
                    foreach (var session in sessionEnumerator)
                    {
                        using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                        using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                        {

                            if (sessionControl.ProcessID == id)
                            {
                                float currentVolume;
                                simpleVolume.GetMasterVolumeNative(out currentVolume);

                                volume = volume / 100;

                                float newVolume = currentVolume + volume;
                                if (newVolume < 0)
                                    newVolume = 0f;
                                if (newVolume > 1)
                                    newVolume = 1f;
                                simpleVolume.MasterVolume = newVolume;
                            }
                        }
                    }

                }
            }

            return true;
        }

        private static string getSoftwaresNames()
        {
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    int[] IDsound = new int[sessionEnumerator.Count];
                    int SoftwareCount = 0;
                    foreach (var session in sessionEnumerator)
                    {
                        using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                        using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            IDsound[SoftwareCount] = sessionControl.ProcessID;
                            SoftwareCount++;
                        }
                    }

                    string[] SoftwareName = new string[SoftwareCount];
                    for (int a = 0; a < SoftwareCount; a++)
                    {
                        SoftwareName[a] = Process.GetProcessById(IDsound[a]).ProcessName;
                    }

                    //Array.ForEach(SoftwareName, Console.WriteLine); //Debug
                    string str = "";
                    foreach (var item in SoftwareName)
                    {
                        if (item.ToString() == "Idle")
                            str += "System" + " ";
                        else
                            str += item.ToString() + " ";
                    }
                    str = str.Remove(str.Length - 1);
                    return str;

                }
            }
        }

        private static string getMusicFromSoftware(String software)
        {
            var proc = Process.GetProcessesByName(software).FirstOrDefault(p => !string.IsNullOrWhiteSpace(p.MainWindowTitle));
            if (proc == null)
            {
                return "Error: Software Not Found";
            }

            return proc.MainWindowTitle;


        }
        private static string getSoftwaresIDs()
        {
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    int[] IDsound = new int[sessionEnumerator.Count];
                    int SoftwareCount = 0;
                    foreach (var session in sessionEnumerator)
                    {
                        using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                        using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            IDsound[SoftwareCount] = sessionControl.ProcessID;
                            SoftwareCount++;
                        }
                    }
                    //Array.ForEach(SoftwareName, Console.WriteLine); //Debug
                    string str = "";
                    foreach (var item in IDsound)
                    {
                        str += item.ToString() + ",";
                    }
                    str = str.Remove(str.Length - 1);
                    return str;

                }
            }
        }
        private static AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    //Console.WriteLine("DefaultDevice: " + device.FriendlyName);
                    var sessionManager = AudioSessionManager2.FromMMDevice(device);
                    return sessionManager;
                }
            }
        }
    }
}
