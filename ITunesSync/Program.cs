using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ITunesSync
{
    class Program
    {
        const int MAX_PATH = 260;

        static void Main(string[] args)
        {
            iTunesLib.iTunesApp app = null;
            Console.WriteLine("ITunes Synchronization");
            try
            {

                app = new iTunesLib.iTunesApp();
                Console.WriteLine("ITunes Version " + app.Version);

                iTunesLib.IITLibraryPlaylist libraryPlaylist = app.LibraryPlaylist;
                if (libraryPlaylist == null)
                {
                    Console.WriteLine("Error: No library playlist detected");
                    Environment.Exit(-1);
                    return;
                }

                if (!Directory.Exists(ITunesSync.Default.MusicDirectory))
                {
                    Console.WriteLine("Error: Music Storage Directory Not Found ({0})");
                    Console.WriteLine("Check configuration settings file");
                    Environment.Exit(-1);
                    return;
                }

                Console.WriteLine("Music Storage Directory: {0}", ITunesSync.Default.MusicDirectory);
                Console.WriteLine("Delete Unresolved Files: {0}", ITunesSync.Default.DeleteUnresolvedTracks);
                Console.WriteLine("No Updates: {0}", ITunesSync.Default.NoUpdates);
                Console.WriteLine("Proceed with sync? (y/N)");
                ConsoleKeyInfo response = Console.ReadKey();
                if (response.Key != ConsoleKey.Y)
                {
                    Console.WriteLine("Info: ITunes Synchronization Aborted");
                    return;
                }

                Console.WriteLine("{0} Playlist, {1} assets", libraryPlaylist.Name, libraryPlaylist.Tracks.Count);

                var fileTracks = from libTrack in libraryPlaylist.Tracks.OfType<iTunesLib.IITTrack>() where libTrack.Kind == iTunesLib.ITTrackKind.ITTrackKindFile select libTrack;
                var cdTracks = from track in fileTracks where track is iTunesLib.IITFileOrCDTrack select track as iTunesLib.IITFileOrCDTrack;
                var resolvedTracks = from resolvedTrack in cdTracks.AsParallel() where !String.IsNullOrEmpty(resolvedTrack.Location) select resolvedTrack.Location.ToLowerInvariant();
                var unresolvedTracks = from unresolvedTrack in cdTracks where String.IsNullOrEmpty(unresolvedTrack.Location) select unresolvedTrack as iTunesLib.IITTrack;
                var discoveredMusic = from file in Directory.GetFiles(ITunesSync.Default.MusicDirectory, "*.mp3", SearchOption.AllDirectories).AsParallel() select file.ToLowerInvariant();

                var newFiles = discoveredMusic.Except(resolvedTracks);

                bool noUpdates = ITunesSync.Default.NoUpdates;

                if (noUpdates)
                {
                    Console.WriteLine("**** NoUpdate Mode Enabled : NO CHANGES WILL BE WRITTEN TO DISK ****");
                }
                
                StringBuilder shortPath = new StringBuilder(MAX_PATH);
                foreach (String file in newFiles)
                {
                    if (!File.Exists(file))
                    {
                        continue;
                    }
                    if (!noUpdates)
                    {
                        if (file.Length > MAX_PATH)
                        {
                            GetShortPathName(file, shortPath, shortPath.Capacity);
                            libraryPlaylist.AddFile(shortPath.ToString());
                            shortPath.Remove(0, shortPath.Length);
                        }
                        else
                        {
                            libraryPlaylist.AddFile(file);
                        }
                    }
                    Console.WriteLine("Info: Adding {0}", file);
                }


                bool deleteUnresolved = ITunesSync.Default.DeleteUnresolvedTracks && !noUpdates;
                // print out all unknown files
                unresolvedTracks.ForEach(unresolvedTrack => Program.DeleteTrack(unresolvedTrack, deleteUnresolved));
            }
            catch
            {
                if (app != null)
                {
                    ((IDisposable)app).Dispose();
                }
            }
        }


        private static void DeleteTrack(iTunesLib.IITTrack track, bool delete)
        {
            Console.WriteLine("Info: Unresolved {0}", track.Name);
            if (delete)
            {
                track.Delete();
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetShortPathName(
            [MarshalAs(UnmanagedType.LPTStr)]
            string path,
            [MarshalAs(UnmanagedType.LPTStr)]
            StringBuilder shortPath,
            int shortPathLength
            );
    }
}
