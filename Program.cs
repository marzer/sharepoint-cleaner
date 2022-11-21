using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Xml.Serialization;

using SharePointFile = Microsoft.SharePoint.Client.File;
using SharePointFolder = Microsoft.SharePoint.Client.Folder;

namespace sharepoint_cleaner
{
	internal static class Extensions
	{
		public static void ExecuteQueryWithDelay(this ClientContext context)
		{
			context.ExecuteQuery();
			Thread.Sleep(20); // necessary so we don't smash SharePoint API limits and get throttled immediately
		}

		public static bool ExecuteQueryWithDelayAndAuthCheck(this ClientContext context)
		{
			try
			{
				context.ExecuteQueryWithDelay();
			}
			catch (ServerUnauthorizedAccessException)
			{
				return false;
			}
			return true;
		}


		public static T XmlDeserialize<T>(this string path) where T : class
		{
			XmlSerializer deserializer = new XmlSerializer(typeof(T));
			using (TextReader reader = new StreamReader(path))
				return deserializer.Deserialize(reader) as T;
		}

		public static void XmlSerialize<T>(this T obj, string path) where T : class
		{
			XmlSerializer serializer = new XmlSerializer(typeof(T));
			using (var stream = System.IO.File.Create(path))
				serializer.Serialize(stream, obj);
		}

		public static SecureString ToSecureString(this string str)
		{
			var pwd = new SecureString();
			foreach (char c in str)
				pwd.AppendChar(c);
			return pwd;
		}

		private static readonly string[] ToFileSizeString_suffixes = new string[]{
			"B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
		};

		public static string ToFileSizeString(this long bytes)
		{
			if (bytes == 0L)
				return "0 B";
			else
			{
				var dbytes = (double)bytes;
				var log = 0;
				while (dbytes >= 1024.0)
				{
					dbytes /= 1024.0;
					log++;
				}
				if (log >= ToFileSizeString_suffixes.Length)
					return $"{bytes} B";

				return $"{dbytes:0.0#} {ToFileSizeString_suffixes[log]}";

			}
		}
	}

	public class Session
	{
		public DateTime Started;
		public DateTime Continued;
		public HashSet<string> Processed { get; private set; } = new HashSet<string>();

		// "Run" vs "Session":
		// - runs are singular runs of the application, disregarding any session continuation
		// - sessions are the cumulative sessions that might be spread over multiple runs

		public long FilesThisSession;
		public long FoldersThisSession;
		public long VersionsThisSession;
		public long FreedSpaceThisSession;

		[XmlIgnore]
		public long FilesThisRun;
		[XmlIgnore]
		public long FoldersThisRun;
		[XmlIgnore]
		public long VersionsThisRun;
		[XmlIgnore]
		public long FreedSpaceThisRun;
	}

	internal class Program
	{
		private static readonly object console_mutex = new object();

		private static string ExceptionName(Exception e)
		{
			if (e as ServerException != null)
				return "ServerException";
			if (e as System.Net.WebException != null)
				return "WebException";
			return e.GetType().FullName;
		}

		private static void ColouredPrint(ConsoleColor col, TextWriter stream, string msg, params object[] args)
		{
			lock (console_mutex)
			{
				var prev = Console.ForegroundColor;
				Console.ForegroundColor = col;

				try
				{
					if (args != null && args.Length > 0)
						stream.Write(msg, args);
					else
						stream.Write(msg);
				}
				finally
				{
					Console.ForegroundColor = prev;
				}
			}
		}

		public static void Error(string msg, params object[] args)
		{
			ColouredPrint(ConsoleColor.Red, Console.Error, msg, args);
		}

		public static void Warning(string msg, params object[] args)
		{
			ColouredPrint(ConsoleColor.Yellow, Console.Error, msg, args);
		}

		public static void Info(string msg, params object[] args)
		{
			ColouredPrint(ConsoleColor.White, Console.Out, msg, args);
		}

		private class State
		{
			public Uri SiteURI;
			public SharePointOnlineCredentials Credentials;
			public string SessionPath;
			public Session Session;
			public ClientContext Context;
			public volatile bool Abort = false;

			private DateTime last_save = DateTime.UtcNow;

			public void LoadSession()
			{
				if (System.IO.File.Exists(SessionPath))
				{
					try
					{
						Session = SessionPath.XmlDeserialize<Session>();

						// reject sessions older than a week or continuations older than a day
						if ((DateTime.UtcNow - Session.Started) >= TimeSpan.FromDays(7)
							|| (DateTime.UtcNow - Session.Continued) >= TimeSpan.FromDays(1))
							Session = null;
					}
					catch (Exception e)
					{
						Error($"[{ExceptionName(e)}] {e.Message}\n");
					}
				}
				if (Session == null)
				{
					Session = new Session
					{
						Started = DateTime.UtcNow
					};
				}
				Session.Continued = DateTime.UtcNow;
			}

			public void SaveSession()
			{
				Session.Continued = DateTime.UtcNow;

				try
				{
					Session.XmlSerialize(SessionPath);
					Info($"Session written to {SessionPath}\n");
				}
				catch (Exception e)
				{
					Error($"[{ExceptionName(e)}] {e.Message}\n");
				}

			}

			private void IntermittentSaveSession()
			{
				var now = DateTime.UtcNow;
				if ((now - last_save) >= TimeSpan.FromSeconds(60))
				{
					SaveSession();
					last_save = now;
				}
			}

			private bool AlreadyProcessed(SharePointFile file)
			{
				return Session.Processed.Contains(file.ServerRelativeUrl);
			}

			private bool AlreadyProcessed(SharePointFolder folder)
			{
				return Session.Processed.Contains(folder.ServerRelativeUrl);
			}

			private void RecordAsProcessed(SharePointFile file)
			{
				if (Session.Processed.Add(file.ServerRelativeUrl))
				{
					Session.FilesThisRun++;
					Session.FilesThisSession++;
				}
			}

			private void RecordAsProcessed(SharePointFolder folder)
			{
				if (Session.Processed.Add(folder.ServerRelativeUrl))
				{
					Session.FoldersThisRun++;
					Session.FoldersThisSession++;
				}
			}
			private void RecordAsProcessed(FileVersionCollection versions)
			{
				Session.VersionsThisRun += versions.Count;
				Session.VersionsThisSession += versions.Count;
				foreach (var v in versions)
				{
					Session.FreedSpaceThisRun += v.Size;
					Session.FreedSpaceThisSession += v.Size;
				}
			}

			public void Clean(SharePointFolder folder)
			{
				if (AlreadyProcessed(folder) || Abort)
					return;

				Context.Load(folder, f => f.Folders, f => f.Files);
				if (!Context.ExecuteQueryWithDelayAndAuthCheck() || Abort)
					return; // silently skip folders requiring higher auth

				Info($"{folder.ServerRelativeUrl}\n");

				List<SharePointFile> batch = new List<SharePointFile>();
				long prev_versions = Session.VersionsThisRun;
				long prev_freed_space = Session.FreedSpaceThisRun;

				var FlushBatch = (bool force) =>
				{
					if (batch.Count == 0 || (batch.Count < 50 && !force))
						return;

					var original_batch = batch.ToList();

					try
					{
						// get versions for whole batch
						for (int i = batch.Count; i --> 0;)
						{
							if (Abort)
								break;

							try
							{
								Context.Load(batch[i], f => f.Versions);
							}
							catch (Exception e)
							{
								Warning($"[{ExceptionName(e)}] requesting versions for {batch[i].ServerRelativeUrl}: {e.Message}\n");
								batch.RemoveAt(i);
							}
						}

						// if all version requests failed to enqueue then there's no work to do
						if (batch.Count == 0 || Abort)
							return;

						// process the version requests
						try
						{
							Context.ExecuteQueryWithDelay();
						}
						catch (Exception e)
						{
							Warning($"[{ExceptionName(e)}] failed to request versions for batch: {e.Message}\n");
							Warning($"Attempting to re-request versions individually...\n");

							for (int i = batch.Count; i --> 0;)
							{
								if (Abort)
									break;

								try
								{
									Context.Load(batch[i], f => f.Versions);
									Context.ExecuteQueryWithDelay();
								}
								catch (Exception e2)
								{
									Warning($"[{ExceptionName(e2)}] requesting versions for {batch[i].ServerRelativeUrl}: {e.Message}\n");
									batch.RemoveAt(i);
								}
							}
						}

						// if all version requests failed then there's no work to do
						if (batch.Count == 0 || Abort)
							return;

						// request history deletion for whole queue
						for (int i = batch.Count; i --> 0;)
						{
							if (Abort)
								break;

							try
							{
								if (batch[i].Versions.Count == 0)
									continue;
								RecordAsProcessed(batch[i].Versions);
								batch[i].Versions.DeleteAll();
							}
							catch (CollectionNotInitializedException e)
							{
								batch.RemoveAt(i); // no history for this object
							}
							catch (Exception e)
							{
								Warning($"[{ExceptionName(e)}] requesting version history deletion for {batch[i].ServerRelativeUrl}: {e.Message}\n");
								batch.RemoveAt(i);
							}
						}

						// if all deletion requests failed to enqueue then there's no work to do
						if (batch.Count == 0 || Abort)
							return;

						// process the deletion requests
						try
						{
							Context.ExecuteQueryWithDelay();
						}
						catch (Exception e)
						{
							Warning($"[{ExceptionName(e)}] failed to delete version histories for batch: {e.Message}\n");
							Warning($"Attempting to delete version histories individually...\n");

							for (int i = batch.Count; i-- > 0;)
							{
								if (Abort)
									break;

								try
								{
									batch[i].Versions.DeleteAll();
									RecordAsProcessed(batch[i].Versions);
									Context.ExecuteQueryWithDelay();
								}
								catch (Exception e2)
								{
									Warning($"[{ExceptionName(e2)}] deleting version history for {batch[i].ServerRelativeUrl}: {e.Message}\n");
									batch.RemoveAt(i);
								}
							}
						}

						batch.Clear();
						long versions = Session.VersionsThisRun - prev_versions;
						long freed_space = Session.FreedSpaceThisRun - prev_freed_space;
						if (versions > 0)
						{
							Info($"Batch deleted {versions} past versions ({freed_space.ToFileSizeString()}).\n");
						}
					}
					finally
					{
						if (!Abort)
						{
							foreach (var file in original_batch)
								RecordAsProcessed(file);
							IntermittentSaveSession();
						}
					}
				};

				// handle files
				foreach (var file in folder.Files)
				{
					if (Abort)
						return;
					if (AlreadyProcessed(file))
						continue;

					batch.Add(file);
					FlushBatch(false);
				}
				FlushBatch(true);
				if (Abort)
					return;

				// handle subfolders
				var subfolders = folder.Folders.ToList();
				subfolders.Sort((a, b) => a.ServerRelativeUrl.CompareTo(b.ServerRelativeUrl));
				foreach (var subfolder in subfolders)
				{
					if (Abort)
						return;

					Clean(subfolder);
				}

				// handle subfolders and finish up
				if (!Abort)
				{
					RecordAsProcessed(folder);
					IntermittentSaveSession();
				}
			}
		}

		private static SecureString ReadSecureString()
		{
			var pwd = new SecureString();
			while (true)
			{
				ConsoleKeyInfo i = Console.ReadKey(true);
				if (i.Key == ConsoleKey.Enter)
				{
					pwd.MakeReadOnly();
					return pwd;
				}
				else if (i.Key == ConsoleKey.Backspace)
				{
					if (pwd.Length > 0)
					{
						pwd.RemoveAt(pwd.Length - 1);
						Info("\b \b");
					}
				}
				else if (i.KeyChar != '\u0000')
				{
					pwd.AppendChar(i.KeyChar);
					Info("*");
				}
			}
		}

		private static Uri ReadURI(string input = null)
		{
			var uri = input == null ? Console.ReadLine().Trim() : input;
			uri = uri.StartsWith("https://") ? uri.Substring("https://".Length) : uri;
			return new Uri($"https://{uri}");
		}

		private static void Run(string[] args)
		{
			Info("---------------------------------------------------------\n");
			Info("sharepoint-cleaner - github.com/marzer/sharepoint-cleaner\n");
			Info("---------------------------------------------------------\n");
			State state = new State();

			Info("Site URI: ");
			state.SiteURI = args.Length >= 1 ? ReadURI(args[0]) : ReadURI();
			if (args.Length >= 1)
				Info($"{state.SiteURI}\n");

			Info("Username: ");
			var username = args.Length >= 2 ? args[1] : Console.ReadLine().Trim();
			if (args.Length >= 2)
				Info($"{username}\n");

			if (args.Length >= 3)
				state.Credentials = new SharePointOnlineCredentials(username, args[2].ToSecureString());
			else
			{
				Info("Password: ");
				state.Credentials = new SharePointOnlineCredentials(username, ReadSecureString());
				Info("\n");
			}

			// initialized session
			state.SessionPath = state.SiteURI.ToString().Trim().ToLower();
			foreach (var character in new char[] { ' ', '+', '-', ':', '/', '\\', '<', '>', '(', ')', '*', '.', '?', '@' })
				state.SessionPath = state.SessionPath.Replace(character, '_');
			state.SessionPath = $@"{state.SessionPath}_{username}";
			state.SessionPath = ((uint)state.SessionPath.GetHashCode()).ToString();
			state.SessionPath = $@"sharepoint-cleaner_{state.SessionPath}.xml";
			Info($"Session: {state.SessionPath}\n");
			Info("---------------------------------------------------------\n");
			state.LoadSession();
			Console.CancelKeyPress += (sender, args) =>
			{
				args.Cancel = true;
				state.Abort = true;
				Info("Aborting...\n");
			};

			// do the thing
			using (var context = new ClientContext(state.SiteURI))
			{
				state.Context = context;
				context.Credentials = state.Credentials;
				context.Load(context.Web, w => w.Folders);
				if (!context.ExecuteQueryWithDelayAndAuthCheck())
					throw new Exception("Access denied.");

				var folders = context.Web.Folders.ToList();
				folders.Sort((a, b) => a.ServerRelativeUrl.CompareTo(b.ServerRelativeUrl));
				foreach (var folder in folders)
				{
					if (state.Abort)
						break;
					state.Clean(folder);
				}
			}

			// finish up
			state.SaveSession();
			Info("---------------------------------------------------------\n");
			Info($"  This run: processed {state.Session.FilesThisRun} files and {state.Session.FoldersThisRun} folders; deleted {state.Session.VersionsThisRun} past versions ({state.Session.FreedSpaceThisRun.ToFileSizeString()}).\n");
			Info($"Cumulative: processed {state.Session.FilesThisSession} files and {state.Session.FoldersThisSession} folders; deleted {state.Session.VersionsThisSession} past versions ({state.Session.FreedSpaceThisSession.ToFileSizeString()}).\n");
		}

		static int Main(string[] args)
		{
			try
			{
				Run(args);
				return 0;
			}
			catch (Exception e)
			{
				Error($"[{ExceptionName(e)}] {e.Message}\n");
				return 1;
			}
		}
	}
}
