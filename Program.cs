using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
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

		[XmlIgnore]
		public long FilesThisSession;
		[XmlIgnore]
		public long FoldersThisSession;
		[XmlIgnore]
		public long VersionsThisSession;
		[XmlIgnore]
		public long FreedSpaceThisSession;
	}

	internal class Program
	{
		private static readonly object console_mutex = new object();


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
						Error($"[{e.GetType()}] {e.Message}\n");
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
					Error($"[{e.GetType()}] {e.Message}\n");
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
				Session.Processed.Add(file.ServerRelativeUrl);
				Session.FilesThisSession++;
			}

			private void RecordAsProcessed(SharePointFolder folder)
			{
				Session.Processed.Add(folder.ServerRelativeUrl);
				Session.FoldersThisSession++;
			}
			private void RecordAsProcessed(FileVersionCollection versions)
			{
				Session.VersionsThisSession += versions.Count;
				foreach (var v in versions)
					Session.FreedSpaceThisSession += v.Size;
			}

			public void Clean(SharePointFolder folder)
			{
				if (AlreadyProcessed(folder) || Abort)
					return;

				Context.Load(folder, f => f.Folders, f => f.Files);
				if (!Context.ExecuteQueryWithDelayAndAuthCheck() || Abort)
					return; // silently skip folders requiring higher auth

				Info($"{folder.ServerRelativeUrl}\n");

				// handle files
				if (folder.Files.Count > 0)
				{
					// request versions for each file
					bool any_versions_requested = false;
					int pending_version_requests = 0;
					var FlushVersionRequests = (bool force) => {
						if (pending_version_requests > 0 && (pending_version_requests >= 10 || force))
						{
							try
							{
								Context.ExecuteQueryWithDelay();
							}
							catch (Exception e)
							{
								Warning($"[{e.GetType()}] {e.Message}\n");
							}
							pending_version_requests = 0;
						}
					};
					foreach (var file in folder.Files)
					{
						if (AlreadyProcessed(file))
							continue;

						try
						{
							Context.Load(file, f => f.Versions);
							pending_version_requests++;
							FlushVersionRequests(false);
							any_versions_requested = true;
						}
						catch (Exception e)
						{
							Warning($"[{e.GetType()}] {e.Message}\n");
						}

						if (Abort)
							return;
					}
					FlushVersionRequests(true);
					if (Abort)
						return;

					// delete the old versions
					if (any_versions_requested)
					{
						var pending_deletes = 0;
						var FlushDeleteRequests = (bool force) => {
							if (pending_deletes > 0 && (pending_deletes >= 50 || force))
							{
								try
								{
									Context.ExecuteQueryWithDelay();
								}
								catch (Exception e)
								{
									Warning($"[{e.GetType()}] {e.Message}\n");
								}
								pending_deletes = 0;
							}
						};

						foreach (var file in folder.Files)
						{
							if (Abort)
								break;

							if (AlreadyProcessed(file))
								continue;

							try
							{
								RecordAsProcessed(file);
								RecordAsProcessed(file.Versions);
								pending_deletes += file.Versions.Count;
							}
							catch (CollectionNotInitializedException e)
							{
								continue; // this would be because the file version request failed
							}
							catch (Exception e)
							{
								Warning($"[{e.GetType()}] {e.Message}\n");
								continue;
							}

							if (file.Versions.Count == 0)
								continue;

							Info($"    Deleting {file.Versions.Count} past versions of {file.ServerRelativeUrl}\n");

							file.Versions.DeleteAll();
							FlushDeleteRequests(false);
						}
						FlushDeleteRequests(true);
					}
				}

				// handle subfolders
				foreach (var subfolder in folder.Folders)
					Clean(subfolder);

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

				foreach (var folder in context.Web.Folders)
					state.Clean(folder);
			}

			// finish up
			state.SaveSession();
			Info("---------------------------------------------------------\n");
			Info($"Processed {state.Session.FilesThisSession} files and {state.Session.FoldersThisSession} folders, deleting {state.Session.VersionsThisSession} past versions and freeing {state.Session.FreedSpaceThisSession.ToFileSizeString()}.\n");
		}

		static void Main(string[] args)
		{
			try
			{
				Run(args);
			}
			catch (Exception e)
			{
				Error($"[{e.GetType()}] {e.Message}\n");
			}
		}
	}
}
