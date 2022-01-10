using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Threading;
using System.Xml.Serialization;

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
			using (var stream = System.IO.File.OpenWrite(path))
				serializer.Serialize(stream, obj);
		}
	}

	internal class SharepointFolderCleanerWorker : IDisposable
	{
		private volatile bool disposed = false;


		private Thread thread;
		private readonly object mutex = new object();
		private Queue<string> queue;

		private void CleanFolder(ClientContext context, string path)
		{
			var folder = context.Web.GetFolderByServerRelativeUrl(path);

			// request files
			context.Load(folder, f => f.Files);
			if (!context.ExecuteQueryWithDelayAndAuthCheck())
				return;

			if (folder.Files.Count > 0)
			{
				// request versions for each file
				foreach (var file in folder.Files)
					context.Load(file, f => f.Versions);
				context.ExecuteQueryWithDelay();

				// delete the old versions
				var num_deletes = 0;
				foreach (var file in folder.Files)
				{
					if (file.Versions.Count > 0)
					{
						file.Versions.DeleteAll();
						num_deletes += file.Versions.Count;

						if (num_deletes >= 20)
						{
							context.ExecuteQueryWithDelay();
							num_deletes = 0;
						}
					}
				}
				if (num_deletes > 0)
					context.ExecuteQueryWithDelay();
			}

			Console.WriteLine($@"{path}: OK");
		}

		public SharepointFolderCleanerWorker(Uri site_uri, SharePointOnlineCredentials credentials)
		{
			queue = new Queue<string>();

			thread = new Thread(() =>
			{
				ClientContext context = null;

				try
				{
					while (!disposed)
					{
						string path = null;
						lock (mutex)
						{
							if (queue.Count == 0)
								continue;
							path = queue.Dequeue();
						}

						if (path != null)
						{
							if (context == null)
							{
								try
								{
									context = new ClientContext(site_uri);
									context.Credentials = credentials;
								}
								catch (Exception e)
								{
									Console.Error.WriteLine($@"[{e.GetType()}] {e.Message}");
								}
							}

							try
							{
								CleanFolder(context, path);
							}
							catch (Exception e)
							{
								Console.Error.WriteLine($@"{path}: [{e.GetType()}] {e.Message}");
							}

							Thread.Sleep(50);
						}
						else
							Thread.Sleep(250);
					}

				}
				finally
				{
					if (context != null)
					{
						context.Dispose();
						context = null;
					}
				}
			});
			thread.Start();
		}

		public void Enqueue(Folder folder)
		{
			lock (mutex)
			{
				queue.Enqueue(folder.ServerRelativeUrl);
			}
		}

		public void Wait()
		{
			while (!disposed)
			{
				lock (mutex)
				{
					if (queue.Count == 0)
						return;
				}

				Thread.Sleep(200);
			}
		}

		public void Dispose()
		{
			if (!disposed)
			{
				disposed = true;
				thread.Join();
				thread = null;
			}
			GC.SuppressFinalize(this);
		}
	}

	internal class SharepointFolderCleaner : IDisposable
	{
		private volatile bool disposed = false;
		private SharepointFolderCleanerWorker[] workers;
		private int next_worker = 0;

		private readonly string history_path;
		private readonly List<string> history;
		private int history_this_session = 0;

		private void EnqueueFolder(ClientContext context, Uri site_uri, SharePointOnlineCredentials credentials, Folder folder)
		{
			// ignore this folder and it's subfolders if it is already in the history
			if (history.Contains(folder.ServerRelativeUrl))
				return;

			try
			{
				// enqueue this folder
				var worker_index = (++next_worker) % workers.Length;
				if (workers[worker_index] == null)
					workers[worker_index] = new SharepointFolderCleanerWorker(site_uri, credentials);
				workers[worker_index].Enqueue(folder);

				// enqueue subfolders
				context.Load(folder, f => f.Folders);
				if (context.ExecuteQueryWithDelayAndAuthCheck())
				{
					foreach (var subfolder in folder.Folders)
						EnqueueFolder(context, site_uri, credentials, subfolder);
				}
			}
			finally
			{
				// record this folder in history
				history.Add(folder.ServerRelativeUrl);
				history_this_session++;
				if (history_this_session % 10 == 0)
					WriteSession();
			}

		}

		public SharepointFolderCleaner(Uri site_uri, SharePointOnlineCredentials credentials, int num_workers = 0)
		{
			history_path = site_uri.ToString().Trim().ToLower();
			foreach (var character in new char[] { ' ', '+', '-', ':', '/', '\\', '<', '>', '(', ')', '*', '.' })
				history_path = history_path.Replace(character, '_');
			history_path = $@"sharepoint_cleaner_{history_path}.xml";
			Console.WriteLine($@"Session path: {history_path}");

			if (System.IO.File.Exists(history_path))
			{
				try
				{
					history = history_path.XmlDeserialize<List<string>>();
				}
				catch (Exception e)
				{
					Console.Error.WriteLine($@"[{e.GetType()}] {e.Message}");
					history = new List<string>();
				}
			}
			else
				history = new List<string>();

			if (num_workers <= 0)
				num_workers = 8;
			workers = new SharepointFolderCleanerWorker[Math.Min(Math.Max(num_workers, 1), 64)];

			using (ClientContext context = new ClientContext(site_uri))
			{
				context.Credentials = credentials;
				context.Load(context.Web, w => w.Folders, w => w.CurrentUser, w => w.Lists);
				context.ExecuteQueryWithDelay();

				if (!context.Web.CurrentUser.IsSiteAdmin)
					throw new Exception("Access denied. User must be a site admin.");

				foreach (var folder in context.Web.Folders)
					EnqueueFolder(context, site_uri, credentials, folder);
			}
		}

		public void Wait()
		{
			foreach (var worker in workers)
			{
				if (worker != null)
					worker.Wait();
			}
		}

		private void WriteSession()
		{
			try
			{
				history.XmlSerialize(history_path);
				Console.WriteLine($@"Session written to {history_path}");
			}
			catch (Exception)
			{

			}
		}

		public void Dispose()
		{
			if (!disposed)
			{
				disposed = true;
				foreach (var worker in workers)
				{
					if (worker != null)
						worker.Dispose();
				}
				WriteSession();
			}
			GC.SuppressFinalize(this);
		}
	}

	internal class Program
	{
		private static SecureString ReadSecureString()
		{
			var pwd = new SecureString();
			while (true)
			{
				ConsoleKeyInfo i = Console.ReadKey(true);
				if (i.Key == ConsoleKey.Enter)
				{
					break;
				}
				else if (i.Key == ConsoleKey.Backspace)
				{
					if (pwd.Length > 0)
					{
						pwd.RemoveAt(pwd.Length - 1);
						Console.Write("\b \b");
					}
				}
				else if (i.KeyChar != '\u0000')
				{
					pwd.AppendChar(i.KeyChar);
					Console.Write("*");
				}
			}
			pwd.MakeReadOnly();
			return pwd;
		}

		private static Uri ReadURI()
		{
			var uri = Console.ReadLine().Trim();
			uri = uri.StartsWith("https://") ? uri.Substring("https://".Length) : uri;
			return new Uri($"https://{uri}");
		}

		private static void Run()
		{
			Console.Write("Sharepoint site URI: ");
			var site_uri = ReadURI();

			Console.Write("Username: ");
			var username = Console.ReadLine().Trim();

			Console.Write("Password: ");
			var password = ReadSecureString();
			Console.WriteLine();

			using (var cleaner = new SharepointFolderCleaner(site_uri, new SharePointOnlineCredentials(username, password)))
				cleaner.Wait();
		}

		static void Main(string[] args)
		{
			try
			{
				Run();
			}
			catch (Exception e)
			{
				Console.Error.WriteLine($@"[{e.GetType()}] {e.Message}");
			}
		}
	}
}
