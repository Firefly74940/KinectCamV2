using System;
using System.Diagnostics;
using Microsoft.Kinect;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

namespace KinectCam
{
	public static class KinectHelper
	{
		class KinectCamApplicationContext : ApplicationContext
		{
			private NotifyIcon TrayIcon;
			private ContextMenuStrip TrayIconContextMenu;
			private ToolStripMenuItem MirroredMenuItem;
			private ToolStripMenuItem DesktopMenuItem;
			private ToolStripMenuItem ZoomMenuItem;

			public KinectCamApplicationContext()
			{
				Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
				InitializeComponent();
				TrayIcon.Visible = true;
				TrayIcon.ShowBalloonTip(30000);
			}

			private void InitializeComponent()
			{
				TrayIcon = new NotifyIcon();

				TrayIcon.BalloonTipIcon = ToolTipIcon.Info;
				TrayIcon.BalloonTipText =
				  "For options use this tray icon.";
				TrayIcon.BalloonTipTitle = "KinectCamV2";
				TrayIcon.Text = "KinectCam";

				TrayIcon.Icon = IconExtractor.Extract(117, false);

				TrayIcon.DoubleClick += TrayIcon_DoubleClick;

				TrayIconContextMenu = new ContextMenuStrip();
				MirroredMenuItem = new ToolStripMenuItem();
				DesktopMenuItem = new ToolStripMenuItem();
				ZoomMenuItem = new ToolStripMenuItem();
				TrayIconContextMenu.SuspendLayout();

				// 
				// TrayIconContextMenu
				// 
				this.TrayIconContextMenu.Items.AddRange(new ToolStripItem[] {
				this.MirroredMenuItem,
				this.DesktopMenuItem,
				this.ZoomMenuItem
				});
				this.TrayIconContextMenu.Name = "TrayIconContextMenu";
				this.TrayIconContextMenu.Size = new Size(153, 70);
				// 
				// MirroredMenuItem
				// 
				this.MirroredMenuItem.Name = "Mirrored";
				this.MirroredMenuItem.Size = new Size(152, 22);
				this.MirroredMenuItem.Text = "Mirrored";
				this.MirroredMenuItem.Click += new EventHandler(this.MirroredMenuItem_Click);

				// 
				// DesktopMenuItem
				// 
				this.DesktopMenuItem.Name = "Desktop";
				this.DesktopMenuItem.Size = new Size(152, 22);
				this.DesktopMenuItem.Text = "Desktop";
				this.DesktopMenuItem.Click += new EventHandler(this.DesktopMenuItem_Click);

				// 
				// ZoomMenuItem
				// 
				this.ZoomMenuItem.Name = "Zoom";
				this.ZoomMenuItem.Size = new Size(152, 22);
				this.ZoomMenuItem.Text = "Zoom";
				this.ZoomMenuItem.Click += new EventHandler(this.ZoomMenuItem_Click);

				TrayIconContextMenu.ResumeLayout(false);
				TrayIcon.ContextMenuStrip = TrayIconContextMenu;
			}

			private void OnApplicationExit(object sender, EventArgs e)
			{
				TrayIcon.Visible = false;
			}

			private void TrayIcon_DoubleClick(object sender, EventArgs e)
			{
				TrayIcon.ShowBalloonTip(30000);
			}

			private void MirroredMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Mirrored = !KinectCamSettigns.Default.Mirrored;
			}

			private void DesktopMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Desktop = !KinectCamSettigns.Default.Desktop;
			}

			private void ZoomMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Zoom = !KinectCamSettigns.Default.Zoom;
			}

			public void Exit()
			{
				TrayIcon.Visible = false;
			}
		}

		static KinectCamApplicationContext context;
		static Thread contexThread;
		static Thread refreshThread;
		static KinectSensor Sensor;

		static void InitializeSensor()
		{
			var sensor = Sensor;
			if (sensor != null) return;

			try
			{
				sensor = KinectSensor.GetDefault();
				if (sensor == null) return;

				var reader = sensor.ColorFrameSource.OpenReader();
				reader.FrameArrived += reader_FrameArrived;

				sensor.Open();

				Sensor = sensor;

				if (context == null)
				{
					contexThread = new Thread(() =>
					{
						context = new KinectCamApplicationContext();
						Application.Run(context);
					});
					refreshThread = new Thread(() =>
					{
						while (true)
						{
							Thread.Sleep(250);
							Application.DoEvents();
						}
					});
					contexThread.IsBackground = true;
					refreshThread.IsBackground = true;
					contexThread.SetApartmentState(ApartmentState.STA);
					refreshThread.SetApartmentState(ApartmentState.STA);
					contexThread.Start();
					refreshThread.Start();
				}
			}
			catch
			{
				Trace.WriteLine("Error of enable the Kinect sensor!");
			}
		}

		public delegate void InvokeDelegate();

		static void reader_FrameArrived(object sender, ColorFrameArrivedEventArgs e)
		{
			using (var colorFrame = e.FrameReference.AcquireFrame())
			{
				if (colorFrame != null)
				{
					ColorFrameReady(colorFrame);
				}
			}
		}

		static unsafe void ColorFrameReady(ColorFrame frame)
		{
			if (frame.RawColorImageFormat == ColorImageFormat.Bgra)
			{
				frame.CopyRawFrameDataToArray(sensorColorFrameData);
			}
			else
			{
				frame.CopyConvertedFrameDataToArray(sensorColorFrameData, ColorImageFormat.Bgra);
			}
		}

		public static void DisposeSensor()
		{
			try
			{
				var sensor = Sensor;
				if (sensor != null && sensor.IsOpen)
				{
					sensor.Close();
					sensor = null;
					Sensor = null;
				}

				if (context != null)
				{
					context.Exit();
					context.Dispose();
					context = null;

					contexThread.Abort();
					refreshThread.Abort();
				}
			}
			catch
			{
				Trace.WriteLine("Error of disable the Kinect sensor!");
			}
		}

		public const int SensorWidth = 1920;
		public const int SensorHeight = 1080;
		public const int ZoomedWidth = 960;
		public const int ZoomedHeight = 540;
		public const int ZoomedWidthStart = ZoomedWidth/2;
		public const int ZoomedHeightStart =ZoomedHeight/2;
		public const int ZoomedWidthEnd = SensorWidth - ZoomedWidthStart;
		public const int ZoomedHeightEnd = SensorHeight - ZoomedHeightStart;

		public const int ZoomedPointerStart = ZoomedHeightStart * 1920 * 4 + ZoomedWidthStart * 4;
		public const int ZoomedPointerEnd = ZoomedHeightEnd * 1920 * 4+ ZoomedWidthEnd * 4;
		static readonly byte[] sensorColorFrameData = new byte[1920 * 1080 * 4];

		public unsafe static void GenerateFrame(IntPtr _ptr, int length, bool mirrored, bool zoom)
		{
			zoom = true;
			byte[] colorFrame = sensorColorFrameData;
			void* camData = _ptr.ToPointer();

			try
			{
				InitializeSensor();

				if (colorFrame != null)
				{
					int colorFramePointerStart = zoom ? ZoomedPointerStart : 0;
					int colorFramePointerEnd = zoom ? ZoomedPointerEnd - 1 : colorFrame.Length - 1;
					int width = zoom ? ZoomedWidth : SensorWidth;

					if (!mirrored)
					{
						fixed (byte* sDataB = &colorFrame[colorFramePointerStart])
						fixed (byte* sDataE = &colorFrame[colorFramePointerEnd])
						{
							byte* pData = (byte*)camData;
							byte* sData = (byte*)sDataE;
							bool redo = true;
						
							for (; sData > sDataB;)
							{
								for (var i = 0; i < width; ++i)
								{
									var p = sData - 3;
									*pData++ = *p++;
									*pData++ = *p++;
									*pData++ = *p++;
									if (zoom)
									{
										p = sData - 3;
										*pData++ = *p++;
										*pData++ = *p++;
										*pData++ = *p++;
									}
									sData -= 4;
								}
								if (zoom)
								{
									if (redo)
									{
										sData += width * 4;
									}
									else
									{
										sData -= ZoomedWidthStart * 2 * 4;
									}
									redo = !redo;
									
								}
							}

						}
					}
					else
					{
						fixed (byte* sDataB = &colorFrame[colorFramePointerStart])
						fixed (byte* sDataE = &colorFrame[colorFramePointerEnd])
						{
							byte* pData = (byte*)camData;
							byte* sData = (byte*)sDataE;

							var sDataBE = sData;
							var p = sData;
							var r = sData;
							bool redo = true;

							while (sData == (sDataBE = sData) &&
								   sDataB <= (sData -= (width * 4 - 1)))
							{
								
								r = sData;
								do
								{
									p = sData;
									*pData++ = *p++;
									*pData++ = *p++;
									*pData++ = *p++;
									if (zoom)
									{
										p = sData;
										*pData++ = *p++;
										*pData++ = *p++;
										*pData++ = *p++;
									}

								}
								while ((sData += 4) <= sDataBE);
								sData = r - 1;
								if (zoom)
								{
									if (redo)
									{
										sData += width * 4;
									}
									else
									{
										sData -= ZoomedWidthStart * 2 * 4;
									}
									redo = !redo;

								}
							}
						}
					}
				}
			}
			catch
			{
				byte* pData = (byte*)camData;
				for (int i = 0; i < length; ++i)
					*pData++ = 0;
			}
		}
	}
}
