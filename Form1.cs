using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace WasapiLoopback
{
    public partial class Form1 : Form
    {
        private string strOut = "out.wav";
        private string strSource = "source.wav";
        public Form1()
        {
            InitializeComponent();
            if (File.Exists(strSource))
                File.Delete(strSource);
            if (File.Exists(strOut))
                File.Delete(strOut);
        }

        private void btnRecord_Click(object sender, EventArgs e)
        {
            captureDevice = CreateWaveInDevice();

            writer = new WaveFileWriter(strOut, captureDevice.WaveFormat);
            captureDevice.StartRecording();
        }

        private IWaveIn CreateWaveInDevice()
        {
            Console.WriteLine($"-------------------{Thread.CurrentThread.ManagedThreadId}");
            var deviceEnum = new MMDeviceEnumerator();
            var renderDevices = deviceEnum.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList();
            renderDevices.ForEach(d=>Console.WriteLine(d.FriendlyName));

            var device = (MMDevice)renderDevices[1];
            IWaveIn newWaveIn = new WasapiCapture(device);
            // Both WASAPI and WaveIn support Sample Rate conversion!
            var sampleRate = (int)16000;
            var channels = 0 + 1;
            newWaveIn.WaveFormat = new WaveFormat(sampleRate, channels);

            newWaveIn.DataAvailable += OnDataAvailable;
            newWaveIn.RecordingStopped += OnRecordingStopped;
            return newWaveIn;
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            Console.WriteLine("OnRecordingStopped");
            writer?.Dispose();
            writer = null;
        }

        private WaveFileWriter writer;
        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            //Console.WriteLine($"OnDataAvailable: {e.BytesRecorded} {Thread.CurrentThread.ManagedThreadId}");

#if true
            //Append binary data to file
            //Test to write callback data directly into file
            using (var fileStream = new FileStream(strSource, FileMode.Append,
                       FileAccess.Write, FileShare.None))
            using (var bw = new BinaryWriter(fileStream))
            {
                bw.Write(e.Buffer, 0, e.BytesRecorded);
            }
#endif


            writer.Write(e.Buffer, 0, e.BytesRecorded);
        }

        private IWaveIn captureDevice;
        private void btnFinish_Click(object sender, EventArgs e)
        {
            captureDevice?.StopRecording();
        }
    }
}
